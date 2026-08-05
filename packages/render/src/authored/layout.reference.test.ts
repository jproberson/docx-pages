import { describe, expect, it } from "vitest";

import {
  blockParagraphs,
  layOutDocument,
  lookupFontMetrics,
  openDocx,
  readBlocks,
  readStyleTable,
  type FaceRequest,
  type LaidOutDocument,
  type MetricsLookup,
  type MetricsResolver,
  type ParagraphBox,
} from "@docx-pages/core";

import { answeringParagraphs } from "./answers.js";
import { characterPlacements } from "./characters.js";
import { authoredDocuments } from "./documents.js";
import { authoredFace } from "./faces.js";
import { LEFT_PT } from "./package.js";
import {
  readMeasured,
  PARAGRAPH_TOLERANCE_PT,
  SHAPE_TOLERANCE_PT,
  type MeasuredDocument,
} from "./measured.js";

// Documents written for this suite, and Word's own answers about them. Unlike the
// seven reference documents these are committed, so a machine with Word's Calibri
// on it can check every rule here without a manifest.
const FACE = authoredFace();
const MEASURED = readMeasured();

const CASES = authoredDocuments().flatMap((each) => {
  const measured: MeasuredDocument | undefined = MEASURED.documents[each.id];
  return measured === undefined ? [] : [{ ...each, measured }];
});

const metricsFor: MetricsResolver = (request: FaceRequest): MetricsLookup =>
  lookupFontMetrics(request, FACE === null ? [] : [FACE]);

const layoutOf = (bytes: Uint8Array): LaidOutDocument => {
  const laid = layOutDocument(openDocx(bytes), metricsFor);
  if (laid.kind !== "laid-out") throw new Error(`blocked: ${JSON.stringify(laid.blocker)}`);
  return laid;
};

// Where every character of every paragraph sits along its line, which is what says
// where a tab landed. Word numbers paragraphs as this project does wherever there
// is no table in the way, and a character inside a table is not worth asking about
// anyway, since Word answers there for the row rather than for the character.
function charactersOf(bytes: Uint8Array): ReadonlyMap<string, number> {
  const pkg = openDocx(bytes);
  const styles = readStyleTable(pkg);
  const boxes = placedBoxes(layoutOf(bytes));

  const found = new Map<string, number>();
  for (const paragraph of blockParagraphs(readBlocks(pkg))) {
    const box = boxes.get(paragraph.index);
    if (box === undefined) continue;
    for (const placed of characterPlacements(paragraph, box, styles, metricsFor, LEFT_PT)) {
      found.set(`${String(paragraph.index + 1)}:${String(placed.index)}`, placed.leftPt);
    }
  }
  return found;
}

function placedBoxes(layout: LaidOutDocument): ReadonlyMap<number, ParagraphBox> {
  const found = new Map<number, ParagraphBox>();
  for (const page of layout.pages) {
    for (const box of page.body) if (!found.has(box.index)) found.set(box.index, box);
  }
  return found;
}

// Word answers for the text of a paragraph rather than for the paragraph's own
// top: one whose line fell past an object reports where the line landed, and one
// whose line rule opened room above its text reports the text under that room. An
// empty paragraph draws no line, and answers from wherever its mark came to rest.
// A paragraph inside a table answers for its row instead, which `answers.ts`
// lines up.
type Placed = { readonly page: number; readonly topPt: number; readonly leftPt: number };

function placedParagraphs(layout: LaidOutDocument): ReadonlyMap<number, Placed> {
  const found = new Map<number, Placed>();

  for (const page of layout.pages) {
    for (const box of page.body) {
      if (found.has(box.index)) continue;
      found.set(box.index, placedAt(box, page.index));
    }
  }
  return found;
}

function placedAt(box: ParagraphBox, pageIndex: number): Placed {
  const line = box.lines[0];
  return {
    page: pageIndex + 1,
    topPt: line === undefined ? box.markTopPt : line.topPt + line.seatPt,
    leftPt: (line?.leftPt ?? LEFT_PT) - LEFT_PT,
  };
}

// Every shape the document holds, by the name it was authored under, so a failure
// says which box is out rather than which index.
function fittedShapes(
  layout: LaidOutDocument,
): ReadonlyMap<string, Placed & { readonly widthPt: number; readonly heightPt: number }> {
  const found = new Map<string, Placed & { readonly widthPt: number; readonly heightPt: number }>();
  for (const page of layout.pages) {
    for (const float of page.floats) {
      if (found.has(float.anchor.name)) continue;
      found.set(float.anchor.name, {
        page: page.index + 1,
        topPt: float.topPt,
        leftPt: float.leftPt,
        widthPt: float.widthPt,
        heightPt: float.heightPt,
      });
    }
  }
  return found;
}

// A document agrees with Word about so many of the places it reported. Whatever it
// does not agree about is named in full when the count itself is wrong, so a
// failure reads as the rule that is out rather than as a number that moved.
function agreeing(
  agreed: number,
  total: number,
  wanted: number | undefined,
  off: readonly string[],
): void {
  if (agreed !== (wanted ?? total)) expect(off).toStrictEqual([]);
  expect(`${String(agreed)} of ${String(total)}`).toBe(
    `${String(wanted ?? total)} of ${String(total)}`,
  );
}

describe.skipIf(CASES.length === 0 || FACE === null)("authored documents against Word", () => {
  for (const each of CASES) {
    describe(`${each.id}: ${each.asks}`, () => {
      it("puts every paragraph where Word put it, on the page Word put it on", () => {
        const placed = placedParagraphs(layoutOf(each.bytes));
        const answers = answeringParagraphs(readBlocks(openDocx(each.bytes)));
        const off: string[] = [];
        let agreed = 0;

        for (const expected of each.measured.paragraphs) {
          const answer = answers[expected.index - 1];
          const ours = placed.get(answer?.paragraph ?? -1);
          if (answer === undefined || ours === undefined) {
            off.push(`paragraph ${String(expected.index)} was not laid out`);
            continue;
          }
          if (
            ours.page === expected.page &&
            Math.abs(ours.topPt - expected.topPt) <= PARAGRAPH_TOLERANCE_PT &&
            (!answer.comparesLeft ||
              Math.abs(ours.leftPt - expected.leftPt) <= PARAGRAPH_TOLERANCE_PT)
          ) {
            agreed += 1;
            continue;
          }
          off.push(
            `paragraph ${String(expected.index)} at ${ours.topPt.toFixed(2)},${ours.leftPt.toFixed(2)} on page ${String(ours.page)}; Word says ${String(expected.topPt)},${String(expected.leftPt)} on page ${String(expected.page)}`,
          );
        }

        agreeing(agreed, each.measured.paragraphs.length, each.paragraphsPlaced, off);
      });

      it.runIf(each.measured.characters.length > 0)(
        "lands every character along its line where Word landed it",
        () => {
          const ours = charactersOf(each.bytes);
          const off: string[] = [];
          let placed = 0;

          for (const expected of each.measured.characters) {
            const key = `${String(expected.paragraph)}:${String(expected.index)}`;
            const mine = ours.get(key);
            if (mine === undefined) {
              off.push(`character ${key} was not laid out`);
              continue;
            }
            if (Math.abs(mine - expected.leftPt) <= PARAGRAPH_TOLERANCE_PT) {
              placed += 1;
              continue;
            }
            off.push(
              `character ${key} at ${mine.toFixed(2)}, Word says ${String(expected.leftPt)}`,
            );
          }

          agreeing(placed, each.measured.characters.length, each.charactersPlaced, off);
        },
      );

      it.runIf(each.measured.shapes.length > 0)("sizes every shape the way Word sized it", () => {
        const shapes = fittedShapes(layoutOf(each.bytes));
        const off: string[] = [];

        for (const expected of each.measured.shapes) {
          const ours = shapes.get(expected.name);
          if (ours === undefined) {
            off.push(`${expected.name} was not laid out`);
            continue;
          }
          if (Math.abs(ours.widthPt - expected.widthPt) > SHAPE_TOLERANCE_PT) {
            off.push(
              `${expected.name} is ${ours.widthPt.toFixed(3)} wide, Word says ${String(expected.widthPt)}`,
            );
          }
          if (Math.abs(ours.heightPt - expected.heightPt) > SHAPE_TOLERANCE_PT) {
            off.push(
              `${expected.name} is ${ours.heightPt.toFixed(3)} tall, Word says ${String(expected.heightPt)}`,
            );
          }
        }

        expect(off).toStrictEqual([]);
      });
    });
  }
});
