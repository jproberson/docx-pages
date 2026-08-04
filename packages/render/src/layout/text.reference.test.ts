import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

import {
  layOutDocument,
  lookupFontMetrics,
  type LaidOutDocument,
  type ParagraphBox,
  type PlacedLine,
} from "@onepager/core";

import { readTextPlacements, type TextPlacement } from "../pdf/text.js";
import { readReferenceDocument } from "../testing/documents.js";
import { referenceCases, suppliedFaces, type ReferenceCase } from "../testing/cases.js";

const CASES = referenceCases().filter(
  (each) => each.renderedPath !== null && each.textLinesMatched !== null,
);

const textOf = (placed: PlacedLine): string =>
  placed.line.segments.map((segment) => (segment.kind === "text" ? segment.text : "")).join("");

// A symbol face's characters reach the page through the private use page Word
// writes them in, and come back out of the pdf in the low byte they shadow.
const outOfSymbolPage = (text: string): string =>
  Array.from(text, (character) => {
    const codePoint = character.codePointAt(0) ?? 0;
    return codePoint >= 0xf020 && codePoint <= 0xf0ff
      ? String.fromCodePoint(codePoint - 0xf000)
      : character;
  }).join("");

const normalise = (text: string): string => outOfSymbolPage(text.replace(/\s+/g, " ").trim());

function boxesOf(layout: LaidOutDocument): readonly ParagraphBox[] {
  const inBoxes = [...layout.headerFloats, ...layout.bodyFloats].flatMap((float) =>
    float.content.kind === "text-box" && float.content.text !== null
      ? [...float.content.text.boxes]
      : [],
  );
  return [...layout.header, ...layout.body, ...inBoxes];
}

const linesOf = (layout: LaidOutDocument): readonly PlacedLine[] =>
  boxesOf(layout).flatMap((box) => box.lines);

type Comparison = {
  readonly matched: number;
  readonly placed: number;
  readonly numbersMatched: number;
  readonly numbersPlaced: number;
};

async function laidOut(each: ReferenceCase): Promise<{
  readonly layout: LaidOutDocument;
  readonly drawn: readonly TextPlacement[];
}> {
  const layout = layOutDocument(readReferenceDocument(each), (request) =>
    lookupFontMetrics(request, suppliedFaces()),
  );
  if (layout.kind !== "laid-out") throw new Error(`blocked: ${layout.blocker.kind}`);

  const drawn = (
    await readTextPlacements(new Uint8Array(readFileSync(each.renderedPath ?? "")))
  ).filter((item) => normalise(item.text) !== "");

  return { layout, drawn };
}

// Word draws each line as its own run of text, so its own output says both which
// lines it broke the paragraph into and where each one sat. A list number is drawn
// as a run of its own too, but every bullet in a document reads the same, so a
// number is identified by the baseline of the line it belongs to rather than by
// its text alone.
async function compare(each: ReferenceCase): Promise<Comparison> {
  const { layout, drawn } = await laidOut(each);

  const taken = new Set<TextPlacement>();
  let matched = 0;
  let placed = 0;

  for (const line of linesOf(layout)) {
    const text = normalise(textOf(line));
    if (text === "") continue;

    const found = drawn.find((item) => !taken.has(item) && normalise(item.text) === text);
    if (found === undefined) continue;
    taken.add(found);
    matched += 1;

    const off = Math.max(
      Math.abs(found.leftPt - line.leftPt),
      Math.abs(found.baselinePt - line.baselinePt),
    );
    if (off <= each.textTolerancePt) placed += 1;
  }

  const numbers = compareNumbers(each, boxesOf(layout), drawn);
  return { matched, placed, ...numbers };
}

// How near a number's line has to sit to Word's own before the number drawn there
// can be taken for the same one.
const SAME_LINE_PT = 3;

function compareNumbers(
  each: ReferenceCase,
  boxes: readonly ParagraphBox[],
  drawn: readonly TextPlacement[],
): { readonly numbersMatched: number; readonly numbersPlaced: number } {
  let numbersMatched = 0;
  let numbersPlaced = 0;

  for (const box of boxes) {
    const marker = box.marker;
    const line = box.lines[0];
    if (marker === null || line === undefined) continue;

    const text = normalise(marker.text);
    if (text === "") continue;

    const near = drawn.filter(
      (item) =>
        normalise(item.text) === text &&
        Math.abs(item.baselinePt - line.baselinePt) <= SAME_LINE_PT,
    );
    if (near.length === 0) continue;
    numbersMatched += 1;

    const off = Math.min(...near.map((item) => Math.abs(item.leftPt - marker.leftPt)));
    if (off <= each.textTolerancePt) numbersPlaced += 1;
  }

  return { numbersMatched, numbersPlaced };
}

describe.skipIf(CASES.length === 0)("text lines against Word", () => {
  for (const each of CASES) {
    describe(each.id, () => {
      it("breaks paragraphs into the lines Word broke them into", async () => {
        const { matched } = await compare(each);
        expect(matched).toBeGreaterThanOrEqual(each.textLinesMatched ?? 0);
      });

      it("puts those lines where Word put them", async () => {
        const { placed } = await compare(each);
        expect(placed).toBeGreaterThanOrEqual(each.textLinesPlaced ?? 0);
      });

      it("puts a list's number where Word put it", async () => {
        const { numbersMatched, numbersPlaced } = await compare(each);
        expect(numbersMatched).toBeGreaterThanOrEqual(each.numbersMatched ?? 0);
        expect(numbersPlaced).toBeGreaterThanOrEqual(each.numbersPlaced ?? 0);
      });
    });
  }
});
