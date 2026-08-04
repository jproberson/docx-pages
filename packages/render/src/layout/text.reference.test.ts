import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

import {
  layOutDocument,
  lookupFontMetrics,
  type LaidOutDocument,
  type LaidOutPage,
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

// The header and the footer are drawn again on every page, so every page holds
// their boxes as well as its own.
function boxesOnPage(layout: LaidOutDocument, page: LaidOutPage): readonly ParagraphBox[] {
  const floats = [...layout.headerFloats, ...layout.footerFloats, ...page.floats];
  const inBoxes = floats.flatMap((float) =>
    float.content.kind === "text-box" && float.content.text !== null
      ? [...float.content.text.boxes]
      : [],
  );
  return [...layout.header, ...layout.footer, ...page.body, ...inBoxes];
}

type Comparison = {
  readonly matched: number;
  readonly placed: number;
  readonly runsMatched: number;
  readonly runsPlaced: number;
  readonly numbersMatched: number;
  readonly numbersPlaced: number;
};

// A stretch of text drawn in one face at one place, which is what Word writes an
// item for and what a line is made of on our side.
type Run = {
  readonly text: string;
  readonly leftPt: number;
};

// Word writes an item per run, so a line whose runs differ in face or size
// arrives as several items in a row. The line is the shortest stretch of items
// from some starting point that spells it out.
function itemsFor(
  text: string,
  drawn: readonly TextPlacement[],
  taken: ReadonlySet<TextPlacement>,
): readonly TextPlacement[] | null {
  for (const [start] of drawn.entries()) {
    let joined = "";
    for (const [end, item] of drawn.slice(start).entries()) {
      if (taken.has(item)) break;
      joined += item.text;
      const spelled = normalise(joined);
      if (spelled === text) return drawn.slice(start, start + end + 1);
      if (!text.startsWith(spelled)) break;
    }
  }
  return null;
}

// Runs line up by the characters that carry ink, since the spaces around them are
// drawn by whichever run happens to hold them. A run opening on a space has no
// inked character of its own to line up at, so it is left out.
const inkOf = (text: string): string => text.replace(/\s+/g, "");

function startsOf(runs: readonly Run[]): ReadonlyMap<number, number> {
  const found = new Map<number, number>();
  let at = 0;
  for (const run of runs) {
    if (inkOf(run.text) !== "" && !/^\s/.test(run.text) && !found.has(at)) {
      found.set(at, run.leftPt);
    }
    at += inkOf(run.text).length;
  }
  return found;
}

const runsOf = (placed: PlacedLine): readonly Run[] =>
  placed.line.segments.flatMap((segment) =>
    segment.kind === "text"
      ? [{ text: outOfSymbolPage(segment.text), leftPt: placed.leftPt + segment.offsetPt }]
      : [],
  );

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

// Word's own output says both which lines it broke each paragraph into and where
// every run of every line sat. Breaking is asked of the whole document and
// placement of the page: a line counts as placed only when the run of items Word
// drew with that text on that same page starts where ours does.
async function compare(each: ReferenceCase): Promise<Comparison> {
  const { layout, drawn } = await laidOut(each);

  const taken = new Set<TextPlacement>();
  const elsewhere: string[] = [];
  let matched = 0;
  let placed = 0;
  let runsMatched = 0;
  let runsPlaced = 0;
  let numbersMatched = 0;
  let numbersPlaced = 0;

  for (const page of layout.pages) {
    const onPage = drawn.filter((item) => item.pageIndex === page.index);
    const boxes = boxesOnPage(layout, page);

    for (const line of boxes.flatMap((box) => box.lines)) {
      const text = normalise(textOf(line));
      if (text === "") continue;

      const items = itemsFor(text, onPage, taken);
      const found = items?.[0];
      if (items === null || found === undefined) {
        elsewhere.push(text);
        continue;
      }
      for (const item of items) taken.add(item);
      matched += 1;

      const off = Math.max(
        Math.abs(found.leftPt - line.leftPt),
        Math.abs(found.baselinePt - line.baselinePt),
      );
      if (off <= each.textTolerancePt) placed += 1;

      // Only where the run starts is asked of it: a superscript sits off the
      // line's own baseline, which the line itself is already pinned against.
      const ours = startsOf(runsOf(line));
      for (const [at, leftPt] of startsOf(items)) {
        const mine = ours.get(at);
        if (mine === undefined) continue;
        runsMatched += 1;
        if (Math.abs(mine - leftPt) <= each.textTolerancePt) runsPlaced += 1;
      }
    }

    const numbers = compareNumbers(each, boxes, onPage);
    numbersMatched += numbers.numbersMatched;
    numbersPlaced += numbers.numbersPlaced;
  }

  // A line Word drew on another page was still broken the way Word broke it.
  for (const text of elsewhere) {
    const items = itemsFor(text, drawn, taken);
    if (items === null) continue;
    for (const item of items) taken.add(item);
    matched += 1;
  }

  return { matched, placed, runsMatched, runsPlaced, numbersMatched, numbersPlaced };
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

      it("puts those lines where Word put them, on the page Word put them on", async () => {
        const { placed } = await compare(each);
        expect(placed).toBeGreaterThanOrEqual(each.textLinesPlaced ?? 0);
      });

      it("starts each of a line's runs where Word started it", async () => {
        const { runsMatched, runsPlaced } = await compare(each);
        expect(runsMatched).toBeGreaterThanOrEqual(each.textRunsMatched ?? 0);
        expect(runsPlaced).toBeGreaterThanOrEqual(each.textRunsPlaced ?? 0);
      });

      it("puts a list's number where Word put it", async () => {
        const { numbersMatched, numbersPlaced } = await compare(each);
        expect(numbersMatched).toBeGreaterThanOrEqual(each.numbersMatched ?? 0);
        expect(numbersPlaced).toBeGreaterThanOrEqual(each.numbersPlaced ?? 0);
      });
    });
  }
});
