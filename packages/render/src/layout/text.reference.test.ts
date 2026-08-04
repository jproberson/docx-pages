import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

import {
  layOutDocument,
  lookupFontMetrics,
  type LaidOutDocument,
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

const normalise = (text: string): string => text.replace(/\s+/g, " ").trim();

function linesOf(layout: LaidOutDocument): readonly PlacedLine[] {
  const inBoxes = [...layout.headerFloats, ...layout.bodyFloats].flatMap((float) =>
    float.content.kind === "text-box" && float.content.text !== null
      ? [...float.content.text.boxes]
      : [],
  );
  return [...layout.header, ...layout.body, ...inBoxes].flatMap((box) => box.lines);
}

type Comparison = {
  readonly matched: number;
  readonly placed: number;
};

// Word draws each line as its own run of text, so its own output says both which
// lines it broke the paragraph into and where each one sat.
async function compare(each: ReferenceCase): Promise<Comparison> {
  const layout = layOutDocument(readReferenceDocument(each), (request) =>
    lookupFontMetrics(request, suppliedFaces()),
  );
  if (layout.kind !== "laid-out") throw new Error(`blocked: ${layout.blocker.kind}`);

  const drawn = (
    await readTextPlacements(new Uint8Array(readFileSync(each.renderedPath ?? "")))
  ).filter((item) => normalise(item.text) !== "");

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

  return { matched, placed };
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
    });
  }
});
