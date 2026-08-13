import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

import {
  layOutDocument,
  lookupFontMetrics,
  type FaceRequest,
  type MetricsResolver,
} from "@docx-pages/core";

import { authoredCases } from "../authored/cases.js";
import { authoredFace, authoredMetrics } from "../authored/faces.js";
import { agreementWith, type Agreement } from "../pdf/agreement.js";
import { readDrawnText } from "../pdf/text.js";
import { readReferenceDocument } from "../testing/documents.js";
import { referenceCases, suppliedFaces, type ReferenceCase } from "../testing/cases.js";

// A document is laid out in the faces it was measured in: the manifest's for the
// seven reference documents, and for the authored ones the faces they name, which
// the manifest knows nothing about.
type Compared = {
  readonly each: ReferenceCase;
  readonly metricsFor: () => MetricsResolver;
};

const FACE = authoredFace();

const CASES: readonly Compared[] = [
  ...referenceCases().map((each) => ({
    each,
    metricsFor: () => {
      const faces = suppliedFaces();
      return (request: FaceRequest) => lookupFontMetrics(request, faces);
    },
  })),
  ...(FACE === null ? [] : authoredCases().map((each) => ({ each, metricsFor: authoredMetrics }))),
].filter(({ each }) => each.renderedPath !== null && each.textLinesMatched !== null);

// The eight real documents alone. The authored ones ask one question each and four of
// them are written around a rule that is measured and not built, so their pages are
// known to disagree; these are the pages nothing is known to be wrong with.
const REAL_IDS = new Set(referenceCases().map((each) => each.id));
const REAL: readonly Compared[] = CASES.filter(({ each }) => REAL_IDS.has(each.id));

async function compare(compared: Compared): Promise<Agreement> {
  const layout = layOutDocument(readReferenceDocument(compared.each), compared.metricsFor());
  if (layout.kind !== "laid-out") throw new Error(`blocked: ${layout.blocker.kind}`);

  const drawn = await readDrawnText(new Uint8Array(readFileSync(compared.each.renderedPath ?? "")));

  return agreementWith(layout, drawn, compared.each.textTolerancePt);
}

// Every assertion below asks the same question of the same document, and laying it
// out and reading Word's pdf again for each one is what the whole suite's time goes
// on. Each case is compared once and the answer handed round.
const comparisons = new Map<string, Promise<Agreement>>();

function comparisonOf(compared: Compared): Promise<Agreement> {
  const found = comparisons.get(compared.each.id);
  if (found !== undefined) return found;

  const made = compare(compared);
  comparisons.set(compared.each.id, made);
  return made;
}

// Laying a document out and reading a pdf of it back is a second or two of work
// each, and the whole suite runs its files side by side, so the first assertion
// of a case waits far longer than a unit test's own patience allows for.
const COMPARISON_TIMEOUT_MS = 60_000;

describe.skipIf(CASES.length === 0)(
  "text lines against Word",
  { timeout: COMPARISON_TIMEOUT_MS },
  () => {
    for (const compared of CASES) {
      const each = compared.each;
      describe(each.id, () => {
        // Word draws a whole page shrunk where it was asked to, and every line of
        // such a drawing is out of place before a rule is consulted. Three of the
        // corpus came back that way, and the ranking read them as a layout wrong
        // about everything. Asked here so that a re-export gone that way says so
        // once rather than as a hundred lines that moved.
        //
        // A percent of slack, since Word writes a size to the tenth of a point and
        // the reference documents come back at 0.994 and 1.006 for it. The shrunk
        // ones are out by a quarter and more.
        it("draws the text at the size the document asks for", async () => {
          const { drawnScale } = await comparisonOf(compared);
          expect(Math.abs((drawnScale ?? 1) - 1)).toBeLessThan(0.01);
        });

        it("breaks paragraphs into the lines Word broke them into", async () => {
          const { matched } = await comparisonOf(compared);
          expect(matched).toBe(each.textLinesMatched ?? 0);
        });

        it("puts those lines where Word put them, on the page Word put them on", async () => {
          const { placed } = await comparisonOf(compared);
          expect(placed).toBe(each.textLinesPlaced ?? 0);
        });

        it("starts each of a line's runs where Word started it", async () => {
          const { runsMatched, runsPlaced } = await comparisonOf(compared);
          expect(runsMatched).toBe(each.textRunsMatched ?? 0);
          expect(runsPlaced).toBe(each.textRunsPlaced ?? 0);
        });

        it("puts a list's number where Word put it", async () => {
          const { numbersMatched, numbersPlaced } = await comparisonOf(compared);
          expect(numbersMatched).toBe(each.numbersMatched ?? 0);
          expect(numbersPlaced).toBe(each.numbersPlaced ?? 0);
        });
      });
    }
  },
);

// **The floor under the reading that classifies a whole page**, and the reason it is
// here rather than in a unit test: a unit test can say that a page whose lines all
// move together is called shifted, and it cannot say that a page nothing is wrong
// with is called nothing at all. These twenty pages are the pages nothing is wrong
// with, and the raster agrees: eighteen of them come out of two rasterisers cell for
// cell equal to Word's own drawing.
//
// It caught its first wolf-cry on the day it was written. Every bullet and every
// list number was being left over as an item Word drew and we did not, because the
// number is drawn by the list rather than by the paragraph and no line ever spells
// it: 213 items over these eight documents, up to 38% of a page, and seven of the
// eight documents were called `missing` for it. `compareNumbers` claims the item it
// matched now, and the seven come out `agrees` on every page.
describe.skipIf(REAL.length === 0)(
  "each page read as a whole",
  { timeout: COMPARISON_TIMEOUT_MS },
  () => {
    for (const compared of REAL) {
      const each = compared.each;
      describe(each.id, () => {
        it("comes out agreed on every page, with no exception anywhere", async () => {
          const { pages } = await comparisonOf(compared);
          const wrong = pages.filter((page) => page.shape !== "agrees");

          expect(wrong.map((page) => `page ${String(page.index + 1)} ${page.shape}`)).toStrictEqual(
            [],
          );
        });

        it("makes the pages Word's own drawing holds", async () => {
          const { pages, pagesDrawn } = await comparisonOf(compared);

          expect(pages).toHaveLength(pagesDrawn);
        });
      });
    }
  },
);
