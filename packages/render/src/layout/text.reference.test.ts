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
import { readTextPlacements } from "../pdf/text.js";
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

async function compare(compared: Compared): Promise<Agreement> {
  const layout = layOutDocument(readReferenceDocument(compared.each), compared.metricsFor());
  if (layout.kind !== "laid-out") throw new Error(`blocked: ${layout.blocker.kind}`);

  const drawn = await readTextPlacements(
    new Uint8Array(readFileSync(compared.each.renderedPath ?? "")),
  );

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
