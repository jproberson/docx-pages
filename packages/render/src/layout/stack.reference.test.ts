import { describe, expect, it } from "vitest";

import { layOutDocument, lookupFontMetrics } from "@onepager/core";

import { readReferenceDocument } from "../testing/documents.js";
import { referenceCases, suppliedFaces, type ReferenceCase } from "../testing/cases.js";

const layoutOf = (each: ReferenceCase) => {
  const supplied = suppliedFaces();
  const result = layOutDocument(readReferenceDocument(each), (request) =>
    lookupFontMetrics(request, supplied),
  );
  if (result.kind !== "laid-out") throw new Error(`blocked: ${result.blocker.kind}`);
  return result;
};

const topOf = (
  boxes: readonly { readonly index: number; readonly topPt: number }[],
  at: number,
) => {
  const box = boxes.find((each) => each.index === at);
  if (box === undefined) throw new Error(`no paragraph ${String(at)}`);
  return box.topPt;
};

// Word reports a paragraph's top against the page it landed on, so the boxes are
// read the same way: page by page, in order.
const bodyBoxes = (each: ReferenceCase) => layoutOf(each).pages.flatMap((page) => page.body);

const CASES = referenceCases();

describe.skipIf(CASES.length === 0)("paragraph stack against Word", () => {
  for (const each of CASES) {
    describe(each.id, () => {
      for (const { index, topPt } of each.headerTopsPt) {
        it(`puts header paragraph ${String(index)} where Word put it`, () => {
          expect(Math.abs(topOf(layoutOf(each).header, index) - topPt)).toBeLessThan(
            each.tolerancePt,
          );
        });
      }

      for (const { index, topPt } of each.bodyTopsPt) {
        it(`puts body paragraph ${String(index)} where Word put it`, () => {
          expect(Math.abs(topOf(bodyBoxes(each), index) - topPt)).toBeLessThan(each.tolerancePt);
        });
      }

      // Each top is only known to about a tenth of a point, but drift would
      // compound across the paragraphs between them into far more than that.
      it.runIf(each.bodyTopsPt.length > 1)("accumulates no drift between two known tops", () => {
        const [first, last] = [each.bodyTopsPt.at(0), each.bodyTopsPt.at(-1)];
        if (first === undefined || last === undefined) throw new Error("need two tops");
        const body = bodyBoxes(each);
        const span = topOf(body, last.index) - topOf(body, first.index);
        expect(Math.abs(span - (last.topPt - first.topPt))).toBeLessThan(0.1);
      });

      it.runIf(each.bodyTopPt !== null)("starts the body where Word starts it", () => {
        expect(Math.abs(layoutOf(each).bodyTopPt - (each.bodyTopPt ?? 0))).toBeLessThan(
          each.tolerancePt,
        );
      });

      it("starts the body below the header rather than at the top margin", () => {
        const { bodyTopPt, headerTopPt, headerHeightPt, page } = layoutOf(each);
        expect(bodyTopPt).toBeGreaterThan(page.margin.topTwips / 20);
        expect(bodyTopPt).toBeCloseTo(headerTopPt + headerHeightPt, 9);
      });
    });
  }
});
