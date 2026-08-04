import { describe, expect, it } from "vitest";

import { layOutDocument, lookupFontMetrics, type PlacedFloat } from "@onepager/core";

import { readReferenceDocument } from "../testing/documents.js";
import { referenceCases, suppliedFaces, type ReferenceCase } from "../testing/cases.js";

const layoutOf = (each: ReferenceCase) => {
  const supplied = suppliedFaces();
  const result = layOutDocument(readReferenceDocument(each), (name) =>
    lookupFontMetrics(name, supplied),
  );
  if (result.kind !== "laid-out") throw new Error(`blocked: ${result.blocker.kind}`);
  return result;
};

const floatsOf = (each: ReferenceCase): readonly PlacedFloat[] => {
  const { headerFloats, bodyFloats } = layoutOf(each);
  return [...headerFloats, ...bodyFloats];
};

const at = (floats: readonly PlacedFloat[], index: number): PlacedFloat => {
  const found = floats[index];
  if (found === undefined) throw new Error(`no float ${String(index)}`);
  return found;
};

const overlaps = (one: PlacedFloat, other: PlacedFloat): boolean =>
  one.leftPt < other.leftPt + other.widthPt &&
  other.leftPt < one.leftPt + one.widthPt &&
  one.topPt < other.topPt + other.heightPt &&
  other.topPt < one.topPt + one.heightPt;

const CASES = referenceCases();

describe.skipIf(CASES.length === 0)("float placement against Word", () => {
  for (const each of CASES) {
    describe(each.id, () => {
      for (const expected of each.floatsPt) {
        it(`places float ${String(expected.index)} horizontally where Word did`, () => {
          expect(at(floatsOf(each), expected.index).leftPt).toBeCloseTo(expected.leftPt, 1);
        });

        it(`places float ${String(expected.index)} vertically where Word did`, () => {
          expect(Math.abs(at(floatsOf(each), expected.index).topPt - expected.topPt)).toBeLessThan(
            each.tolerancePt,
          );
        });
      }

      for (const expected of each.inlinesPt) {
        it.runIf(expected.leftPt !== null)(
          `places inline drawing ${String(expected.index)} horizontally where Word did`,
          () => {
            const found = layoutOf(each).bodyInlines[expected.index];
            expect(found?.leftPt).toBeCloseTo(expected.leftPt ?? 0, 1);
          },
        );

        it.runIf(expected.topPt !== null)(
          `places inline drawing ${String(expected.index)} vertically where Word did`,
          () => {
            const found = layoutOf(each).bodyInlines[expected.index];
            expect(Math.abs((found?.topPt ?? 0) - (expected.topPt ?? 0))).toBeLessThan(
              each.tolerancePt,
            );
          },
        );
      }

      for (const [one, other] of each.disjointFloatPairs) {
        it(`keeps floats ${String(one)} and ${String(other)} clear of each other`, () => {
          const floats = floatsOf(each);
          expect(overlaps(at(floats, one), at(floats, other))).toBe(false);
        });
      }

      it.runIf(each.headerFloatCount !== null)("finds every float in the header", () => {
        expect(layoutOf(each).headerFloats).toHaveLength(each.headerFloatCount ?? 0);
      });

      it.runIf(each.leastBodyFloatCount !== null)("finds every float in the body", () => {
        expect(layoutOf(each).bodyFloats.length).toBeGreaterThanOrEqual(
          each.leastBodyFloatCount ?? 0,
        );
      });

      // A negative paragraph-relative offset is the case other engines get wrong:
      // the object belongs above the paragraph that anchors it, not clamped to it.
      it("lets a negatively offset float rise above the paragraph that anchors it", () => {
        const { header, body, headerFloats, bodyFloats } = layoutOf(each);
        const rising = [
          ...headerFloats.map((float) => ({ float, boxes: header })),
          ...bodyFloats.map((float) => ({ float, boxes: body })),
        ].filter(
          ({ float }) =>
            float.anchor.vertical.kind === "offset" &&
            float.anchor.vertical.from === "paragraph" &&
            float.anchor.vertical.offsetEmu < 0,
        );

        for (const { float, boxes } of rising) {
          const anchoring = boxes.find((box) => box.index === float.anchor.paragraphIndex);
          expect(float.topPt).toBeLessThan(anchoring?.topPt ?? Number.POSITIVE_INFINITY);
        }
      });
    });
  }
});
