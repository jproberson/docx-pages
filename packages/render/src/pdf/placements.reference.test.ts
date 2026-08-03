import { describe, expect, it } from "vitest";

import { readImagePlacements } from "./placements.js";
import { readRenderedPages } from "../testing/documents.js";
import { referenceCases } from "../testing/cases.js";

const round = (value: number): number => Math.round(value * 100) / 100;

const CASES = referenceCases().filter((each) => each.renderedPath !== null);

describe.skipIf(CASES.length === 0)("placements against Word output", () => {
  for (const each of CASES) {
    describe(each.id, () => {
      it.runIf(each.renderedImagesPt.length > 0)(
        "finds every image Word drew on the first page, where Word drew it",
        async () => {
          const placements = await readImagePlacements(readRenderedPages(each));
          const firstPage = placements
            .filter((placement) => placement.rect.pageIndex === 0)
            .map(({ rect }) => ({
              leftPt: round(rect.leftPt),
              topPt: round(rect.topPt),
              widthPt: round(rect.widthPt),
              heightPt: round(rect.heightPt),
            }));

          expect(firstPage).toStrictEqual([...each.renderedImagesPt]);
        },
      );

      it.runIf(each.renderedPageIndexes.length > 0)(
        "draws images on the pages Word drew them on",
        async () => {
          const placements = await readImagePlacements(readRenderedPages(each));
          const pages = new Set(placements.map((placement) => placement.rect.pageIndex));
          expect([...pages].sort()).toStrictEqual([...each.renderedPageIndexes]);
        },
      );
    });
  }
});
