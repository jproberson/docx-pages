import { describe, expect, it } from "vitest";

import { readImagePlacements } from "./placements.js";
import { hasReference, readReference } from "../testing/references.js";

const REFERENCE = "reference.pdf";

const round = (value: number): number => Math.round(value * 100) / 100;

describe.skipIf(!hasReference(REFERENCE))("placements against Word output", () => {
  it("finds every image Word drew on the first page, where Word drew it", async () => {
    const placements = await readImagePlacements(readReference(REFERENCE));
    const firstPage = placements
      .filter((placement) => placement.rect.pageIndex === 0)
      .map(({ rect }) => ({
        leftPt: round(rect.leftPt),
        topPt: round(rect.topPt),
        widthPt: round(rect.widthPt),
        heightPt: round(rect.heightPt),
      }));

    expect(firstPage).toStrictEqual([
      { leftPt: 445.35, topPt: 8.85, widthPt: 254.6, heightPt: 71.2 },
      { leftPt: 25.5, topPt: 21.75, widthPt: 42, heightPt: 40.2 },
      { leftPt: 303.75, topPt: 494.97, widthPt: 265.44, heightPt: 149.15 },
      { leftPt: 303.75, topPt: 259.73, widthPt: 265.5, heightPt: 197.55 },
    ]);
  });

  it("repeats the header images on the second page", async () => {
    const placements = await readImagePlacements(readReference(REFERENCE));
    const pages = new Set(placements.map((placement) => placement.rect.pageIndex));
    expect([...pages].sort()).toStrictEqual([0, 1]);
  });
});
