import { describe, expect, it } from "vitest";

import { buildPdf } from "../testing/build-pdf.js";
import { readImagePlacements } from "./placements.js";

const rectsOf = async (contents: string) =>
  (await readImagePlacements(buildPdf({ contents }))).map((placement) => placement.rect);

describe("readImagePlacements", () => {
  it("converts a cm matrix into a page-top-relative rectangle", async () => {
    expect(await rectsOf("q 100 0 0 50 200 300 cm /Im0 Do Q")).toStrictEqual([
      { pageIndex: 0, leftPt: 200, topPt: 442, widthPt: 100, heightPt: 50 },
    ]);
  });

  it("composes nested transforms", async () => {
    expect(await rectsOf("q 2 0 0 2 100 100 cm 50 0 0 25 50 100 cm /Im0 Do Q")).toStrictEqual([
      { pageIndex: 0, leftPt: 200, topPt: 442, widthPt: 100, heightPt: 50 },
    ]);
  });

  it("restores the matrix after a Q so later images are unaffected", async () => {
    expect(
      await rectsOf("q 10 0 0 10 500 500 cm Q q 100 0 0 50 200 300 cm /Im0 Do Q"),
    ).toStrictEqual([{ pageIndex: 0, leftPt: 200, topPt: 442, widthPt: 100, heightPt: 50 }]);
  });

  it("takes the bounding box of a rotated placement", async () => {
    const rotated = await rectsOf("q 0 40 -20 0 300 300 cm /Im0 Do Q");
    expect(rotated).toStrictEqual([
      { pageIndex: 0, leftPt: 280, topPt: 452, widthPt: 20, heightPt: 40 },
    ]);
  });

  it("normalises a negative scale into a positive rectangle", async () => {
    expect(await rectsOf("q -100 0 0 50 300 300 cm /Im0 Do Q")).toStrictEqual([
      { pageIndex: 0, leftPt: 200, topPt: 442, widthPt: 100, heightPt: 50 },
    ]);
  });

  it("returns nothing when the page paints no images", async () => {
    expect(await rectsOf("q 1 0 0 1 0 0 cm Q")).toStrictEqual([]);
  });
});
