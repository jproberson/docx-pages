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

// **A cropped picture is written as the whole image scaled up and clipped to the
// part that is wanted**, so where the image was put and what the page holds are two
// different rectangles and only the second is ink. Word writes every crop this way:
// one reference document's logo is placed 254.6pt wide and drawn 180.
describe("readImagePlacements over a clipped image", () => {
  const inkOf = async (contents: string) =>
    (await readImagePlacements(buildPdf({ contents }))).map((placement) => placement.inkRect);

  const CROPPED = "q 200 300 100 50 re W n 100 0 0 50 200 300 cm /Im0 Do Q";

  it("keeps the whole placement in the rect", async () => {
    expect(await rectsOf(CROPPED)).toStrictEqual([
      { pageIndex: 0, leftPt: 200, topPt: 442, widthPt: 100, heightPt: 50 },
    ]);
  });

  it("cuts the ink down to what the clip leaves", async () => {
    expect(await inkOf("q 200 300 40 50 re W n 100 0 0 50 200 300 cm /Im0 Do Q")).toStrictEqual([
      { pageIndex: 0, leftPt: 200, topPt: 442, widthPt: 40, heightPt: 50 },
    ]);
  });

  it("leaves the ink whole where the clip does not reach it", async () => {
    expect(await inkOf(CROPPED)).toStrictEqual([
      { pageIndex: 0, leftPt: 200, topPt: 442, widthPt: 100, heightPt: 50 },
    ]);
  });

  it("takes the smaller of two clips laid over each other", async () => {
    const twice = "q 200 300 80 50 re W n 200 300 40 50 re W n 100 0 0 50 200 300 cm /Im0 Do Q";
    expect(await inkOf(twice)).toStrictEqual([
      { pageIndex: 0, leftPt: 200, topPt: 442, widthPt: 40, heightPt: 50 },
    ]);
  });

  it("says an image the clip leaves nothing of is ink nowhere", async () => {
    expect(await inkOf("q 0 0 50 50 re W n 100 0 0 50 200 300 cm /Im0 Do Q")).toStrictEqual([null]);
  });

  // A clip belongs to the `q` it was set inside, so an image after the `Q` is whole.
  it("gives a clip back at the restore that closes it", async () => {
    const after = "q 0 0 50 50 re W n Q q 100 0 0 50 200 300 cm /Im0 Do Q";
    expect(await inkOf(after)).toStrictEqual([
      { pageIndex: 0, leftPt: 200, topPt: 442, widthPt: 100, heightPt: 50 },
    ]);
  });
});
