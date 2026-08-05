import { describe, expect, it } from "vitest";

import { buildPdf } from "../testing/build-pdf.js";
import { readFillPlacements } from "./fills.js";

const fillsOf = async (contents: string) =>
  (await readFillPlacements(buildPdf({ contents }))).map(
    ({ color, leftPt, topPt, widthPt, heightPt }) => ({
      color,
      leftPt,
      topPt,
      widthPt,
      heightPt,
    }),
  );

describe("readFillPlacements", () => {
  it("reports a filled rectangle from the top of the page down", async () => {
    expect(await fillsOf("1 0 0 rg 200 300 100 50 re f")).toStrictEqual([
      { color: "#ff0000", leftPt: 200, topPt: 442, widthPt: 100, heightPt: 50 },
    ]);
  });

  it("takes a path drawn line by line for the rectangle it covers", async () => {
    expect(await fillsOf("0 0 1 rg 10 10 m 40 10 l 40 30 l 10 30 l h f")).toStrictEqual([
      { color: "#0000ff", leftPt: 10, topPt: 762, widthPt: 30, heightPt: 20 },
    ]);
  });

  it("puts a path through the transform in force where it was drawn", async () => {
    expect(await fillsOf("q 2 0 0 2 100 100 cm 0 g 10 10 20 20 re f Q")).toStrictEqual([
      { color: "#000000", leftPt: 120, topPt: 632, widthPt: 40, heightPt: 40 },
    ]);
  });

  it("leaves a path that is stroked or only clips, neither being a block of colour", async () => {
    expect(await fillsOf("1 0 0 rg 10 10 20 20 re S 10 50 20 20 re W n")).toStrictEqual([]);
  });

  it("keeps the colour each path was filled in", async () => {
    const drawn = await fillsOf("1 0 0 rg 0 0 10 10 re f 0 1 0 rg 20 0 10 10 re f");
    expect(drawn.map((fill) => fill.color)).toStrictEqual(["#ff0000", "#00ff00"]);
  });
});
