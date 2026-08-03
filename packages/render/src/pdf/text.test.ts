import { describe, expect, it } from "vitest";

import { buildPdf } from "../testing/build-pdf.js";
import { readTextPlacements } from "./text.js";

const placementsOf = (contents: string) => readTextPlacements(buildPdf({ contents }));

describe("readTextPlacements", () => {
  it("reports the baseline origin relative to the top of the page", async () => {
    const [placement] = await placementsOf("BT /F0 12 Tf 100 700 Td (Hello) Tj ET");

    expect(placement?.text).toBe("Hello");
    expect(placement?.leftPt).toBe(100);
    expect(placement?.baselinePt).toBe(92);
    expect(placement?.fontSizePt).toBe(12);
    expect(placement?.pageIndex).toBe(0);
  });

  it("accounts for a transform applied outside the text object", async () => {
    const [placement] = await placementsOf(
      "q 1 0 0 1 50 -100 cm BT /F0 12 Tf 100 700 Td (Hello) Tj ET Q",
    );

    expect(placement?.leftPt).toBe(150);
    expect(placement?.baselinePt).toBe(192);
  });

  it("scales the reported font size by the transform", async () => {
    const [placement] = await placementsOf("q 2 0 0 2 0 0 cm BT /F0 12 Tf 0 300 Td (Hi) Tj ET Q");

    expect(placement?.fontSizePt).toBe(24);
    expect(placement?.baselinePt).toBe(192);
  });

  it("keeps each shown string as its own placement, in drawing order", async () => {
    const placements = await placementsOf(
      "BT /F0 12 Tf 10 700 Td (first) Tj 0 -20 Td (second) Tj ET",
    );

    expect(placements.map((placement) => placement.text)).toStrictEqual(["first", "second"]);
    expect(placements.map((placement) => placement.baselinePt)).toStrictEqual([92, 112]);
  });

  it("returns nothing when the page shows no text", async () => {
    expect(await placementsOf("q 100 0 0 50 200 300 cm /Im0 Do Q")).toStrictEqual([]);
  });
});
