import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

import { lineHeightPt, readFontMetrics } from "@onepager/core";

// Not copied into this repo: the face is proprietary. Point at wherever it lives.
const FONT_PATH = resolve(
  process.env["ONEPAGER_TEST_FONT"] ??
    `${process.env["HOME"] ?? ""}/fonts/TestFace-Medium.woff`,
);

describe.skipIf(!existsSync(FONT_PATH))("readFontMetrics on a real corporate font", () => {
  it("reads the metrics that Word used to lay out the reference documents", () => {
    const result = readFontMetrics(new Uint8Array(readFileSync(FONT_PATH)));

    expect(result.format).toBe("woff");
    expect(result.metrics).toStrictEqual({
      unitsPerEm: 1000,
      ascender: 971,
      descender: -242,
      lineGap: 0,
    });
  });

  it("yields the line height ratio measured from Word's own output", () => {
    const { metrics } = readFontMetrics(new Uint8Array(readFileSync(FONT_PATH)));
    expect(lineHeightPt(metrics, 1)).toBeCloseTo(1.213, 3);
  });
});
