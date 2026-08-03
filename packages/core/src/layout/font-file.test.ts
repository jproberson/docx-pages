import { describe, expect, it } from "vitest";

import { isOnePagerError, type OnePagerError } from "../errors.js";
import { buildSfnt, buildWoff, buildWoff2 } from "../testing/build-font.js";
import { readFontMetrics } from "./font-file.js";
import { lineHeightPt } from "./font-metrics.js";

const MERIDIAN_SANS = { unitsPerEm: 1000, ascender: 971, descender: -242, lineGap: 0 };

const caught = (run: () => unknown): OnePagerError => {
  try {
    run();
  } catch (error: unknown) {
    if (isOnePagerError(error)) return error;
    throw error;
  }
  throw new Error("expected a OnePagerError");
};

describe("readFontMetrics", () => {
  it("reads hhea metrics out of a bare sfnt", () => {
    expect(readFontMetrics(buildSfnt(MERIDIAN_SANS))).toStrictEqual({
      format: "sfnt",
      metrics: MERIDIAN_SANS,
    });
  });

  it("reads an uncompressed woff", () => {
    expect(readFontMetrics(buildWoff(MERIDIAN_SANS))).toStrictEqual({
      format: "woff",
      metrics: MERIDIAN_SANS,
    });
  });

  it("inflates a zlib-compressed woff table", () => {
    expect(readFontMetrics(buildWoff(MERIDIAN_SANS, true))).toStrictEqual({
      format: "woff",
      metrics: MERIDIAN_SANS,
    });
  });

  it("keeps a negative descender negative", () => {
    const metrics = readFontMetrics(buildSfnt(MERIDIAN_SANS)).metrics;
    expect(metrics.descender).toBe(-242);
    expect(lineHeightPt(metrics, 22)).toBeCloseTo(26.686, 3);
  });

  it("rejects woff2 by name, because brotli is not available everywhere", () => {
    const error = caught(() => readFontMetrics(buildWoff2()));
    expect(error.code).toBe("font-format-unsupported");
    expect(error.context["format"]).toBe("woff2");
  });

  it("reports a missing hhea table rather than inventing metrics", () => {
    const error = caught(() => readFontMetrics(buildSfnt({ ...MERIDIAN_SANS, omit: "hhea" })));
    expect(error.code).toBe("font-table-missing");
    expect(error.context["table"]).toBe("hhea");
  });

  it("reports a missing head table", () => {
    const error = caught(() => readFontMetrics(buildSfnt({ ...MERIDIAN_SANS, omit: "head" })));
    expect(error.code).toBe("font-table-missing");
    expect(error.context["table"]).toBe("head");
  });

  it("reports unreadable bytes", () => {
    expect(caught(() => readFontMetrics(new Uint8Array([1, 2, 3, 4, 5]))).code).toBe(
      "font-unreadable",
    );
  });
});
