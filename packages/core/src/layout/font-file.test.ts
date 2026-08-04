import { describe, expect, it } from "vitest";

import { isOnePagerError, type OnePagerError } from "../errors.js";
import { buildSfnt, buildWoff, buildWoff2, type FontFixture } from "../testing/build-font.js";
import { readFontFile, readFontMetrics } from "./font-file.js";
import { lineHeightPt, type GlyphAdvances } from "./font-metrics.js";

const FACE: FontFixture = { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 };

const METRICS = {
  unitsPerEm: FACE.unitsPerEm,
  ascender: FACE.ascender,
  descender: FACE.descender,
  lineGap: FACE.lineGap,
};

const WIDTHS = { A: 660, B: 640, " ": 260 };

const caught = (run: () => unknown): OnePagerError => {
  try {
    run();
  } catch (error: unknown) {
    if (isOnePagerError(error)) return error;
    throw error;
  }
  throw new Error("expected a OnePagerError");
};

function advancesOf(fixture: FontFixture): GlyphAdvances {
  const table = readFontFile(buildSfnt(fixture)).advances;
  if (table.kind !== "advances") throw new Error(`expected advances, got ${table.reason}`);
  return table.advanceFor;
}

const advanceOf = (advanceFor: GlyphAdvances, character: string): number | null =>
  advanceFor(character.codePointAt(0) ?? 0);

describe("readFontMetrics", () => {
  it("reads hhea metrics out of a bare sfnt", () => {
    expect(readFontMetrics(buildSfnt(FACE))).toStrictEqual({ format: "sfnt", metrics: METRICS });
  });

  it("reads an uncompressed woff", () => {
    expect(readFontMetrics(buildWoff(FACE))).toStrictEqual({ format: "woff", metrics: METRICS });
  });

  it("inflates a zlib-compressed woff table", () => {
    expect(readFontMetrics(buildWoff(FACE, true))).toStrictEqual({
      format: "woff",
      metrics: METRICS,
    });
  });

  it("keeps a negative descender negative", () => {
    const metrics = readFontMetrics(buildSfnt(FACE)).metrics;
    expect(metrics.descender).toBe(-200);
    expect(lineHeightPt(metrics, 22)).toBeCloseTo(22, 3);
  });

  it("rejects woff2 by name, because brotli is not available everywhere", () => {
    const error = caught(() => readFontMetrics(buildWoff2()));
    expect(error.code).toBe("font-format-unsupported");
    expect(error.context["format"]).toBe("woff2");
  });

  it("reports a missing hhea table rather than inventing metrics", () => {
    const error = caught(() => readFontMetrics(buildSfnt({ ...FACE, omit: "hhea" })));
    expect(error.code).toBe("font-table-missing");
    expect(error.context["table"]).toBe("hhea");
  });

  it("reports a missing head table", () => {
    const error = caught(() => readFontMetrics(buildSfnt({ ...FACE, omit: "head" })));
    expect(error.code).toBe("font-table-missing");
    expect(error.context["table"]).toBe("head");
  });

  it("reports unreadable bytes", () => {
    expect(caught(() => readFontMetrics(new Uint8Array([1, 2, 3, 4, 5]))).code).toBe(
      "font-unreadable",
    );
  });
});

describe("readFontFile", () => {
  it("reads metrics alongside advances", () => {
    const file = readFontFile(buildSfnt({ ...FACE, advances: WIDTHS }));
    expect(file.format).toBe("sfnt");
    expect(file.metrics).toStrictEqual(METRICS);
  });

  it("reads an advance for each character a format 4 cmap maps", () => {
    const advanceFor = advancesOf({ ...FACE, advances: WIDTHS });
    expect(advanceOf(advanceFor, "A")).toBe(660);
    expect(advanceOf(advanceFor, "B")).toBe(640);
    expect(advanceOf(advanceFor, " ")).toBe(260);
  });

  it("reads advances through a format 12 cmap, including outside the basic plane", () => {
    const advanceFor = advancesOf({
      ...FACE,
      cmapFormat: 12,
      advances: { ...WIDTHS, "\u{1d400}": 720 },
    });
    expect(advanceOf(advanceFor, "A")).toBe(660);
    expect(advanceOf(advanceFor, "\u{1d400}")).toBe(720);
  });

  it("reports an unmapped character rather than guessing a width for it", () => {
    expect(advanceOf(advancesOf({ ...FACE, advances: WIDTHS }), "Z")).toBeNull();
  });

  it("reads through a woff container as well as a bare sfnt", () => {
    const table = readFontFile(buildWoff({ ...FACE, advances: WIDTHS }, true)).advances;
    expect(table.kind === "advances" && advanceOf(table.advanceFor, "A")).toBe(660);
  });

  it("gives glyphs past the last long metric that metric's advance", () => {
    const advanceFor = advancesOf({ ...FACE, advances: WIDTHS, longMetrics: 2 });
    expect(advanceOf(advanceFor, " ")).toBe(260);
    expect(advanceOf(advanceFor, "A")).toBe(260);
    expect(advanceOf(advanceFor, "B")).toBe(260);
  });

  it("reports a face with no cmap as unmeasurable", () => {
    const table = readFontFile(buildSfnt({ ...FACE, advances: WIDTHS, omit: "cmap" })).advances;
    expect(table).toStrictEqual({ kind: "unavailable", reason: "cmap-missing" });
  });

  it("reports a face with no hmtx as unmeasurable", () => {
    const table = readFontFile(buildSfnt({ ...FACE, advances: WIDTHS, omit: "hmtx" })).advances;
    expect(table).toStrictEqual({ kind: "unavailable", reason: "hmtx-missing" });
  });

  it("reports a cmap subtable format it cannot read", () => {
    const table = readFontFile(buildSfnt({ ...FACE, advances: WIDTHS, cmapFormat: 6 })).advances;
    expect(table).toStrictEqual({ kind: "unavailable", reason: "cmap-unsupported" });
  });

  it("still reports metrics for a face it cannot measure text with", () => {
    expect(readFontFile(buildSfnt(FACE)).metrics).toStrictEqual(METRICS);
  });
});
