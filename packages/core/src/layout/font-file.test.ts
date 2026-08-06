import { describe, expect, it } from "vitest";

import { isDocxPagesError, type DocxPagesError } from "../errors.js";
import {
  buildCollection,
  buildSfnt,
  buildWoff,
  buildWoff2,
  type FontFixture,
} from "../testing/build-font.js";
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

const caught = (run: () => unknown): DocxPagesError => {
  try {
    run();
  } catch (error: unknown) {
    if (isDocxPagesError(error)) return error;
    throw error;
  }
  throw new Error("expected a DocxPagesError");
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

  // Word ships Cambria, which is what it falls back on when it can resolve a name
  // no other way, in a collection and in no other form.
  it("reads the face a collection is named for, which is the first of them", () => {
    const other: FontFixture = { ...FACE, ascender: 1, advances: { A: 111 } };
    const file = readFontFile(buildCollection([{ ...FACE, advances: WIDTHS }, other]));

    expect(file.metrics).toStrictEqual(METRICS);
    expect(file.advances.kind === "advances" && advanceOf(file.advances.advanceFor, "A")).toBe(660);
  });

  it("reports a collection with no face in it", () => {
    const collection = buildCollection([FACE]);
    new DataView(collection.buffer).setUint32(8, 0);

    expect(caught(() => readFontMetrics(collection)).code).toBe("font-unreadable");
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

// A symbol face maps its glyphs in the F020 to F0FF page and offers no unicode
// subtable at all, so it is unmeasurable unless that page is read.
describe("a symbol cmap", () => {
  const SYMBOL: FontFixture = {
    ...FACE,
    subtables: ["symbol"],
    advances: { "\uF0A7": 480, "\uF041": 500 },
  };

  it("is read when the face offers nothing else", () => {
    expect(advanceOf(advancesOf(SYMBOL), "\uF0A7")).toBe(480);
  });

  it("answers for a character written outside the page the face maps it in", () => {
    expect(advanceOf(advancesOf(SYMBOL), "A")).toBe(500);
  });

  it("answers for a character the face maps in the low byte", () => {
    const low: FontFixture = { ...FACE, subtables: ["symbol"], advances: { A: 500 } };
    expect(advanceOf(advancesOf(low), "\uF041")).toBe(500);
  });

  it("still reports a character in neither page as unmapped", () => {
    expect(advanceOf(advancesOf(SYMBOL), "•")).toBeNull();
  });

  it("is not how a unicode face is read, which maps only what it says it maps", () => {
    const unicode: FontFixture = { ...FACE, advances: { A: 500 } };
    expect(advanceOf(advancesOf(unicode), "\uF041")).toBeNull();
  });

  // Symbol declares both encodings, and the unicode one is what is read here, since
  // it maps the face's own page and the Greek and the maths beside it.
  it("makes a face a symbol face by what it declares, not by the subtable read", () => {
    const both: FontFixture = {
      ...FACE,
      subtables: ["unicode", "symbol"],
      notdefAdvance: 600,
      advances: { A: 500 },
    };

    expect(advanceOf(advancesOf(both), "A")).toBe(500);
    expect(advanceOf(advancesOf(both), "\u00A0")).toBe(600);
  });
});

// Word draws a symbol face's own notdef for a character its own page has a place
// for and no glyph at, rather than reaching for another face. Measured off Word's
// pdf of the authored `unmapped-characters` document.
describe("a character a symbol face has no glyph for", () => {
  const SYMBOL: FontFixture = {
    ...FACE,
    subtables: ["symbol"],
    notdefAdvance: 600,
    advances: { "\uF0A7": 480 },
  };

  it("is drawn at the face's own notdef where its page has a place for it", () => {
    expect(advanceOf(advancesOf(SYMBOL), "\u00A0")).toBe(600);
    expect(advanceOf(advancesOf(SYMBOL), "\uF0A0")).toBe(600);
  });

  it("is unmapped where the page has no place for it at all", () => {
    expect(advanceOf(advancesOf(SYMBOL), "\u2022")).toBeNull();
  });

  it("leaves a unicode face reporting it unmapped, which is a question of its own", () => {
    const unicode: FontFixture = { ...FACE, notdefAdvance: 600, advances: { A: 500 } };
    expect(advanceOf(advancesOf(unicode), "\u00A0")).toBeNull();
  });
});

// The narrow no-break space is drawn out of another of the face's own glyphs, its
// thin space, which is measured off the same pdf.
describe("a narrow no-break space", () => {
  it("takes the face's thin space where the face maps no glyph for it", () => {
    const thin: FontFixture = { ...FACE, advances: { "\u2009": 205 } };
    expect(advanceOf(advancesOf(thin), "\u202F")).toBe(205);
  });

  it("keeps the face's own glyph where it has one", () => {
    const both: FontFixture = { ...FACE, advances: { "\u2009": 205, "\u202F": 260 } };
    expect(advanceOf(advancesOf(both), "\u202F")).toBe(260);
  });

  it("is unmapped in a face that has no thin space either", () => {
    expect(advanceOf(advancesOf({ ...FACE, advances: WIDTHS }), "\u202F")).toBeNull();
  });
});
