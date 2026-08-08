import { describe, expect, it } from "vitest";

import { isDocxPagesError, type DocxPagesError } from "../errors.js";
import {
  buildCollection,
  buildSfnt,
  buildWoff,
  buildWoff2,
  type FontFixture,
} from "../testing/build-font.js";
import { readFontFile, readFontMetrics, readGlyphIndex } from "./font-file.js";
import { lineHeightPt, type GlyphAdvances } from "./font-metrics.js";
import type { CodeToGlyph } from "./glyphs.js";

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

  // Word borrows a maths letter and a hyphen from Cambria Math, which is the
  // second face of the file Cambria is the first of.
  it("reads a face out of a collection by name rather than by where it sits", () => {
    const maths: FontFixture = { ...FACE, faceName: "Meridian Math", advances: { A: 111 } };
    const file = readFontFile(
      buildCollection([{ ...FACE, faceName: "Meridian", advances: WIDTHS }, maths]),
      "Meridian Math",
    );

    expect(file.advances.kind === "advances" && advanceOf(file.advances.advanceFor, "A")).toBe(111);
  });

  it("reports a collection that holds no face of the name asked for", () => {
    const collection = buildCollection([{ ...FACE, faceName: "Meridian", advances: WIDTHS }]);

    expect(caught(() => readFontFile(collection, "Meridian Math")).code).toBe("font-face-missing");
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

// Which face Word draws a character out of when the stated one has no glyph for it
// turns on whether that face has serifs, and nothing in a document says: the file
// itself is asked. PANOSE classifies a Latin text face's serifs, and its styles
// from 11 up are the sans ones.
describe("what a face says it is", () => {
  const classified = (panoseFamily: number, panoseSerifStyle: number): boolean =>
    readFontFile(buildSfnt({ ...FACE, panoseFamily, panoseSerifStyle })).sansSerif;

  const LATIN_TEXT = 2;
  const PICTORIAL = 5;

  it("reads a Latin text face of normal sans as one without serifs", () => {
    expect(classified(LATIN_TEXT, 11)).toBe(true);
  });

  // Calibri classifies itself as rounded, which is the last of the sans styles.
  it("reads the rounded and flared styles as sans as well", () => {
    expect(classified(LATIN_TEXT, 14)).toBe(true);
    expect(classified(LATIN_TEXT, 15)).toBe(true);
  });

  it("reads the cove styles Times New Roman and Cambria claim as serif", () => {
    expect(classified(LATIN_TEXT, 2)).toBe(false);
    expect(classified(LATIN_TEXT, 4)).toBe(false);
  });

  // The byte means something else entirely under another family type: Wingdings is
  // a pictorial face whose style byte is 0, and a pictorial face is not a sans one
  // whatever it says there.
  it("reads a face of any other family as not a sans one", () => {
    expect(classified(PICTORIAL, 11)).toBe(false);
    expect(classified(PICTORIAL, 0)).toBe(false);
  });

  it("reads a face that classifies itself not at all as not a sans one", () => {
    expect(readFontFile(buildSfnt(FACE)).sansSerif).toBe(false);
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

// The number itself, which layout never needs and a writer embedding the face
// cannot do without: a font written under Identity-H is addressed by glyph.
describe("readGlyphIndex", () => {
  // `buildSfnt` numbers a fixture's glyphs from one in code point order, leaving
  // zero to .notdef as a real face does.
  const glyphsOf = (fixture: FontFixture): CodeToGlyph => readGlyphIndex(buildSfnt(fixture));

  const glyphOf = (glyphFor: CodeToGlyph, character: string): number =>
    glyphFor(character.codePointAt(0) ?? 0);

  it("answers the glyph the face draws each character with", () => {
    const glyphFor = glyphsOf({ ...FACE, advances: WIDTHS });

    expect(glyphOf(glyphFor, " ")).toBe(1);
    expect(glyphOf(glyphFor, "A")).toBe(2);
    expect(glyphOf(glyphFor, "B")).toBe(3);
  });

  it("answers .notdef for a character the face does not map", () => {
    expect(glyphOf(glyphsOf({ ...FACE, advances: WIDTHS }), "Z")).toBe(0);
  });

  it("reads a format 12 cmap, which maps past the basic plane", () => {
    const wide: FontFixture = { ...FACE, cmapFormat: 12, advances: { A: 500, "\u{1F600}": 900 } };
    expect(glyphOf(glyphsOf(wide), "\u{1F600}")).toBe(2);
  });

  // The same two aliases the advances are read through, so a character measured at
  // one glyph's width is never written as another.
  it("follows a symbol face into the page it maps its glyphs in", () => {
    const symbol: FontFixture = { ...FACE, subtables: ["symbol"], advances: { "\uF041": 500 } };
    expect(glyphOf(glyphsOf(symbol), "A")).toBe(1);
  });

  it("takes the face's thin space for a narrow no-break space it has no glyph for", () => {
    const thin: FontFixture = { ...FACE, advances: { "\u2009": 205 } };
    expect(glyphOf(glyphsOf(thin), "\u202F")).toBe(1);
  });

  it("picks one face out of a collection by name, as the metrics reader does", () => {
    const collection = buildCollection([
      { ...FACE, faceName: "First", advances: { A: 500 } },
      { ...FACE, faceName: "Second", advances: { A: 500, B: 600 } },
    ]);

    expect(readGlyphIndex(collection, "Second")("B".codePointAt(0) ?? 0)).toBe(2);
  });

  // Refused rather than written wrongly: no glyph number stands in for one that
  // could not be read, and a wrong one draws a different letter.
  it("refuses a face whose character map cannot be read", () => {
    const error = caught(() =>
      readGlyphIndex(buildSfnt({ ...FACE, advances: WIDTHS, omit: "cmap" })),
    );

    expect(error.code).toBe("font-glyphs-unreadable");
    expect(error.at).toBe("core/layout/font-file.readGlyphIndex");
    expect(error.context["reason"]).toBe("cmap-missing");
  });

  it("refuses a face whose cmap is in a format nothing here reads", () => {
    const error = caught(() =>
      readGlyphIndex(buildSfnt({ ...FACE, cmapFormat: 6, advances: WIDTHS })),
    );

    expect(error.code).toBe("font-glyphs-unreadable");
    expect(error.context["reason"]).toBe("cmap-unsupported");
  });
});

// Word draws an underline where the drawn face says to rather than at a place of
// its own, which is measured in `font-file.ts` beside the type.
describe("where a face puts the line under its letters", () => {
  it("reads the position and the thickness the post table states", () => {
    const face = readFontFile(
      buildSfnt({ ...FACE, underlinePosition: 232, underlineThickness: 134 }),
    );

    expect(face.underline).toStrictEqual({ position: 232, thickness: 134 });
  });

  // The file states the position as a distance up from the baseline, and a line
  // under the letters is below one. Turned the right way up on the way out, so a
  // caller adding it to a baseline goes down the page as everything else does.
  it("answers a line below the baseline as a positive distance", () => {
    const face = readFontFile(buildSfnt({ ...FACE, underlinePosition: 100 }));

    expect(face.underline?.position).toBeGreaterThan(0);
  });

  it("reads how far the letters lean", () => {
    expect(readFontFile(buildSfnt({ ...FACE, italicAngle: -12 })).italicAngle).toBeCloseTo(-12, 4);
    expect(readFontFile(buildSfnt({ ...FACE, underlinePosition: 1 })).italicAngle).toBe(0);
  });

  // Nothing can be invented for a face that does not say, so a renderer needing a
  // line is told there is no answer rather than given a made-up one.
  it("answers nothing at all for a face stating no post table", () => {
    const face = readFontFile(buildSfnt(FACE));

    expect(face.underline).toBeNull();
    expect(face.italicAngle).toBe(0);
  });
});
