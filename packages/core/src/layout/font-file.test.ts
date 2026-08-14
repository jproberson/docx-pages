import { describe, expect, it } from "vitest";

import { isDocxPagesError, type DocxPagesError } from "../errors.js";
import {
  buildCollection,
  buildFace,
  buildSfnt,
  buildWoff,
  buildWoff2,
  type FontFixture,
} from "../testing/build-font.js";
import { readFontFaces, readFontFile, readFontMetrics, readGlyphIndex } from "./font-file.js";
import { lineHeightPt, type GlyphAdvances, type KerningTable } from "./font-metrics.js";
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

// A line is the sum of its characters' advances plus what each pair of them moves,
// which a face states in a legacy `kern` table, in GPOS pair positioning, or in
// both. Everything here is read off the font formats; what Word does with the
// answer is measured elsewhere.
describe("what a pair of a face's characters moves", () => {
  const KERNABLE = { ...FACE, advances: { ...WIDTHS, V: 620, a: 500 } };

  const kerningOf = (fixture: FontFixture): KerningTable =>
    readFontFile(buildSfnt(fixture)).kerning;

  function movementIn(fixture: FontFixture): (pair: string) => number {
    const table = kerningOf(fixture);
    if (table.kind !== "kerning") throw new Error(`expected kerning, got ${table.reason}`);
    return (pair) => table.kerningBetween(pair.codePointAt(0) ?? 0, pair.codePointAt(1) ?? 0);
  }

  it("reads a pair out of a legacy kern table", () => {
    const movedBy = movementIn({ ...KERNABLE, kernPairs: { AV: -80, Aa: -15 } });

    expect(movedBy("AV")).toBe(-80);
    expect(movedBy("Aa")).toBe(-15);
  });

  // A subtable states its length in two bytes and Word's own faces overflow it:
  // Calibri's table is 160254 bytes long and its subtable states 29178, which is
  // 160250 less two whole turns of the field. The pair count is what survives.
  it("reads a kern subtable whose stated length is too small for its own pairs", () => {
    const overflowed = { ...KERNABLE, kernPairs: { AV: -80 }, claimedKernLength: 6 };

    expect(movementIn(overflowed)("AV")).toBe(-80);
  });

  it("says a pair the table does not name moves nothing", () => {
    const movedBy = movementIn({ ...KERNABLE, kernPairs: { AV: -80 } });

    expect(movedBy("VA")).toBe(0);
    expect(movedBy("AB")).toBe(0);
  });

  it("says nothing for a character the face has no glyph for", () => {
    expect(movementIn({ ...KERNABLE, kernPairs: { AV: -80 } })("AZ")).toBe(0);
  });

  it("names the table the pairs came out of", () => {
    expect(kerningOf({ ...KERNABLE, kernPairs: { AV: -80 } })).toMatchObject({ source: "kern" });
    expect(kerningOf({ ...KERNABLE, gposPairs: { AV: -80 } })).toMatchObject({ source: "gpos" });
  });

  it("reads a pair GPOS states one pair at a time", () => {
    const movedBy = movementIn({ ...KERNABLE, gposPairs: { AV: -55, AB: -10, Va: -30 } });

    expect(movedBy("AV")).toBe(-55);
    expect(movedBy("AB")).toBe(-10);
    expect(movedBy("Va")).toBe(-30);
    expect(movedBy("BA")).toBe(0);
  });

  // How a face states the pairs of whole alphabets without naming any of them: a
  // class of first glyphs against a class of second ones.
  it("reads a pair GPOS states as a pair of classes", () => {
    const movedBy = movementIn({
      ...KERNABLE,
      gposClassPairs: {
        firstClasses: ["AV"],
        secondClasses: ["a", "B"],
        values: [
          [0, 0, 0],
          [0, -40, -12],
        ],
      },
    });

    expect(movedBy("Aa")).toBe(-40);
    expect(movedBy("Va")).toBe(-40);
    expect(movedBy("AB")).toBe(-12);
  });

  it("says nothing for a pair whose classes state no movement", () => {
    const movedBy = movementIn({
      ...KERNABLE,
      gposClassPairs: {
        firstClasses: ["A"],
        secondClasses: ["a"],
        values: [
          [0, 0],
          [0, -40],
        ],
      },
    });

    // The second glyph is in no class the table names, and the first is in none
    // the coverage does.
    expect(movedBy("AB")).toBe(0);
    expect(movedBy("Va")).toBe(0);
  });

  // A real face reaches a subtable further into the file than a two-byte offset can
  // name this way, so a reader that stops at the wrapper reads no pairs at all.
  it("follows an extension lookup to the pair positioning under it", () => {
    const movedBy = movementIn({
      ...KERNABLE,
      gposPairs: { AV: -55 },
      gposThroughExtension: true,
    });

    expect(movedBy("AV")).toBe(-55);
  });

  // Which table Word reads where a face states both is unmeasured; `font-file.ts`
  // states which is taken and why, and this is what says so out loud.
  it("takes GPOS over the legacy table where a face states both", () => {
    const both = { ...KERNABLE, kernPairs: { AV: -80 }, gposPairs: { AV: -55 } };

    expect(kerningOf(both)).toMatchObject({ source: "gpos" });
    expect(movementIn(both)("AV")).toBe(-55);
  });

  it("answers nothing at all for a face that states neither table", () => {
    expect(kerningOf(KERNABLE)).toStrictEqual({ kind: "unavailable", reason: "unkerned" });
  });

  it("answers nothing for a face whose pair positioning names no pair", () => {
    const table = kerningOf({ ...KERNABLE, gposPairs: {} });

    expect(table).toStrictEqual({ kind: "unavailable", reason: "unkerned" });
  });

  // Refused rather than half-read: a face states its pairs once, so a table that
  // cannot be read is a face whose pairs are unknown, and some of them moves text
  // to a place neither Word nor the face asked for.
  it("refuses a kern table that states more pairs than it holds", () => {
    const lying = { ...KERNABLE, kernPairs: { AV: -80 }, claimedKernPairs: 40 };

    expect(kerningOf(lying)).toStrictEqual({ kind: "unavailable", reason: "kern-malformed" });
  });

  it("refuses GPOS whose offsets run past the end of it", () => {
    const cut = { ...KERNABLE, gposPairs: { AV: -55 }, cutFromGpos: 8 };

    expect(kerningOf(cut)).toStrictEqual({ kind: "unavailable", reason: "gpos-malformed" });
  });

  // The only movement read is an X advance on the first glyph, which is what
  // kerning a line of text is. A subtable stating anything else is left unread and
  // counted, rather than read at half of what it says or taken as a reason to
  // refuse the pairs beside it: measured on 2026-08-13 over the 472 faces on this
  // machine, every subtable of that kind belongs to a script running the other way
  // and covers no Latin glyph, and Calibri states its Latin pairs beside eight.
  it("reads the pairs it can beside a subtable whose movement it will not guess at", () => {
    const both = { ...KERNABLE, gposPairs: { AV: -55 }, gposRightToLeftPairs: { Ta: -30 } };

    expect(movementIn(both)("AV")).toBe(-55);
    expect(movementIn(both)("Ta")).toBe(0);
    expect(kerningOf(both)).toMatchObject({ subtablesLeftUnread: 1 });
  });

  it("counts nothing left unread where every subtable was read", () => {
    expect(kerningOf({ ...KERNABLE, gposPairs: { AV: -55 } })).toMatchObject({
      subtablesLeftUnread: 0,
    });
    expect(kerningOf({ ...KERNABLE, kernPairs: { AV: -80 } })).toMatchObject({
      subtablesLeftUnread: 0,
    });
  });

  // A face whose only pair positioning states such a movement has kerning and not
  // one pair of it could be read, which is a different answer from a face that
  // states none.
  it("refuses a face whose every pair states a movement it will not guess at", () => {
    const placed = { ...KERNABLE, gposPairs: { AV: -55 }, gposValueFormats: [0x0005, 0] as const };
    const second = {
      ...KERNABLE,
      gposPairs: { AV: -55 },
      gposValueFormats: [0x0004, 0x0004] as const,
    };

    expect(kerningOf(placed)).toStrictEqual({ kind: "unavailable", reason: "gpos-unsupported" });
    expect(kerningOf(second)).toStrictEqual({ kind: "unavailable", reason: "gpos-unsupported" });
  });

  it("refuses a malformed table even where the other one could be read", () => {
    const both = {
      ...KERNABLE,
      kernPairs: { AV: -80 },
      claimedKernPairs: 40,
      gposPairs: { AV: -55 },
    };

    expect(kerningOf(both)).toStrictEqual({ kind: "unavailable", reason: "kern-malformed" });
  });

  // The pairs are stated in glyphs, so a face whose characters cannot be mapped to
  // glyphs has nothing to look them up by.
  it("answers nothing for a face whose character map cannot be read", () => {
    const table = kerningOf({ ...KERNABLE, kernPairs: { AV: -80 }, omit: "cmap" });

    expect(table).toStrictEqual({ kind: "unavailable", reason: "cmap-missing" });
  });

  // The road a measurer takes to the pairs: off the file and onto the face a run
  // resolves to.
  it("hands a face built for a test the pairs its file states", () => {
    const supplied = buildFace({ name: "Meridian", metrics: METRICS, kernPairs: { AV: -80 } });
    const between = supplied.kerning?.kind === "kerning" ? supplied.kerning.kerningBetween : null;

    expect(between?.("A".codePointAt(0) ?? 0, "V".codePointAt(0) ?? 0)).toBe(-80);
    expect(buildFace({ name: "Meridian", metrics: METRICS }).kerning).toBeUndefined();
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

// A file on disk is named for whoever shipped it, so the only way to know what a
// machine can offer a document is to open each font and ask it.
describe("readFontFaces", () => {
  it("reads the family a face belongs to apart from the whole of its own name", () => {
    const light: FontFixture = { ...FACE, faceName: "Meridian Light", familyName: "Meridian" };

    expect(readFontFaces(buildSfnt(light))).toStrictEqual([
      { family: "Meridian", fullName: "Meridian Light", bold: false, italic: false },
    ]);
  });

  it("reads the weight and the slope the face states in head", () => {
    const cuts = [
      { bold: false, italic: false },
      { bold: true, italic: false },
      { bold: false, italic: true },
      { bold: true, italic: true },
    ];

    for (const cut of cuts) {
      const faces = readFontFaces(buildSfnt({ ...FACE, faceName: "Meridian", ...cut }));
      expect(faces[0]).toStrictEqual({ family: "Meridian", fullName: "Meridian", ...cut });
    }
  });

  it("answers for every face of a collection, which holds several", () => {
    const collection = buildCollection([
      { ...FACE, faceName: "Meridian" },
      { ...FACE, faceName: "Meridian Bold", familyName: "Meridian", bold: true },
      { ...FACE, faceName: "Meridian Maths" },
    ]);

    expect(readFontFaces(collection).map((each) => each.fullName)).toStrictEqual([
      "Meridian",
      "Meridian Bold",
      "Meridian Maths",
    ]);
    expect(readFontFaces(collection)[1]?.bold).toBe(true);
  });

  it("calls a face that names no family by its own name, since one is a family of one", () => {
    expect(readFontFaces(buildSfnt({ ...FACE, faceName: "Meridian" }))).toStrictEqual([
      { family: "Meridian", fullName: "Meridian", bold: false, italic: false },
    ]);
  });

  it("reads a face that carries no name table at all rather than refusing it", () => {
    expect(readFontFaces(buildSfnt(FACE))).toStrictEqual([
      { family: "", fullName: "", bold: false, italic: false },
    ]);
  });

  it("refuses what is not a font at all", () => {
    expect(caught(() => readFontFaces(new Uint8Array(4))).code).toBe("font-unreadable");
  });
});
