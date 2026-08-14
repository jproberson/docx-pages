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
import {
  readFontFaces,
  readFontFile,
  readFontMetrics,
  readGlyphIndex,
  type GlyphOutline,
  type OutlineTable,
} from "./font-file.js";
import {
  lineHeightPt,
  type GlyphAdvances,
  type InkBox,
  type InkTable,
  type KerningTable,
  type MathTable,
} from "./font-metrics.js";
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

// A letter's own box is not its line: Word measures a fraction's height off the
// ink of its halves rather than off the face's ascent, measured on 2026-08-13 over
// two fractions of one size that came out 2.64pt apart in height for no reason but
// which letters their halves held.
describe("what a glyph draws", () => {
  const BOXES = {
    A: { left: 20, bottom: 0, right: 640, top: 700 },
    B: { left: 40, bottom: -10, right: 600, top: 690 },
  };
  const DRAWN: FontFixture = { ...FACE, advances: WIDTHS, boxes: BOXES };

  const inkIn = (fixture: FontFixture): InkTable => readFontFile(buildSfnt(fixture)).ink;

  function inkOf(fixture: FontFixture, character: string): InkBox | null {
    const table = inkIn(fixture);
    if (table.kind !== "ink") throw new Error(`expected ink, got ${table.reason}`);
    return table.inkOf(character.codePointAt(0) ?? 0);
  }

  it("reads the box a TrueType glyph states in its own header", () => {
    expect(inkOf(DRAWN, "A")).toStrictEqual(BOXES.A);
    expect(inkOf(DRAWN, "B")).toStrictEqual(BOXES.B);
  });

  // A space is written as a glyph of no length at all, which is not a fault and
  // not a box of zero either: it draws nothing.
  it("answers nothing for a character whose glyph has no outline", () => {
    expect(inkOf(DRAWN, " ")).toBeNull();
  });

  it("answers nothing for a character the face does not map", () => {
    expect(inkOf(DRAWN, "Z")).toBeNull();
  });

  // A face with more outline in it than a two-byte offset can reach states its
  // glyph offsets whole rather than in pairs of bytes.
  it("reads a long loca as well as a short one", () => {
    expect(inkOf({ ...DRAWN, locaFormat: "long" }, "A")).toStrictEqual(BOXES.A);
  });

  // The glyphs are numbered from one in code point order, so A is the second of
  // the three this fixture maps.
  it("answers for a glyph reached without a character, which is how a variant is", () => {
    const table = inkIn(DRAWN);
    expect(table.kind === "ink" && table.inkOfGlyph(2)).toStrictEqual(BOXES.A);
  });

  it("reports a face carrying no outlines at all", () => {
    expect(inkIn({ ...FACE, advances: WIDTHS })).toStrictEqual({
      kind: "unavailable",
      reason: "outlines-missing",
    });
  });

  it("reports a face whose character map cannot be read", () => {
    expect(inkIn({ ...DRAWN, omit: "cmap" })).toStrictEqual({
      kind: "unavailable",
      reason: "cmap-missing",
    });
  });
});

// A PostScript face states no box for a glyph anywhere, so the box is what the
// outline comes to and the charstring has to be followed to find it.
describe("what a PostScript outline comes to", () => {
  // A rectangle, and a curve whose control points stand a quarter above the
  // highest ink it draws.
  const OUTLINES: FontFixture = {
    ...FACE,
    advances: WIDTHS,
    outlines: {
      A: {
        from: [20, 0],
        steps: [{ line: [600, 0] }, { line: [0, 700] }, { line: [-600, 0] }],
      },
      B: { from: [0, 0], steps: [{ curve: [0, 1000, 1000, 0, 0, -1000] }] },
    },
  };

  function inkOf(fixture: FontFixture, character: string): InkBox | null {
    const table = readFontFile(buildSfnt(fixture)).ink;
    if (table.kind !== "ink") throw new Error(`expected ink, got ${table.reason}`);
    return table.inkOf(character.codePointAt(0) ?? 0);
  }

  it("reads the box a path of lines fills", () => {
    expect(inkOf(OUTLINES, "A")).toStrictEqual({ left: 20, bottom: 0, right: 620, top: 700 });
  });

  // The control points reach 1000 and the curve itself reaches 750, which is where
  // it turns. Taking the control points for the ink would draw the letter a
  // quarter taller than it is.
  it("reads a curve at where it turns rather than at where its controls stand", () => {
    expect(inkOf(OUTLINES, "B")).toStrictEqual({ left: 0, bottom: 0, right: 1000, top: 750 });
  });

  it("answers nothing for a glyph whose charstring draws nothing", () => {
    expect(inkOf(OUTLINES, " ")).toBeNull();
  });
});

// What a face says about setting mathematics, which is a table all but a handful
// of faces state nothing of.
describe("what a face says about setting mathematics", () => {
  const CONSTANTS = {
    scriptPercentScaleDown: 73,
    scriptScriptPercentScaleDown: 60,
    delimitedSubFormulaMinHeight: 1500,
    displayOperatorMinHeight: 1250,
    axisHeight: 250,
    fractionRuleThickness: 60,
    fractionNumeratorGapMin: 60,
    fractionDenominatorGapMin: 60,
    radicalDegreeBottomRaisePercent: 65,
  };

  const SETTING: FontFixture = {
    ...FACE,
    advances: { "(": 400, A: 660, B: 640 },
    boxes: {
      "(": { left: 60, bottom: -200, right: 340, top: 700 },
      A: { left: 20, bottom: 0, right: 640, top: 700 },
    },
    math: {
      constants: CONSTANTS,
      minConnectorOverlap: 100,
      italicCorrections: { A: 40, "(": 0 },
      tallerVariants: {
        "(": [
          { character: "(", measurement: 900 },
          { character: "A", measurement: 1400 },
        ],
      },
      widerVariants: { A: [{ character: "B", measurement: 1200 }] },
      tallerPieces: {
        "(": {
          italicCorrection: 30,
          parts: [
            { character: "B", startConnector: 0, endConnector: 200, fullAdvance: 800 },
            {
              character: "A",
              startConnector: 200,
              endConnector: 200,
              fullAdvance: 400,
              extender: true,
            },
          ],
        },
      },
    },
  };

  const mathIn = (fixture: FontFixture): MathTable => readFontFile(buildSfnt(fixture)).math;

  function settingIn(fixture: FontFixture): Extract<MathTable, { kind: "math" }> {
    const table = mathIn(fixture);
    if (table.kind !== "math") throw new Error(`expected math, got ${table.reason}`);
    return table;
  }

  const codePointOf = (character: string): number => character.codePointAt(0) ?? 0;

  it("reads the constants the table states as a value", () => {
    const constants = settingIn(SETTING).constants;

    expect(constants.axisHeight).toBe(250);
    expect(constants.fractionRuleThickness).toBe(60);
    expect(constants.fractionNumeratorGapMin).toBe(60);
    expect(constants.fractionDenominatorGapMin).toBe(60);
  });

  // The four the table states plainly, ahead of the fifty-one it states as a value
  // and a device, and the one percentage it states after all of them.
  it("reads the heights and the percentages the table states plainly", () => {
    const constants = settingIn(SETTING).constants;

    expect(constants.scriptPercentScaleDown).toBe(73);
    expect(constants.scriptScriptPercentScaleDown).toBe(60);
    expect(constants.delimitedSubFormulaMinHeight).toBe(1500);
    expect(constants.displayOperatorMinHeight).toBe(1250);
    expect(constants.radicalDegreeBottomRaisePercent).toBe(65);
  });

  it("answers zero for a constant the table leaves at nothing", () => {
    expect(settingIn(SETTING).constants.overbarRuleThickness).toBe(0);
  });

  it("reads how far a glyph leans past its own advance", () => {
    const setting = settingIn(SETTING);

    expect(setting.italicCorrectionOf(codePointOf("A"))).toBe(40);
    expect(setting.italicCorrectionOf(codePointOf("("))).toBe(0);
  });

  it("answers nothing for a character the corrections do not cover", () => {
    expect(settingIn(SETTING).italicCorrectionOf(codePointOf("B"))).toBe(0);
  });

  // A grown parenthesis is one glyph at a larger size rather than pieces stacked,
  // measured on 2026-08-13 off Word's pdf: its ink was continuous over 21.60pt.
  // Each variant comes back resolved, since a variant glyph has no character to
  // ask the advances or the outlines about.
  it("reads the taller shapes a character grows through, with what each draws", () => {
    const variants = settingIn(SETTING).tallerVariantsOf(codePointOf("("));

    expect(variants.map((each) => each.measurement)).toStrictEqual([900, 1400]);
    expect(variants[0]?.advance).toBe(400);
    expect(variants[1]?.advance).toBe(660);
    expect(variants[1]?.ink).toStrictEqual({ left: 20, bottom: 0, right: 640, top: 700 });
  });

  it("reads the wider shapes apart from the taller ones", () => {
    const setting = settingIn(SETTING);

    expect(setting.widerVariantsOf(codePointOf("A")).map((each) => each.measurement)).toStrictEqual(
      [1200],
    );
    expect(setting.widerVariantsOf(codePointOf("("))).toStrictEqual([]);
    expect(setting.tallerVariantsOf(codePointOf("A"))).toStrictEqual([]);
  });

  // What a character grows through where no one variant reaches far enough: the
  // pieces are stacked, and the middle one may be repeated to fill what is left.
  it("reads the pieces a character grows through, and which of them repeats", () => {
    const pieces = settingIn(SETTING).piecesToGrowTaller(codePointOf("("));

    expect(pieces?.italicCorrection).toBe(30);
    expect(pieces?.parts.map((part) => part.extender)).toStrictEqual([false, true]);
    expect(pieces?.parts[0]).toStrictEqual({
      glyph: 3,
      startConnector: 0,
      endConnector: 200,
      fullAdvance: 800,
      extender: false,
      advance: 640,
      ink: null,
    });
    expect(settingIn(SETTING).minConnectorOverlap).toBe(100);
  });

  it("answers nothing for a character the face grows no other way", () => {
    expect(settingIn(SETTING).piecesToGrowTaller(codePointOf("A"))).toBeNull();
  });

  // What a test elsewhere builds a face out of, so that a rule about setting
  // mathematics can be pinned without a font file off this machine: the same three
  // answers `math.ts` asks a face for, off a face made up for the case.
  it("hands a face built for a test its ink and what it says about mathematics", () => {
    const built = buildFace({
      name: "Meridian Math",
      metrics: METRICS,
      advances: { "(": 400, "𝑏": 550 },
      boxes: {
        "𝑏": { left: 30, bottom: -10, right: 520, top: 720 },
        "(": { left: 60, bottom: -200, right: 340, top: 700 },
      },
      math: {
        constants: { axisHeight: 250, fractionRuleThickness: 60 },
        tallerVariants: {
          "(": [
            { character: "(", measurement: 900 },
            { character: "A", measurement: 1400 },
          ],
        },
      },
    });

    expect(built.ink?.kind === "ink" && built.ink.inkOf(codePointOf("𝑏"))).toStrictEqual({
      left: 30,
      bottom: -10,
      right: 520,
      top: 720,
    });
    expect(built.math?.kind === "math" && built.math.constants.axisHeight).toBe(250);
    expect(
      built.math?.kind === "math" &&
        built.math.tallerVariantsOf(codePointOf("(")).map((each) => each.measurement),
    ).toStrictEqual([900, 1400]);
    expect(built.advances.kind === "advances" && built.advances.advanceFor(codePointOf("("))).toBe(
      400,
    );
  });

  // A variant is a glyph with no character of its own, so what it draws and what it
  // advances have to come back with it.
  it("resolves a built face's variants against the glyphs it wrote them for", () => {
    const built = buildFace({
      name: "Meridian Math",
      metrics: METRICS,
      boxes: { A: { left: 20, bottom: 0, right: 640, top: 700 } },
      math: { tallerVariants: { "(": [{ character: "A", measurement: 1400 }] } },
    });
    const variant =
      built.math?.kind === "math" ? built.math.tallerVariantsOf(codePointOf("("))[0] : null;

    expect(variant?.advance).toBe(METRICS.unitsPerEm / 2);
    expect(variant?.ink).toStrictEqual({ left: 20, bottom: 0, right: 640, top: 700 });
  });

  it("leaves a built face that states neither saying nothing about either", () => {
    const plain = buildFace({ name: "Meridian", metrics: METRICS });

    expect(plain.ink).toBeUndefined();
    expect(plain.math).toBeUndefined();
  });

  it("reports a face that says nothing about mathematics, which is nearly all of them", () => {
    expect(mathIn({ ...FACE, advances: WIDTHS })).toStrictEqual({
      kind: "unavailable",
      reason: "math-missing",
    });
  });

  // Refused rather than half-read, as every other table here is.
  it("refuses a MATH table whose offsets run past the end of it", () => {
    expect(mathIn({ ...SETTING, cutFromMath: 24 })).toStrictEqual({
      kind: "unavailable",
      reason: "math-malformed",
    });
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

// The shape itself, which is what a backend that cannot name a glyph draws
// instead: a browser addresses a face by character and by nothing else.
describe("the outline a face states for a glyph", () => {
  const BOX = { left: 20, bottom: 0, right: 640, top: 700 };
  const DRAWN: FontFixture = { ...FACE, advances: WIDTHS, boxes: { A: BOX } };

  const outlinesIn = (fixture: FontFixture): OutlineTable =>
    readFontFile(buildSfnt(fixture)).outlines;

  function outlineOf(fixture: FontFixture, character: string): GlyphOutline | null {
    const table = outlinesIn(fixture);
    if (table.kind !== "outlines") throw new Error(`expected outlines, got ${table.reason}`);
    return table.outlineOf(character.codePointAt(0) ?? 0);
  }

  // The fixture writes a box as the four corners it has, all of them on the curve.
  it("reads a TrueType contour, closed on the point it started at", () => {
    expect(outlineOf(DRAWN, "A")).toStrictEqual({
      unitsPerEm: 1000,
      contours: [
        {
          from: [20, 0],
          steps: [
            { kind: "line", to: [640, 0] },
            { kind: "line", to: [640, 700] },
            { kind: "line", to: [20, 700] },
            { kind: "line", to: [20, 0] },
          ],
        },
      ],
    });
  });

  it("answers nothing for a character whose glyph draws nothing", () => {
    expect(outlineOf(DRAWN, " ")).toBeNull();
    expect(outlineOf(DRAWN, "Z")).toBeNull();
  });

  it("answers by glyph as well, which is how a math variant is reached", () => {
    const table = outlinesIn(DRAWN);
    expect(table.kind === "outlines" && table.outlineOfGlyph(2)?.contours.length).toBe(1);
  });

  it("reports a face carrying no outlines at all", () => {
    expect(outlinesIn({ ...FACE, advances: WIDTHS })).toStrictEqual({
      kind: "unavailable",
      reason: "outlines-missing",
    });
  });

  // A PostScript face states its outline as the charstring itself, which is the
  // very walk the box is worked out by.
  it("reads a PostScript path back as the face wrote it", () => {
    const outlines: FontFixture = {
      ...FACE,
      advances: WIDTHS,
      outlines: {
        A: {
          from: [20, 0],
          steps: [{ line: [600, 0] }, { curve: [0, 300, 100, 200, -300, 100] }],
        },
      },
    };

    expect(outlineOf(outlines, "A")).toStrictEqual({
      unitsPerEm: 1000,
      contours: [
        {
          from: [20, 0],
          steps: [
            { kind: "line", to: [620, 0] },
            { kind: "cubic", first: [620, 300], second: [720, 500], to: [420, 600] },
          ],
        },
      ],
    });
  });
});

// A font built table by table, for the two things a face states that the fixture
// builder has no way to write: a contour of curves, and a glyph made of another.
function tablesOfSfnt(font: Uint8Array): Map<string, Uint8Array> {
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
  const tables = new Map<string, Uint8Array>();
  for (let index = 0; index < view.getUint16(4); index += 1) {
    const record = 12 + index * 16;
    const tag = String.fromCharCode(...font.subarray(record, record + 4));
    const at = view.getUint32(record + 8);
    tables.set(tag, font.subarray(at, at + view.getUint32(record + 12)));
  }
  return tables;
}

function sfntOfTables(tables: ReadonlyMap<string, Uint8Array>): Uint8Array {
  const written = [...tables];
  const padded = (length: number): number => (length + 3) & ~3;
  const directory = 12 + written.length * 16;
  const total = directory + written.reduce((sum, [, data]) => sum + padded(data.length), 0);

  const out = new Uint8Array(total);
  const view = new DataView(out.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint16(4, written.length);

  let at = directory;
  written.forEach(([tag, data], index) => {
    const record = 12 + index * 16;
    for (let byte = 0; byte < 4; byte += 1) out[record + byte] = tag.charCodeAt(byte);
    view.setUint32(record + 8, at);
    view.setUint32(record + 12, data.length);
    out.set(data, at);
    at += padded(data.length);
  });
  return out;
}

// One contour of four points, the middle two off the curve, which is where a face
// leaves an on-curve point implied halfway between them.
function curvedGlyph(): Uint8Array {
  const glyph = new Uint8Array(34);
  const view = new DataView(glyph.buffer);
  view.setInt16(0, 1);
  view.setInt16(2, 0);
  view.setInt16(4, 0);
  view.setInt16(6, 300);
  view.setInt16(8, 200);
  view.setUint16(10, 3);
  view.setUint16(12, 0);
  glyph[14] = 1;
  glyph[15] = 0;
  glyph[16] = 0;
  glyph[17] = 1;
  [0, 100, 100, 100].forEach((step, at) => {
    view.setInt16(18 + at * 2, step);
  });
  [0, 200, 0, -200].forEach((step, at) => {
    view.setInt16(26 + at * 2, step);
  });
  return glyph;
}

// A glyph that is another glyph, moved where it stands.
function composedGlyph(of: number, acrossPt: number, upPt: number): Uint8Array {
  const glyph = new Uint8Array(18);
  const view = new DataView(glyph.buffer);
  view.setInt16(0, -1);
  view.setUint16(10, 0x0003);
  view.setUint16(12, of);
  view.setInt16(14, acrossPt);
  view.setInt16(16, upPt);
  return glyph;
}

function faceDrawing(glyphs: readonly Uint8Array[]): Uint8Array {
  const tables = tablesOfSfnt(
    buildSfnt({
      ...FACE,
      advances: { A: 660, B: 640 },
      boxes: { A: { left: 0, bottom: 0, right: 1, top: 1 } },
    }),
  );

  const glyf = new Uint8Array(glyphs.reduce((sum, each) => sum + each.length, 0));
  const loca = new Uint8Array((glyphs.length + 1) * 2);
  const view = new DataView(loca.buffer);
  let at = 0;
  glyphs.forEach((glyph, index) => {
    view.setUint16(index * 2, at / 2);
    glyf.set(glyph, at);
    at += glyph.length;
  });
  view.setUint16(glyphs.length * 2, at / 2);

  tables.set("glyf", glyf);
  tables.set("loca", loca);
  return sfntOfTables(tables);
}

describe("a TrueType outline of curves", () => {
  const outlineOfGlyph = (font: Uint8Array, glyph: number): GlyphOutline | null => {
    const table = readFontFile(font).outlines;
    if (table.kind !== "outlines") throw new Error(`expected outlines, got ${table.reason}`);
    return table.outlineOfGlyph(glyph);
  };

  // **Two off-curve points in a row have an on-curve point implied halfway between
  // them.** A reader that misses that draws a straight line through every second
  // curve: here the implied point is (150, 200), halfway between the two controls.
  it("puts back the point a face leaves out between two controls", () => {
    const outline = outlineOfGlyph(faceDrawing([new Uint8Array(0), curvedGlyph()]), 1);

    expect(outline?.contours[0]).toStrictEqual({
      from: [0, 0],
      steps: [
        { kind: "quadratic", control: [100, 200], to: [150, 200] },
        { kind: "quadratic", control: [200, 200], to: [300, 0] },
        { kind: "line", to: [0, 0] },
      ],
    });
  });

  // A composite is other glyphs, each moved where it stands: an accented letter is
  // written this way, and so is much of a maths face.
  it("draws a glyph made of another, moved where it stands", () => {
    const font = faceDrawing([new Uint8Array(0), curvedGlyph(), composedGlyph(1, 500, -50)]);
    const outline = outlineOfGlyph(font, 2);

    expect(outline?.contours[0]?.from).toStrictEqual([500, -50]);
    expect(outline?.contours[0]?.steps[0]).toStrictEqual({
      kind: "quadratic",
      control: [600, 150],
      to: [650, 150],
    });
  });

  it("answers nothing for a glyph of no length at all", () => {
    expect(outlineOfGlyph(faceDrawing([new Uint8Array(0), curvedGlyph()]), 0)).toBeNull();
  });
});
