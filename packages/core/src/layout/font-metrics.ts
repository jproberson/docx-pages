export type FontMetrics = {
  readonly unitsPerEm: number;
  readonly ascender: number;
  readonly descender: number;
  readonly lineGap: number;
};

// An advance is in font units, so it scales by the same unitsPerEm as the metrics
// that came out of the same file. Unmapped characters read back as null rather
// than as a guessed width.
export type GlyphAdvances = (codePoint: number) => number | null;

export type AdvancesUnavailable =
  | "unsupplied"
  | "style-unsupplied"
  | "cmap-missing"
  | "cmap-unsupported"
  | "cmap-malformed"
  | "hmtx-missing"
  | "hmtx-malformed";

export type AdvanceTable =
  | {
      readonly kind: "advances";
      readonly advanceFor: GlyphAdvances;
      // What the face's missing-glyph box advances, which is the last width left
      // when no face at all maps a character and the caller would still rather
      // have a page than a refusal. A face built by hand rather than read out of
      // a file may not say.
      readonly notDefAdvance?: number;
    }
  | { readonly kind: "unavailable"; readonly reason: AdvancesUnavailable };

// Vertical metrics alone place empty paragraphs and floats; measuring text needs
// the advances too, and only a caller-supplied font file carries them.
export const NO_ADVANCES: AdvanceTable = { kind: "unavailable", reason: "unsupplied" };

// How far one character of a face sits from the next beyond the sum of their
// advances, in the same font units the advances are in. Negative for a pair that
// closes up, which is nearly every pair a face states. Zero for a pair the face
// says nothing about, so a caller adds it without asking whether there is one.
export type PairKerning = (leftCodePoint: number, rightCodePoint: number) => number;

// Kerning is read through the same character map the advances are, so every way
// they can be unavailable is a way it can be. The rest are its own: `unkerned` is
// a face that states no pairs anywhere, and the other three are a face whose
// pairs are there and could not be honoured.
export type KerningUnavailable =
  AdvancesUnavailable | "unkerned" | "kern-malformed" | "gpos-malformed" | "gpos-unsupported";

// Which of the two tables the pairs were read out of. Carried out so that a
// measurement can say which one Word agrees with; `font-file.ts` states which is
// preferred and why.
export type KerningSource = "kern" | "gpos";

export type KerningTable =
  | {
      readonly kind: "kerning";
      readonly source: KerningSource;
      readonly kerningBetween: PairKerning;
      // How many of the face's pair subtables stated their movement in a way the
      // reader will not guess at and were left unread. Every one measured belongs
      // to a script that runs the other way and covers no Latin glyph at all;
      // `font-file.ts` records the faces and why one of these is not a refusal.
      readonly subtablesLeftUnread: number;
    }
  | { readonly kind: "unavailable"; readonly reason: KerningUnavailable };

export const NO_KERNING: KerningTable = { kind: "unavailable", reason: "unsupplied" };

// What one glyph actually draws, in font units, measured from the origin its
// advance starts at and with `top` above the baseline as the face states it. A
// letter's box is not its line: `l` reaches higher than `r` does, and Word
// measures a fraction's height off this rather than off the face's own ascent
// (measured 2026-08-13, two fractions of the same size 2.64pt apart in height for
// no reason but which letters their halves held).
export type InkBox = {
  readonly left: number;
  readonly right: number;
  readonly bottom: number;
  readonly top: number;
};

// Null for a character the face draws nothing for, a space above all, which has
// an advance and no ink at all.
export type GlyphInk = (codePoint: number) => InkBox | null;

// The outlines are reached through the same character map the advances are, so
// every way they can be unavailable is a way ink can be. The rest are its own:
// a face carrying neither kind of outline, one whose outlines say something this
// reader cannot follow, and one whose tables run past themselves.
export type InkUnavailable =
  AdvancesUnavailable | "outlines-missing" | "outlines-unsupported" | "outlines-malformed";

export type InkTable =
  | {
      readonly kind: "ink";
      readonly inkOf: GlyphInk;
      // The same answer for a glyph reached without a character, which is how a
      // math variant is named: the taller parenthesis has no code point at all.
      readonly inkOfGlyph: (glyph: number) => InkBox | null;
    }
  | { readonly kind: "unavailable"; readonly reason: InkUnavailable };

export const NO_INK: InkTable = { kind: "unavailable", reason: "unsupplied" };

// The constants the MATH table states as a value and a device offset, in the
// order the table writes them. The order is the table's own layout, so this array
// is what both the reader and a fixture walk, and neither can drift from the
// other.
export const MATH_VALUE_CONSTANTS = [
  "mathLeading",
  "axisHeight",
  "accentBaseHeight",
  "flattenedAccentBaseHeight",
  "subscriptShiftDown",
  "subscriptTopMax",
  "subscriptBaselineDropMin",
  "superscriptShiftUp",
  "superscriptShiftUpCramped",
  "superscriptBottomMin",
  "superscriptBaselineDropMax",
  "subSuperscriptGapMin",
  "superscriptBottomMaxWithSubscript",
  "spaceAfterScript",
  "upperLimitGapMin",
  "upperLimitBaselineRiseMin",
  "lowerLimitGapMin",
  "lowerLimitBaselineDropMin",
  "stackTopShiftUp",
  "stackTopDisplayStyleShiftUp",
  "stackBottomShiftDown",
  "stackBottomDisplayStyleShiftDown",
  "stackGapMin",
  "stackDisplayStyleGapMin",
  "stretchStackTopShiftUp",
  "stretchStackBottomShiftDown",
  "stretchStackGapAboveMin",
  "stretchStackGapBelowMin",
  "fractionNumeratorShiftUp",
  "fractionNumeratorDisplayStyleShiftUp",
  "fractionDenominatorShiftDown",
  "fractionDenominatorDisplayStyleShiftDown",
  "fractionNumeratorGapMin",
  "fractionNumDisplayStyleGapMin",
  "fractionRuleThickness",
  "fractionDenominatorGapMin",
  "fractionDenomDisplayStyleGapMin",
  "skewedFractionHorizontalGap",
  "skewedFractionVerticalGap",
  "overbarVerticalGap",
  "overbarRuleThickness",
  "overbarExtraAscender",
  "underbarVerticalGap",
  "underbarRuleThickness",
  "underbarExtraDescender",
  "radicalVerticalGap",
  "radicalDisplayStyleVerticalGap",
  "radicalRuleThickness",
  "radicalExtraAscender",
  "radicalKernBeforeDegree",
  "radicalKernAfterDegree",
] as const;

export type MathValueConstant = (typeof MATH_VALUE_CONSTANTS)[number];

// Everything the face says about setting mathematics. The value constants are in
// font units like every other measurement here; the four heights and two
// percentages the table states plainly are what they say they are, a percentage
// being a hundredth.
export type MathConstants = Readonly<Record<MathValueConstant, number>> & {
  readonly scriptPercentScaleDown: number;
  readonly scriptScriptPercentScaleDown: number;
  readonly delimitedSubFormulaMinHeight: number;
  readonly displayOperatorMinHeight: number;
  readonly radicalDegreeBottomRaisePercent: number;
};

// One of the shapes a face keeps for a character that grows, resolved so that a
// caller never has to go back to the file: `measurement` is how far it reaches
// along the axis it grows on, which is its height for a parenthesis, and the
// advance and the ink are the glyph's own.
//
// Cambria Math grows a parenthesis through these rather than by stacking pieces:
// measured 2026-08-13, a grown paren came out of Word's pdf as one glyph with
// continuous ink 21.60pt tall.
export type MathVariant = {
  readonly glyph: number;
  readonly measurement: number;
  readonly advance: number;
  readonly ink: InkBox | null;
};

// A piece of a shape assembled out of several, for a character that has to grow
// further than any one of its variants reaches. `extender` is a piece that may be
// repeated to fill what is left.
export type MathAssemblyPart = {
  readonly glyph: number;
  readonly startConnector: number;
  readonly endConnector: number;
  readonly fullAdvance: number;
  readonly extender: boolean;
  readonly advance: number;
  readonly ink: InkBox | null;
};

export type MathAssembly = {
  readonly italicCorrection: number;
  readonly parts: readonly MathAssemblyPart[];
};

export type MathUnavailable = AdvancesUnavailable | "math-missing" | "math-malformed";

export type MathTable =
  | {
      readonly kind: "math";
      readonly constants: MathConstants;
      // Zero where the face states none, since a correction nobody stated moves
      // nothing.
      readonly italicCorrectionOf: (codePoint: number) => number;
      // In the order the face keeps them, which is smallest first and usually
      // starts with the character's own plain glyph.
      readonly tallerVariantsOf: (codePoint: number) => readonly MathVariant[];
      readonly widerVariantsOf: (codePoint: number) => readonly MathVariant[];
      readonly piecesToGrowTaller: (codePoint: number) => MathAssembly | null;
      readonly piecesToGrowWider: (codePoint: number) => MathAssembly | null;
      // How far two pieces of an assembly must overlap, in font units.
      readonly minConnectorOverlap: number;
    }
  | { readonly kind: "unavailable"; readonly reason: MathUnavailable };

export const NO_MATH: MathTable = { kind: "unavailable", reason: "unsupplied" };

// Bold and italic are separate files with their own advances, so a face is asked
// for by style as well as by name.
export type FaceRequest = {
  readonly name: string;
  readonly bold: boolean;
  readonly italic: boolean;
};

// What a run states about kerning and what it is set in, which is the whole of
// what decides whether its pairs move. `w:kern` states the smallest size that
// kerns, in half-points as Word states every size, and null is a run whose
// cascade states none at all.
export type RunKerning = FaceRequest & {
  readonly kernFromHalfPoints: number | null;
  readonly sizePt: number;
};

const HALF_POINTS_IN_A_POINT = 2;

/**
 * Whether a run kerns at all.
 *
 * **Kerning is opt-in.** Measured on 2026-08-13 off Word's own pdf, a right
 * aligned line of kerning pairs in Calibri 12pt: it starts at 427.95 where the
 * run states nothing and at that same 427.95 where it states `w:kern w:val="0"`,
 * against 432.66 where it states `w:kern w:val="1"`. So a run that says nothing
 * is a run that does not kern, and so is one that says zero.
 *
 * **The size the threshold names is a size that kerns.** Measured the same day
 * against `w:kern w:val="32"`: at 15.5pt the line starts at 384.77, which is where
 * that size unkerned starts, and at 16pt it starts at 384.88, which is where that
 * size kerned starts. The comparison is at or above, not above.
 */
export function runKerns(run: RunKerning): boolean {
  const from = run.kernFromHalfPoints;
  if (from === null || from <= 0) return false;
  return run.sizePt * HALF_POINTS_IN_A_POINT >= from;
}

/**
 * Whether the last character of one run and the first of the next may move
 * together.
 *
 * **A pair crosses a run boundary only where the runs are set alike.** Measured
 * on 2026-08-13: `WAVY` written as one run and as `WA` beside `VY` was drawn
 * identically, from 546.79 and 29.20 wide both times, so a boundary is nothing on
 * its own. Where the second run is bold its `V` starts at 562.67 and the first
 * run ends at 562.68, without the half point an `AV` pulls, and a second run set
 * larger is drawn apart the same way.
 *
 * Both runs have to kern. Which of the two marks rules where only one of them
 * states `w:kern` is unmeasured, and this is the reading that moves nothing
 * nobody asked to move.
 */
export function runsKernAcross(before: RunKerning, after: RunKerning): boolean {
  return (
    runKerns(before) &&
    runKerns(after) &&
    normalise(before.name) === normalise(after.name) &&
    before.bold === after.bold &&
    before.italic === after.italic &&
    before.sizePt === after.sizePt
  );
}

export type SuppliedFace = FaceRequest & {
  readonly metrics: FontMetrics;
  readonly advances: AdvanceTable;
  // Absent where nothing asked the file for its pairs, which is not the same as a
  // face that states none: a caller that never read them is told nothing rather
  // than told there is nothing. What each glyph draws and what the face says about
  // setting mathematics travel the same way.
  readonly kerning?: KerningTable;
  readonly ink?: InkTable;
  readonly math?: MathTable;
  // Whether the face draws its letters without serifs, which decides the face Word
  // borrows a character from where this one has no glyph for it: a sans face
  // borrows from Arial and every other face from Times New Roman, measured on
  // 2026-08-06. A face that does not say is not a sans one.
  readonly sansSerif?: boolean;
};

// A character drawn out of a face the run never named, and the metrics of the face
// that drew it: Word measures the line over that face as well as over the stated
// one, so the width is no use without them.
export type BorrowedGlyph = {
  readonly metrics: FontMetrics;
  readonly advance: number;
};

// What draws a character the face itself has no glyph for. Which face answers
// turns on the character as well as on the face that asked, so the two come back
// together. `lookupFontMetrics` never offers one: only a resolver holding every
// face the machine has can say, and `substitutingMetrics` is that resolver.
export type FaceElsewhere = (codePoint: number) => BorrowedGlyph | null;

export type MetricsLookup =
  | {
      readonly kind: "found";
      readonly source: "builtin" | "supplied";
      readonly metrics: FontMetrics;
      readonly advances: AdvanceTable;
      readonly kerning?: KerningTable;
      readonly ink?: InkTable;
      readonly math?: MathTable;
      readonly elsewhere?: FaceElsewhere;
    }
  | { readonly kind: "missing"; readonly fontName: string };

export const lineHeightPt = (metrics: FontMetrics, fontSizePt: number): number =>
  (fontSizePt * (metrics.ascender - metrics.descender + metrics.lineGap)) / metrics.unitsPerEm;

// A face's line gap leads the line from above, so the baseline sits below it as
// well as below the ascender.
export const ascentPt = (metrics: FontMetrics, fontSizePt: number): number =>
  (fontSizePt * (metrics.ascender + metrics.lineGap)) / metrics.unitsPerEm;

export const advanceWidthPt = (advance: number, metrics: FontMetrics, fontSizePt: number): number =>
  (fontSizePt * advance) / metrics.unitsPerEm;

// How tall a line is in a face this project has met, whether or not the machine
// holds one. **A face here states no advances**, so a document written in one still
// has another face stood in to measure its words; only the height comes from here,
// which is the half a stand-in gets most wrong and the half that moves every line
// under it rather than one line.
//
// **Every number is a face's own `hhea`, read rather than typed.** The first ten were
// read off this machine's own copies on 2026-08-14, and each agrees with what Word
// embedded in its own pdfs of the corpus: Arial in 695 of them, Verdana in 482,
// Calibri in 284, Courier New in 102 and Tahoma in 3.
//
// **`times new roman` stated a line gap of 87 until that day**, which is 0.51pt a
// line at 12pt. What Word drew with states none: `TimesNewRomanPSMT` reads the same
// in every one of the 151 pdfs it is embedded in, and so does the copy under
// `/Library/Fonts` here. **This machine does hold a copy stating 87**, and the face
// set reaches for it: `Times New Roman` resolves to 87 while `Times New Roman Bold`
// resolves to 0, so one document's regular and bold text is laid out at two line
// heights. That is a fault in the set of files rather than in this table, and it is
// unfixed; what this row answers is the machine that holds no copy at all, and for
// that one the face Word drew with is the only answer worth giving.
//
// The last two are faces no machine here holds at all, and their numbers come out of
// the font programs Word embedded in its own pdfs of the documents that name them,
// which for an uninstalled face is the only primary source there is: a subsetter
// rewrites the outlines and the character map and copies `hhea` whole. `SegoeUI`
// reads the same across 31 of those pdfs and `ArialNova` across 20.
//
// A style is not asked for here. Every family met so far states one set of vertical
// metrics for all of its cuts: Aptos over six of them, Arial over four, Calibri over
// six and Times New Roman over four all read the same.
const BUILTIN: ReadonlyMap<string, FontMetrics> = new Map([
  ["arial", { unitsPerEm: 2048, ascender: 1854, descender: -434, lineGap: 67 }],
  ["calibri", { unitsPerEm: 2048, ascender: 1950, descender: -550, lineGap: 0 }],
  ["times new roman", { unitsPerEm: 2048, ascender: 1825, descender: -443, lineGap: 0 }],
  ["courier new", { unitsPerEm: 2048, ascender: 1705, descender: -615, lineGap: 0 }],
  ["georgia", { unitsPerEm: 2048, ascender: 1878, descender: -449, lineGap: 0 }],
  ["verdana", { unitsPerEm: 2048, ascender: 2059, descender: -430, lineGap: 0 }],
  ["trebuchet ms", { unitsPerEm: 2048, ascender: 1923, descender: -455, lineGap: 0 }],
  ["tahoma", { unitsPerEm: 2048, ascender: 2049, descender: -423, lineGap: 0 }],
  ["comic sans ms", { unitsPerEm: 2048, ascender: 2257, descender: -597, lineGap: 0 }],
  ["impact", { unitsPerEm: 2048, ascender: 2066, descender: -432, lineGap: 0 }],
  ["segoe ui", { unitsPerEm: 2048, ascender: 2210, descender: -514, lineGap: 0 }],
  ["arial nova", { unitsPerEm: 2048, ascender: 2011, descender: -466, lineGap: 0 }],
]);

const normalise = (fontName: string): string => fontName.trim().toLowerCase();

const sameStyle = (face: SuppliedFace, request: FaceRequest): boolean =>
  face.bold === request.bold && face.italic === request.italic;

export function lookupFontMetrics(
  request: FaceRequest,
  supplied: readonly SuppliedFace[] = [],
): MetricsLookup {
  const key = normalise(request.name);
  const named = supplied.filter((face) => normalise(face.name) === key);
  const exact = named.find((face) => sameStyle(face, request));

  if (exact !== undefined) {
    return {
      kind: "found",
      source: "supplied",
      metrics: exact.metrics,
      advances: exact.advances,
      ...(exact.kerning === undefined ? {} : { kerning: exact.kerning }),
      ...(exact.ink === undefined ? {} : { ink: exact.ink }),
      ...(exact.math === undefined ? {} : { math: exact.math }),
    };
  }

  // A family's styles share their vertical metrics in practice, so a near miss
  // still places paragraphs; it just cannot measure text, since the widths it
  // would use are the wrong style's.
  const nearest = named.find((face) => !face.bold && !face.italic) ?? named[0];
  if (nearest !== undefined) {
    return {
      kind: "found",
      source: "supplied",
      metrics: nearest.metrics,
      advances: { kind: "unavailable", reason: "style-unsupplied" },
      // The pairs of the near miss are as much the wrong style's as its widths
      // are, so they are refused for the same reason, and so is the ink each of
      // its glyphs draws.
      ...(nearest.kerning === undefined
        ? {}
        : { kerning: { kind: "unavailable", reason: "style-unsupplied" } as const }),
      ...(nearest.ink === undefined
        ? {}
        : { ink: { kind: "unavailable", reason: "style-unsupplied" } as const }),
      // The MATH table goes the same way, and for a reason of its own beyond the
      // constants: its corrections and its variants are named in the glyphs of the
      // cut that stated them, and the near miss is another cut.
      ...(nearest.math === undefined
        ? {}
        : { math: { kind: "unavailable", reason: "style-unsupplied" } as const }),
    };
  }

  const builtin = BUILTIN.get(key);
  if (builtin !== undefined) {
    return { kind: "found", source: "builtin", metrics: builtin, advances: NO_ADVANCES };
  }

  return { kind: "missing", fontName: request.name };
}
