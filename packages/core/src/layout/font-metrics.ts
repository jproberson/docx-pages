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
  // than told there is nothing.
  readonly kerning?: KerningTable;
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

const BUILTIN: ReadonlyMap<string, FontMetrics> = new Map([
  ["arial", { unitsPerEm: 2048, ascender: 1854, descender: -434, lineGap: 67 }],
  ["calibri", { unitsPerEm: 2048, ascender: 1950, descender: -550, lineGap: 0 }],
  ["times new roman", { unitsPerEm: 2048, ascender: 1825, descender: -443, lineGap: 87 }],
  ["courier new", { unitsPerEm: 2048, ascender: 1705, descender: -615, lineGap: 0 }],
  ["georgia", { unitsPerEm: 2048, ascender: 1878, descender: -449, lineGap: 0 }],
  ["verdana", { unitsPerEm: 2048, ascender: 2059, descender: -430, lineGap: 0 }],
  ["trebuchet ms", { unitsPerEm: 2048, ascender: 1923, descender: -455, lineGap: 0 }],
  ["tahoma", { unitsPerEm: 2048, ascender: 2049, descender: -423, lineGap: 0 }],
  ["comic sans ms", { unitsPerEm: 2048, ascender: 2257, descender: -597, lineGap: 0 }],
  ["impact", { unitsPerEm: 2048, ascender: 2066, descender: -432, lineGap: 0 }],
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
      // are, so they are refused for the same reason.
      ...(nearest.kerning === undefined
        ? {}
        : { kerning: { kind: "unavailable", reason: "style-unsupplied" } as const }),
    };
  }

  const builtin = BUILTIN.get(key);
  if (builtin !== undefined) {
    return { kind: "found", source: "builtin", metrics: builtin, advances: NO_ADVANCES };
  }

  return { kind: "missing", fontName: request.name };
}
