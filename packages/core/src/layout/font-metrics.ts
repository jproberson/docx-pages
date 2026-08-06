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

// Bold and italic are separate files with their own advances, so a face is asked
// for by style as well as by name.
export type FaceRequest = {
  readonly name: string;
  readonly bold: boolean;
  readonly italic: boolean;
};

export type SuppliedFace = FaceRequest & {
  readonly metrics: FontMetrics;
  readonly advances: AdvanceTable;
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
    return { kind: "found", source: "supplied", metrics: exact.metrics, advances: exact.advances };
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
    };
  }

  const builtin = BUILTIN.get(key);
  if (builtin !== undefined) {
    return { kind: "found", source: "builtin", metrics: builtin, advances: NO_ADVANCES };
  }

  return { kind: "missing", fontName: request.name };
}
