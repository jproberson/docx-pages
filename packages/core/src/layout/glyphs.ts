import type { AdvanceTable, AdvancesUnavailable } from "./font-metrics.js";

const CMAP = "cmap";
const HMTX = "hmtx";

export type CodeToGlyph = (codePoint: number) => number;

// Which glyph a face draws a character with, and whether the face is a symbol one,
// which is what decides its answer for a character it has no glyph for.
//
// Layout reads this to reach the character's advance; a writer embedding the face
// reads the number itself, since an Identity-H encoding writes glyphs where the
// document writes characters. One cmap, read once, so the two cannot disagree
// about which glyph a character is.
export type GlyphIndex =
  | { readonly kind: "glyphs"; readonly glyphFor: CodeToGlyph; readonly symbol: boolean }
  | { readonly kind: "unavailable"; readonly reason: AdvancesUnavailable };

type CharacterMap =
  | { readonly kind: "map"; readonly glyphFor: CodeToGlyph }
  | { readonly kind: "unsupported" }
  | { readonly kind: "malformed" };

type Segment = {
  readonly start: number;
  readonly end: number;
  readonly delta: number;
  readonly rangeOffset: number;
  readonly rangeOffsetAt: number;
};

type Group = {
  readonly start: number;
  readonly end: number;
  readonly glyph: number;
};

const viewOf = (bytes: Uint8Array): DataView =>
  new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

// A symbol face declares this encoding to say its glyphs are addressed by byte,
// through the private use page rather than by what its characters mean. Declaring
// it is what makes a face one, whether or not the subtable under it is what is
// read here.
const SYMBOL_ENCODING = 0;

type Subtable = {
  readonly bytes: Uint8Array;
  readonly symbol: boolean;
};

// Full-repertoire Unicode subtables first, so a face that maps beyond the basic
// plane is read at its widest rather than through its legacy subtable. The symbol
// subtable comes last: Symbol carries a Unicode one beside it that maps its own
// page and the Greek and the maths its glyphs stand for besides.
function preferenceOf(platform: number, encoding: number): number {
  if (platform === 3 && encoding === 10) return 5;
  if (platform === 0 && encoding >= 4) return 4;
  if (platform === 3 && encoding === 1) return 3;
  if (platform === 0) return 2;
  if (platform === 3 && encoding === SYMBOL_ENCODING) return 1;
  return 0;
}

function selectSubtable(cmap: Uint8Array): Subtable | null {
  if (cmap.byteLength < 4) return null;
  const view = viewOf(cmap);
  const count = view.getUint16(2);

  let bestScore = 0;
  let bestOffset = 0;
  let symbol = false;
  for (let index = 0; index < count; index += 1) {
    const record = 4 + index * 8;
    if (record + 8 > cmap.byteLength) break;
    const platform = view.getUint16(record);
    const encoding = view.getUint16(record + 2);
    if (platform === 3 && encoding === SYMBOL_ENCODING) symbol = true;
    const score = preferenceOf(platform, encoding);
    const offset = view.getUint32(record + 4);
    if (score > bestScore && offset + 4 <= cmap.byteLength) {
      bestScore = score;
      bestOffset = offset;
    }
  }

  return bestScore === 0 ? null : { bytes: cmap.subarray(bestOffset), symbol };
}

// Word writes a symbol face's characters in the F020 to F0FF page, and a face may
// map them either there or in the low byte they shadow. Each spelling stands for
// the other glyph, so a miss is worth retrying at the one the face did not use.
function symbolAlias(codePoint: number): number | null {
  if (codePoint >= 0x20 && codePoint <= 0xff) return 0xf000 + codePoint;
  if (codePoint >= 0xf020 && codePoint <= 0xf0ff) return codePoint - 0xf000;
  return null;
}

// The narrow no-break space is drawn out of the face's own thin space where the
// face maps no glyph of its own for it. Measured on 2026-08-06 off Word's pdf of
// the authored `unmapped-characters` document: Word drew U+202F in Arial from
// Arial, embedding one glyph whose advance is 410 units, and the one glyph in the
// whole face carrying that advance is U+2009.
const NARROW_NO_BREAK_SPACE = 0x202f;
const THIN_SPACE = 0x2009;

function throughThinSpace(glyphFor: CodeToGlyph): CodeToGlyph {
  return (codePoint) => {
    const direct = glyphFor(codePoint);
    if (direct !== 0 || codePoint !== NARROW_NO_BREAK_SPACE) return direct;
    return glyphFor(THIN_SPACE);
  };
}

// A symbol face answers for every character its own page has a place for, drawing
// .notdef where it has no glyph there rather than letting the character go
// elsewhere. Measured on 2026-08-06: Word drew U+00A0 in Symbol from Symbol, at
// Symbol's .notdef advance of 1229 units, where the face maps neither U+00A0 nor
// the F0A0 it stands for. A character with no place in that page is drawn from
// another face altogether, which is a question no one file can answer: it comes
// back unmapped here and `substitutingMetrics` reaches for the face Word does.
const NOTDEF_GLYPH = 0;

function parseFormat4(table: Uint8Array): CharacterMap {
  if (table.byteLength < 16) return { kind: "malformed" };
  const view = viewOf(table);

  const segCount = view.getUint16(6) / 2;
  const endAt = 14;
  const startAt = endAt + segCount * 2 + 2;
  const deltaAt = startAt + segCount * 2;
  const rangeAt = deltaAt + segCount * 2;
  if (!Number.isInteger(segCount) || segCount < 1 || rangeAt + segCount * 2 > table.byteLength) {
    return { kind: "malformed" };
  }

  const segments: Segment[] = [];
  for (let index = 0; index < segCount; index += 1) {
    segments.push({
      start: view.getUint16(startAt + index * 2),
      end: view.getUint16(endAt + index * 2),
      delta: view.getUint16(deltaAt + index * 2),
      rangeOffset: view.getUint16(rangeAt + index * 2),
      rangeOffsetAt: rangeAt + index * 2,
    });
  }

  return {
    kind: "map",
    glyphFor: (codePoint) => {
      if (codePoint > 0xffff) return 0;
      for (const segment of segments) {
        if (codePoint > segment.end) continue;
        if (codePoint < segment.start) return 0;
        if (segment.rangeOffset === 0) return (codePoint + segment.delta) & 0xffff;

        const at = segment.rangeOffsetAt + segment.rangeOffset + (codePoint - segment.start) * 2;
        if (at + 2 > table.byteLength) return 0;
        const glyph = view.getUint16(at);
        return glyph === 0 ? 0 : (glyph + segment.delta) & 0xffff;
      }
      return 0;
    },
  };
}

function parseFormat12(table: Uint8Array): CharacterMap {
  if (table.byteLength < 16) return { kind: "malformed" };
  const view = viewOf(table);

  const count = view.getUint32(12);
  if (16 + count * 12 > table.byteLength) return { kind: "malformed" };

  const groups: Group[] = [];
  for (let index = 0; index < count; index += 1) {
    const at = 16 + index * 12;
    groups.push({
      start: view.getUint32(at),
      end: view.getUint32(at + 4),
      glyph: view.getUint32(at + 8),
    });
  }

  return {
    kind: "map",
    glyphFor: (codePoint) => {
      for (const group of groups) {
        if (codePoint > group.end) continue;
        return codePoint < group.start ? 0 : group.glyph + (codePoint - group.start);
      }
      return 0;
    },
  };
}

function parseSubtable(table: Uint8Array): CharacterMap {
  const format = viewOf(table).getUint16(0);
  if (format === 4) return parseFormat4(table);
  if (format === 12) return parseFormat12(table);
  return { kind: "unsupported" };
}

export function readGlyphIndex(tables: ReadonlyMap<string, Uint8Array>): GlyphIndex {
  const cmap = tables.get(CMAP);
  if (cmap === undefined) return { kind: "unavailable", reason: "cmap-missing" };

  const subtable = selectSubtable(cmap);
  if (subtable === null) return { kind: "unavailable", reason: "cmap-unsupported" };

  const parsed = parseSubtable(subtable.bytes);
  if (parsed.kind === "unsupported") return { kind: "unavailable", reason: "cmap-unsupported" };
  if (parsed.kind === "malformed") return { kind: "unavailable", reason: "cmap-malformed" };

  return {
    kind: "glyphs",
    symbol: subtable.symbol,
    glyphFor: throughThinSpace(
      subtable.symbol ? throughSymbolPage(parsed.glyphFor) : parsed.glyphFor,
    ),
  };
}

export function readAdvanceTable(
  tables: ReadonlyMap<string, Uint8Array>,
  numberOfHMetrics: number,
): AdvanceTable {
  const index = readGlyphIndex(tables);
  if (index.kind === "unavailable") return index;

  const hmtx = tables.get(HMTX);
  if (hmtx === undefined) return { kind: "unavailable", reason: "hmtx-missing" };

  const { glyphFor, symbol } = index;

  // Glyphs past the last long metric all repeat it, which is how a face gives its
  // monospaced tail a single advance.
  const long = Math.min(numberOfHMetrics, Math.floor(hmtx.byteLength / 4));
  if (long < 1) return { kind: "unavailable", reason: "hmtx-malformed" };

  const view = viewOf(hmtx);
  const advanceOf = (glyph: number): number => view.getUint16(Math.min(glyph, long - 1) * 4);

  return {
    kind: "advances",
    advanceFor: (codePoint) => {
      const glyph = glyphFor(codePoint);
      if (glyph !== 0) return advanceOf(glyph);
      return symbol && symbolAlias(codePoint) !== null ? advanceOf(NOTDEF_GLYPH) : null;
    },
    notDefAdvance: advanceOf(NOTDEF_GLYPH),
  };
}

function throughSymbolPage(glyphFor: CodeToGlyph): CodeToGlyph {
  return (codePoint) => {
    const direct = glyphFor(codePoint);
    if (direct !== 0) return direct;
    const alias = symbolAlias(codePoint);
    return alias === null ? 0 : glyphFor(alias);
  };
}
