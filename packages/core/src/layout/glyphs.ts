import type { AdvanceTable } from "./font-metrics.js";

const CMAP = "cmap";
const HMTX = "hmtx";

type CodeToGlyph = (codePoint: number) => number;

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

// A symbol face has no Unicode subtable at all; its glyphs are reachable only
// through the private use page its own subtable maps.
const SYMBOL_ENCODING = 0;

type Subtable = {
  readonly bytes: Uint8Array;
  readonly symbol: boolean;
};

// Full-repertoire Unicode subtables first, so a face that maps beyond the basic
// plane is read at its widest rather than through its legacy subtable. The symbol
// subtable comes last, since a face that has one usually has nothing else.
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
    const score = preferenceOf(platform, encoding);
    const offset = view.getUint32(record + 4);
    if (score > bestScore && offset + 4 <= cmap.byteLength) {
      bestScore = score;
      bestOffset = offset;
      symbol = platform === 3 && encoding === SYMBOL_ENCODING;
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

export function readAdvanceTable(
  tables: ReadonlyMap<string, Uint8Array>,
  numberOfHMetrics: number,
): AdvanceTable {
  const cmap = tables.get(CMAP);
  if (cmap === undefined) return { kind: "unavailable", reason: "cmap-missing" };

  const hmtx = tables.get(HMTX);
  if (hmtx === undefined) return { kind: "unavailable", reason: "hmtx-missing" };

  const subtable = selectSubtable(cmap);
  if (subtable === null) return { kind: "unavailable", reason: "cmap-unsupported" };

  const parsed = parseSubtable(subtable.bytes);
  if (parsed.kind === "unsupported") return { kind: "unavailable", reason: "cmap-unsupported" };
  if (parsed.kind === "malformed") return { kind: "unavailable", reason: "cmap-malformed" };
  const glyphFor = subtable.symbol ? throughSymbolPage(parsed.glyphFor) : parsed.glyphFor;

  // Glyphs past the last long metric all repeat it, which is how a face gives its
  // monospaced tail a single advance.
  const long = Math.min(numberOfHMetrics, Math.floor(hmtx.byteLength / 4));
  if (long < 1) return { kind: "unavailable", reason: "hmtx-malformed" };

  const view = viewOf(hmtx);
  return {
    kind: "advances",
    advanceFor: (codePoint) => {
      const glyph = glyphFor(codePoint);
      return glyph === 0 ? null : view.getUint16(Math.min(glyph, long - 1) * 4);
    },
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
