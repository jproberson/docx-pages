import { unzlibSync } from "fflate";

import { DocxPagesError } from "../errors.js";
import type { AdvanceTable, FontMetrics, KerningSource, KerningTable } from "./font-metrics.js";
import { readAdvanceTable, readGlyphIndex as glyphIndexIn, type CodeToGlyph } from "./glyphs.js";

export type FontFileFormat = "sfnt" | "woff";

export type ReadFontMetricsResult = {
  readonly format: FontFileFormat;
  readonly metrics: FontMetrics;
};

// Where the face puts the line under its own letters, in the face's own units.
// `position` is the top of the line measured down from the baseline, so it is
// positive below one.
//
// Word draws an underline where the face says to rather than at a place of its
// own. Measured on 2026-08-07 off Word's own pdf of a reference document, which
// drew every underline 0.1207 em below the baseline and 0.0690 em thick,
// consistently across three runs at 13.92pt, and those are not the ratios of any
// face this machine could stand in: they are the drawn face's own.
export type UnderlineMetrics = {
  readonly position: number;
  readonly thickness: number;
};

// What one face in a font file calls itself. A file may hold several.
export type FontFaceName = {
  readonly family: string;
  readonly fullName: string;
  readonly bold: boolean;
  readonly italic: boolean;
};

export type ReadFontFileResult = ReadFontMetricsResult & {
  readonly advances: AdvanceTable;
  // Whether the face draws its letters without serifs, which is the half of the
  // question the file itself answers about which face Word borrows a character
  // from. See `sansSerif` in `font-metrics.ts` for what turns on it.
  readonly sansSerif: boolean;
  // What each pair of the face's characters moves beyond its advances, or why the
  // face's pairs could not be read. A face stating none is `unkerned` rather than
  // an error: most of them state none.
  readonly kerning: KerningTable;
  // Null where the face states no `post` table, which nothing can be invented
  // for: a renderer that needs a line has to say what it did instead.
  readonly underline: UnderlineMetrics | null;
  // How far the face's letters lean, in degrees, negative to the right. Zero for
  // an upright face and for one that does not say.
  readonly italicAngle: number;
};

const AT = "core/layout/font-file.readFontMetrics";

const GPOS = "GPOS";
const HEAD = "head";
const HHEA = "hhea";
const KERN = "kern";
const NAME = "name";
const OS2 = "OS/2";
const POST = "post";

const tagAt = (bytes: Uint8Array, offset: number): string =>
  String.fromCharCode(...bytes.subarray(offset, offset + 4));

const viewOf = (bytes: Uint8Array): DataView =>
  new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

const unreadable = (reason: string, byteLength: number): DocxPagesError =>
  new DocxPagesError({
    code: "font-unreadable",
    message: reason,
    at: AT,
    context: { byteLength },
  });

function sliceTable(bytes: Uint8Array, offset: number, length: number, tag: string): Uint8Array {
  if (offset + length > bytes.byteLength) {
    throw unreadable(`the ${tag} table runs past the end of the file`, bytes.byteLength);
  }
  return bytes.subarray(offset, offset + length);
}

function readSfntTables(bytes: Uint8Array, at = 0): ReadonlyMap<string, Uint8Array> {
  const view = viewOf(bytes);
  const count = view.getUint16(at + 4);
  const tables = new Map<string, Uint8Array>();
  for (let index = 0; index < count; index += 1) {
    const record = at + 12 + index * 16;
    if (record + 16 > bytes.byteLength) break;
    const tag = tagAt(bytes, record);
    const offset = view.getUint32(record + 8);
    const length = view.getUint32(record + 12);
    tables.set(tag, sliceTable(bytes, offset, length, tag));
  }
  return tables;
}

function readWoffTables(bytes: Uint8Array): ReadonlyMap<string, Uint8Array> {
  const view = viewOf(bytes);
  const count = view.getUint16(12);
  const tables = new Map<string, Uint8Array>();
  for (let index = 0; index < count; index += 1) {
    const record = 44 + index * 20;
    if (record + 20 > bytes.byteLength) break;
    const tag = tagAt(bytes, record);
    const offset = view.getUint32(record + 4);
    const compressedLength = view.getUint32(record + 8);
    const originalLength = view.getUint32(record + 12);
    const stored = sliceTable(bytes, offset, compressedLength, tag);
    if (compressedLength === originalLength) {
      tables.set(tag, stored);
      continue;
    }
    try {
      tables.set(tag, unzlibSync(stored));
    } catch (error: unknown) {
      throw new DocxPagesError({
        code: "font-unreadable",
        message: `the ${tag} table could not be inflated`,
        at: AT,
        context: { table: tag, byteLength: bytes.byteLength },
        cause: error,
      });
    }
  }
  return tables;
}

function requireTable(
  tables: ReadonlyMap<string, Uint8Array>,
  tag: string,
  minimumLength: number,
): Uint8Array {
  const table = tables.get(tag);
  if (table === undefined || table.byteLength < minimumLength) {
    throw new DocxPagesError({
      code: "font-table-missing",
      message: `the font has no usable ${tag} table`,
      at: AT,
      context: { table: tag, presentTables: [...tables.keys()].sort() },
    });
  }
  return table;
}

const HHEA_METRIC_COUNT_AT = 34;

// A face's own name inside a collection: 1 is the family it belongs to and 4 the
// whole of what it is called. Both are read, since a face that is a family of one
// answers the same either way.
const FAMILY_NAME = 1;
const FULL_NAME = 4;

// The Unicode and Windows platforms write two bytes to a character; the Macintosh
// platform writes one. No name read here is outside Latin, so the low byte of a
// wide character is the whole of it.
const MACINTOSH_PLATFORM = 1;

function namesOf(table: Uint8Array): ReadonlyMap<number, string> {
  const found = new Map<number, string>();
  if (table.byteLength < 6) return found;
  const view = viewOf(table);
  const count = view.getUint16(2);
  const stringsAt = view.getUint16(4);

  for (let index = 0; index < count; index += 1) {
    const record = 6 + index * 12;
    if (record + 12 > table.byteLength) break;

    const nameId = view.getUint16(record + 6);
    if (nameId !== FAMILY_NAME && nameId !== FULL_NAME) continue;

    const length = view.getUint16(record + 8);
    const at = stringsAt + view.getUint16(record + 10);
    if (at + length > table.byteLength) continue;

    const bytes = table.subarray(at, at + length);
    const wide = view.getUint16(record) !== MACINTOSH_PLATFORM;
    const step = wide ? 2 : 1;
    let name = "";
    for (let byte = wide ? 1 : 0; byte < bytes.byteLength; byte += step) {
      name += String.fromCharCode(bytes[byte] ?? 0);
    }
    // The first record of an id wins, which is the Macintosh one where a face
    // writes both. They say the same thing, and reading either twice is waste.
    if (!found.has(nameId)) found.set(nameId, name);
  }
  return found;
}

// Whether a face says it is bold or italic, which `head` carries in two bits and
// every sfnt has. A face names its weight in its own full name too, but only
// these say so in a way that does not have to be parsed out of English.
const MAC_STYLE_AT = 44;
const BOLD_BIT = 1;
const ITALIC_BIT = 2;

function styleOf(tables: ReadonlyMap<string, Uint8Array>): { bold: boolean; italic: boolean } {
  const head = tables.get(HEAD);
  if (head === undefined || head.byteLength < MAC_STYLE_AT + 2)
    return { bold: false, italic: false };
  const macStyle = viewOf(head).getUint16(MAC_STYLE_AT);
  return { bold: (macStyle & BOLD_BIT) !== 0, italic: (macStyle & ITALIC_BIT) !== 0 };
}

const normalise = (name: string): string => name.trim().toLowerCase();

// A collection holds several faces in one file, each with a table directory of its
// own whose records still count from the start of the file. Word ships Cambria and
// the Yu Gothic family this way and no other, and the face a document wants may
// not be the one the file is named for: Cambria Math, which Word borrows a maths
// letter and a hyphen from, is the second face of `Cambria.ttc`. It is asked for
// by name rather than by index, since where a face sits in the file is the
// business of whoever shipped it.
function collectionOffsets(bytes: Uint8Array): readonly number[] {
  const view = viewOf(bytes);
  const count = bytes.byteLength < 16 ? 0 : view.getUint32(8);
  const offsets: number[] = [];
  for (let index = 0; index < count && 16 + index * 4 <= bytes.byteLength; index += 1) {
    offsets.push(view.getUint32(12 + index * 4));
  }
  return offsets;
}

function faceOfCollection(bytes: Uint8Array, wanted: string | undefined): number {
  const offsets = collectionOffsets(bytes);

  const first = offsets[0];
  if (first === undefined) throw unreadable("the font collection holds no faces", bytes.byteLength);
  if (wanted === undefined) return first;

  const key = normalise(wanted);
  for (const at of offsets) {
    const table = readSfntTables(bytes, at).get(NAME);
    if (table === undefined) continue;
    if ([...namesOf(table).values()].some((each) => normalise(each) === key)) return at;
  }

  throw new DocxPagesError({
    code: "font-face-missing",
    message: `the font collection holds no face called ${wanted}`,
    at: AT,
    context: { faceName: wanted, faceCount: offsets.length },
  });
}

const PANOSE_AT = 32;
const LATIN_TEXT_FAMILY = 2;

// PANOSE classifies a Latin text face's serifs, and its styles from 11 up are the
// sans ones: 11 to 13 are the plain sans cuts, 14 flared and 15 rounded. Measured
// against the four faces the authored `unmapped-in-a-text-face` document states:
// Arial and Verdana say 11, Calibri says 15, and Times New Roman and Cambria and
// Georgia say 2 or 4.
//
// The family type is read as well as the style, since the byte means something
// else entirely under another one: Wingdings is a pictorial face whose style byte
// is 0, and a pictorial face is not a sans one whatever it says there.
const FIRST_SANS_SERIF_STYLE = 11;
const LAST_SANS_SERIF_STYLE = 15;

function readsSansSerif(tables: ReadonlyMap<string, Uint8Array>): boolean {
  const os2 = tables.get(OS2);
  if (os2 === undefined || os2.byteLength < PANOSE_AT + 2) return false;

  const family = os2[PANOSE_AT] ?? 0;
  const serifStyle = os2[PANOSE_AT + 1] ?? 0;
  return (
    family === LATIN_TEXT_FAMILY &&
    serifStyle >= FIRST_SANS_SERIF_STYLE &&
    serifStyle <= LAST_SANS_SERIF_STYLE
  );
}

type OpenedFont = {
  readonly format: FontFileFormat;
  readonly tables: ReadonlyMap<string, Uint8Array>;
};

// What every reader below starts from: the format the file is in, and the tables
// under whichever face of it was asked for.
function openFontFile(bytes: Uint8Array, faceName: string | undefined): OpenedFont {
  if (bytes.byteLength < 12)
    throw unreadable("the file is too short to be a font", bytes.byteLength);

  const signature = tagAt(bytes, 0);
  if (signature === "wOF2") {
    throw new DocxPagesError({
      code: "font-format-unsupported",
      message: "woff2 needs brotli, which is not available in every runtime; supply woff or otf",
      at: AT,
      context: { format: "woff2" },
    });
  }

  const view = viewOf(bytes);
  const version = view.getUint32(0);
  const isSfnt = version === 0x00010000 || signature === "OTTO" || signature === "true";
  const isCollection = signature === "ttcf";
  if (!isSfnt && !isCollection && signature !== "wOFF") {
    throw unreadable("the file is not an sfnt, woff or woff2 font", bytes.byteLength);
  }

  const format: FontFileFormat = signature === "wOFF" ? "woff" : "sfnt";
  return {
    format,
    tables:
      format === "woff"
        ? readWoffTables(bytes)
        : readSfntTables(bytes, isCollection ? faceOfCollection(bytes, faceName) : 0),
  };
}

// The `post` table opens with its version, then the angle the letters lean at as
// a fixed-point number, then where the line under them goes and how thick it is.
const ITALIC_ANGLE_AT = 4;
const UNDERLINE_POSITION_AT = 8;
const UNDERLINE_THICKNESS_AT = 10;

// A fixed-point number in the font formats is a whole part and a fraction of
// sixty-five thousand.
const FIXED = 65536;

const POST_LENGTH = 12;

function readPost(tables: ReadonlyMap<string, Uint8Array>): {
  readonly underline: UnderlineMetrics | null;
  readonly italicAngle: number;
} {
  const post = tables.get(POST);
  if (post === undefined || post.byteLength < POST_LENGTH) {
    return { underline: null, italicAngle: 0 };
  }

  const view = viewOf(post);
  return {
    // Stated as a distance up from the baseline, which for a line under the
    // letters is negative. Turned the right way up here so that a caller adding it
    // to a baseline is going down the page, as everything else here does.
    underline: {
      position: -view.getInt16(UNDERLINE_POSITION_AT),
      thickness: view.getInt16(UNDERLINE_THICKNESS_AT),
    },
    italicAngle: view.getInt32(ITALIC_ANGLE_AT) / FIXED,
  };
}

// What a pair of glyphs moves, in font units, or null where the table read says
// nothing about that pair. Null rather than zero, so that one subtable saying
// nothing lets the next one answer.
type PairMovement = (leftGlyph: number, rightGlyph: number) => number | null;

// What a table of pairs came to. `unstated` is a table that holds no pairs this
// reader is looking for, which is not a fault. `unhonoured` counts the pair
// lookups that state their movement in a way this reader will not guess at, which
// are left unread and said out loud rather than read at half of what they say.
type PairReading =
  | { readonly kind: "pairs"; readonly movementFor: PairMovement; readonly unhonoured: number }
  | { readonly kind: "unstated"; readonly unhonoured: number }
  | { readonly kind: "malformed" };

const UNSTATED: PairReading = { kind: "unstated", unhonoured: 0 };
const UNHONOURED: PairReading = { kind: "unstated", unhonoured: 1 };
const MALFORMED: PairReading = { kind: "malformed" };

const pairKey = (leftGlyph: number, rightGlyph: number): number => leftGlyph * 0x10000 + rightGlyph;

// The legacy table's coverage byte. Only horizontal values that add to an advance
// are read: a minimum subtable bounds a pair rather than moving it, and a
// cross-stream one moves the pair off the line of writing, and neither is a width.
const HORIZONTAL_COVERAGE = 1;
const MINIMUM_COVERAGE = 2;
const CROSS_STREAM_COVERAGE = 4;

const KERN_HEADER = 4;
const KERN_SUBTABLE_HEADER = 6;
const KERN_FORMAT_0_HEADER = 8;
const KERN_PAIR_LENGTH = 6;

// The version every Windows font states. Version 1 is Apple's table, whose header
// and subtables are shaped differently; it is left unread rather than read as if
// it were this one.
const MICROSOFT_KERN_VERSION = 0;

function readLegacyKern(kern: Uint8Array | undefined): PairReading {
  if (kern === undefined) return UNSTATED;
  if (kern.byteLength < KERN_HEADER) return MALFORMED;

  const view = viewOf(kern);
  if (view.getUint16(0) !== MICROSOFT_KERN_VERSION) return UNSTATED;

  const values = new Map<number, number>();
  let at = KERN_HEADER;
  for (let index = 0; index < view.getUint16(2); index += 1) {
    if (at + KERN_SUBTABLE_HEADER > kern.byteLength) return MALFORMED;

    const coverage = view.getUint16(at + 4);
    const format = coverage >> 8;
    const length = subtableLength(kern, at, format);
    if (length === null) return MALFORMED;

    const horizontalAdvances =
      (coverage & HORIZONTAL_COVERAGE) !== 0 &&
      (coverage & (MINIMUM_COVERAGE | CROSS_STREAM_COVERAGE)) === 0;
    if (horizontalAdvances && format === 0) {
      const body = kern.subarray(at + KERN_SUBTABLE_HEADER, at + length);
      if (!addFormat0Pairs(body, values)) return MALFORMED;
    }
    at += length;
  }

  if (values.size === 0) return UNSTATED;
  return {
    kind: "pairs",
    unhonoured: 0,
    movementFor: (left, right) => values.get(pairKey(left, right)) ?? null,
  };
}

// How long a subtable really is, or null where it runs past the table.
//
// **A subtable states its length in two bytes and Word's own faces overflow it.**
// Measured on 2026-08-13: Calibri's `kern` table is 160254 bytes and its one
// subtable states 29178, which is what is left of 160250 after two whole turns of
// the field; Cambria states 47232 where its pairs need 178304. The pair count is
// the field that survives, so where the pairs reach further than the stated length
// and still fit inside the table, they are what is read. Reading the stated length
// instead loses every pair in both faces.
function subtableLength(kern: Uint8Array, at: number, format: number): number | null {
  const view = viewOf(kern);
  const stated = view.getUint16(at + 2);

  let length = stated;
  if (format === 0 && at + KERN_SUBTABLE_HEADER + KERN_FORMAT_0_HEADER <= kern.byteLength) {
    const pairs = view.getUint16(at + KERN_SUBTABLE_HEADER);
    const implied = KERN_SUBTABLE_HEADER + KERN_FORMAT_0_HEADER + pairs * KERN_PAIR_LENGTH;
    if (implied > stated && at + implied <= kern.byteLength) length = implied;
  }

  return length < KERN_SUBTABLE_HEADER || at + length > kern.byteLength ? null : length;
}

// False where the subtable states more pairs than it holds, which is refused
// rather than read as far as it goes.
function addFormat0Pairs(body: Uint8Array, values: Map<number, number>): boolean {
  if (body.byteLength < KERN_FORMAT_0_HEADER) return false;

  const view = viewOf(body);
  const count = view.getUint16(0);
  if (KERN_FORMAT_0_HEADER + count * KERN_PAIR_LENGTH > body.byteLength) return false;

  for (let index = 0; index < count; index += 1) {
    const at = KERN_FORMAT_0_HEADER + index * KERN_PAIR_LENGTH;
    const key = pairKey(view.getUint16(at), view.getUint16(at + 2));
    // A face may state the same pair in more than one subtable, and what those
    // say adds up.
    values.set(key, (values.get(key) ?? 0) + view.getInt16(at + 4));
  }
  return true;
}

const GPOS_HEADER = 10;
const KERN_FEATURE = "kern";
const PAIR_ADJUSTMENT = 2;
const EXTENSION_POSITIONING = 9;
const EXTENSION_LENGTH = 8;
const X_ADVANCE = 0x0004;

// The only movement read: the first glyph's own advance, which is what kerning a
// line of text is. A subtable that also places a glyph, or that moves the second
// one, is left unread rather than read at half of what it says.
//
// **A face is not refused whole over one of those.** Measured on 2026-08-13 over
// the 472 faces on this machine: the subtables stating anything else place the
// glyph by exactly what they advance it, which is how a pair is stated for text
// running the other way, and not one of them covers a Latin glyph. Calibri states
// its Latin pairs the readable way and eight subtables the other, all eight of
// them Latin-free; Liberation Sans covers 105 glyphs the readable way and 49
// Hebrew ones the other. Refusing the face over them loses every Latin pair in
// Calibri, which is the face this repository's own documents are written in. So a
// subtable is what is left unread, and how many were is carried out beside the
// pairs.
const onlyFirstAdvance = (first: number, second: number): boolean =>
  (first === 0 || first === X_ADVANCE) && second === 0;

function readGposPairs(gpos: Uint8Array | undefined): PairReading {
  if (gpos === undefined) return UNSTATED;
  if (gpos.byteLength < GPOS_HEADER) return MALFORMED;

  const view = viewOf(gpos);
  const wanted = kernFeatureLookups(gpos, view.getUint16(6));
  if (wanted === null) return MALFORMED;
  if (wanted.length === 0) return UNSTATED;

  const lookupList = view.getUint16(8);
  const movements: PairMovement[] = [];
  let unhonoured = 0;
  for (const index of wanted) {
    const reading = readPairLookup(gpos, lookupList, index);
    if (reading.kind === "malformed") return reading;
    unhonoured += reading.unhonoured;
    if (reading.kind === "pairs") movements.push(reading.movementFor);
  }

  if (movements.length === 0) return { kind: "unstated", unhonoured };
  return { kind: "pairs", movementFor: added(movements), unhonoured };
}

// What every lookup of the pair says, added up, since each is applied in its turn
// over the same two glyphs.
function added(movements: readonly PairMovement[]): PairMovement {
  return (left, right) => {
    let moved: number | null = null;
    for (const movementFor of movements) {
      const value = movementFor(left, right);
      if (value !== null) moved = (moved ?? 0) + value;
    }
    return moved;
  };
}

// Which lookups the `kern` feature points at. The script list is not read: a face
// states its pairs once and every script that kerns points at the same lookups,
// and which script a document's text is in is a question the layout does not ask.
function kernFeatureLookups(gpos: Uint8Array, featureListAt: number): readonly number[] | null {
  if (featureListAt + 2 > gpos.byteLength) return null;

  const view = viewOf(gpos);
  const count = view.getUint16(featureListAt);
  if (featureListAt + 2 + count * 6 > gpos.byteLength) return null;

  const wanted = new Set<number>();
  for (let index = 0; index < count; index += 1) {
    const record = featureListAt + 2 + index * 6;
    if (tagAt(gpos, record) !== KERN_FEATURE) continue;

    const feature = featureListAt + view.getUint16(record + 4);
    if (feature + 4 > gpos.byteLength) return null;
    const lookupCount = view.getUint16(feature + 2);
    if (feature + 4 + lookupCount * 2 > gpos.byteLength) return null;

    for (let each = 0; each < lookupCount; each += 1) {
      wanted.add(view.getUint16(feature + 4 + each * 2));
    }
  }
  return [...wanted];
}

function readPairLookup(gpos: Uint8Array, lookupListAt: number, index: number): PairReading {
  if (lookupListAt + 2 > gpos.byteLength) return MALFORMED;

  const view = viewOf(gpos);
  const count = view.getUint16(lookupListAt);
  if (index >= count || lookupListAt + 2 + count * 2 > gpos.byteLength) return MALFORMED;

  const lookup = lookupListAt + view.getUint16(lookupListAt + 2 + index * 2);
  if (lookup + 6 > gpos.byteLength) return MALFORMED;

  const type = view.getUint16(lookup);
  const subtableCount = view.getUint16(lookup + 4);
  if (lookup + 6 + subtableCount * 2 > gpos.byteLength) return MALFORMED;

  // The feature may point at a lookup of another type, which refines what the
  // pairs say rather than stating pairs of its own: contextual positioning is the
  // one real faces state. What it would change is left unread, since the pair
  // value under it is nearer than no movement at all.
  if (type !== PAIR_ADJUSTMENT && type !== EXTENSION_POSITIONING) return UNSTATED;

  const movements: PairMovement[] = [];
  let unhonoured = 0;
  for (let each = 0; each < subtableCount; each += 1) {
    const at = lookup + view.getUint16(lookup + 6 + each * 2);
    const reading =
      type === EXTENSION_POSITIONING
        ? readExtension(gpos, at)
        : readPairSubtable(gpos.subarray(at));
    if (reading.kind === "malformed") return reading;
    unhonoured += reading.unhonoured;
    if (reading.kind === "pairs") movements.push(reading.movementFor);
  }

  if (movements.length === 0) return { kind: "unstated", unhonoured };
  // The subtables of one lookup are tried in turn and the first that covers the
  // pair answers for it, which is how a face states a general rule and an
  // exception to it.
  return {
    kind: "pairs",
    unhonoured,
    movementFor: (left, right) => {
      for (const movementFor of movements) {
        const value = movementFor(left, right);
        if (value !== null) return value;
      }
      return null;
    },
  };
}

// How a face reaches a subtable sitting further into the file than a two-byte
// offset can name. Word ships faces whose pair positioning is only reachable this
// way, so a reader that stops here reads no pairs at all out of them.
function readExtension(gpos: Uint8Array, at: number): PairReading {
  if (at + EXTENSION_LENGTH > gpos.byteLength) return MALFORMED;

  const view = viewOf(gpos);
  if (view.getUint16(at) !== 1) return MALFORMED;
  if (view.getUint16(at + 2) !== PAIR_ADJUSTMENT) return UNSTATED;
  return readPairSubtable(gpos.subarray(at + view.getUint32(at + 4)));
}

function readPairSubtable(table: Uint8Array): PairReading {
  if (table.byteLength < 2) return MALFORMED;

  const format = viewOf(table).getUint16(0);
  if (format === 1) return readPairSets(table);
  if (format === 2) return readClassPairs(table);
  return MALFORMED;
}

const PAIR_SET_HEADER = 10;

// A pair set names the second glyph of every pair the first one kerns with, which
// is how a face states the few hundred pairs it cares about one by one.
function readPairSets(table: Uint8Array): PairReading {
  if (table.byteLength < PAIR_SET_HEADER) return MALFORMED;

  const view = viewOf(table);
  const first = view.getUint16(4);
  const second = view.getUint16(6);
  if (!onlyFirstAdvance(first, second)) return UNHONOURED;

  const covered = readCoverage(table, view.getUint16(2));
  if (covered === null) return MALFORMED;

  const setCount = view.getUint16(8);
  if (PAIR_SET_HEADER + setCount * 2 > table.byteLength) return MALFORMED;
  // A face states one pair set for each glyph its coverage names, so a table where
  // the two disagree is one this reader cannot line up.
  if (covered.glyphs.length !== setCount) return MALFORMED;

  const recordLength = 2 + (first === 0 ? 0 : 2);
  const values = new Map<number, number>();
  for (const [index, glyph] of covered.glyphs.entries()) {
    const set = view.getUint16(PAIR_SET_HEADER + index * 2);
    if (set + 2 > table.byteLength) return MALFORMED;

    const pairCount = view.getUint16(set);
    if (set + 2 + pairCount * recordLength > table.byteLength) return MALFORMED;

    for (let pair = 0; pair < pairCount; pair += 1) {
      const at = set + 2 + pair * recordLength;
      const moved = first === 0 ? 0 : view.getInt16(at + 2);
      values.set(pairKey(glyph, view.getUint16(at)), moved);
    }
  }

  if (values.size === 0) return UNSTATED;
  return {
    kind: "pairs",
    unhonoured: 0,
    movementFor: (left, right) => values.get(pairKey(left, right)) ?? null,
  };
}

const CLASS_PAIR_HEADER = 16;

// Every glyph of one class kerns the same against every glyph of another, which is
// how a face states the pairs of whole alphabets without naming them.
function readClassPairs(table: Uint8Array): PairReading {
  if (table.byteLength < CLASS_PAIR_HEADER) return MALFORMED;

  const view = viewOf(table);
  const first = view.getUint16(4);
  const second = view.getUint16(6);
  if (!onlyFirstAdvance(first, second)) return UNHONOURED;
  if (first === 0) return UNSTATED;

  const covered = readCoverage(table, view.getUint16(2));
  const firstClassOf = readClassDefinition(table, view.getUint16(8));
  const secondClassOf = readClassDefinition(table, view.getUint16(10));
  if (covered === null || firstClassOf === null || secondClassOf === null) return MALFORMED;

  const firstCount = view.getUint16(12);
  const secondCount = view.getUint16(14);
  if (CLASS_PAIR_HEADER + firstCount * secondCount * 2 > table.byteLength) return MALFORMED;

  return {
    kind: "pairs",
    unhonoured: 0,
    movementFor: (left, right) => {
      if (covered.indexOf(left) === null) return null;
      const firstClass = firstClassOf(left);
      const secondClass = secondClassOf(right);
      if (firstClass >= firstCount || secondClass >= secondCount) return null;
      return view.getInt16(CLASS_PAIR_HEADER + (firstClass * secondCount + secondClass) * 2);
    },
  };
}

type Coverage = {
  readonly indexOf: (glyph: number) => number | null;
  readonly glyphs: readonly number[];
};

// A face has at most this many glyphs, so a coverage naming more of them is one
// whose ranges are not what they claim.
const GLYPH_LIMIT = 0x10000;

function readCoverage(table: Uint8Array, at: number): Coverage | null {
  if (at + 4 > table.byteLength) return null;

  const view = viewOf(table);
  const format = view.getUint16(at);
  const count = view.getUint16(at + 2);
  const glyphs: number[] = [];

  if (format === 1) {
    if (at + 4 + count * 2 > table.byteLength) return null;
    for (let index = 0; index < count; index += 1) glyphs.push(view.getUint16(at + 4 + index * 2));
  } else if (format === 2) {
    if (at + 4 + count * 6 > table.byteLength) return null;
    for (let index = 0; index < count; index += 1) {
      const record = at + 4 + index * 6;
      const start = view.getUint16(record);
      const end = view.getUint16(record + 2);
      if (end < start || glyphs.length + (end - start) >= GLYPH_LIMIT) return null;
      for (let glyph = start; glyph <= end; glyph += 1) glyphs.push(glyph);
    }
  } else return null;

  const indices = new Map(glyphs.map((glyph, index) => [glyph, index]));
  return { glyphs, indexOf: (glyph) => indices.get(glyph) ?? null };
}

// Which class a glyph belongs to. Class 0 is every glyph the table does not name,
// and a face states a movement for it like any other.
type GlyphClass = (glyph: number) => number;

function readClassDefinition(table: Uint8Array, at: number): GlyphClass | null {
  if (at + 4 > table.byteLength) return null;

  const view = viewOf(table);
  const format = view.getUint16(at);

  if (format === 1) {
    if (at + 6 > table.byteLength) return null;
    const start = view.getUint16(at + 2);
    const count = view.getUint16(at + 4);
    if (at + 6 + count * 2 > table.byteLength) return null;
    return (glyph) =>
      glyph < start || glyph >= start + count ? 0 : view.getUint16(at + 6 + (glyph - start) * 2);
  }

  if (format !== 2) return null;
  const count = view.getUint16(at + 2);
  if (at + 4 + count * 6 > table.byteLength) return null;
  return (glyph) => {
    for (let index = 0; index < count; index += 1) {
      const record = at + 4 + index * 6;
      if (glyph < view.getUint16(record)) return 0;
      if (glyph <= view.getUint16(record + 2)) return view.getUint16(record + 4);
    }
    return 0;
  };
}

// Where a face states both tables, this is the one read and the other is left
// alone. GPOS is taken because it is what a shaper on Windows applies: a face
// carrying both states the same pairs in GPOS for anything that shapes text and
// keeps the legacy table for software that does not.
//
// **Which one Word agrees with is unmeasured**, and on the faces here it may not
// matter: Arial states both, and over the 256 pairs of sixteen letters the two
// agree on all 53 that move and disagree on none (measured 2026-08-13). It is one
// constant so that a measurement can move it: `"kern"` reads the legacy table
// wherever a face has one and falls back on GPOS.
const PREFERRED_KERNING_TABLE: KerningSource = "gpos";

function readKerning(tables: ReadonlyMap<string, Uint8Array>, glyphFor: CodeToGlyph): KerningTable {
  const gpos = readGposPairs(tables.get(GPOS));
  const legacy = readLegacyKern(tables.get(KERN));

  // A table that cannot be read at all refuses the face's kerning whole rather than
  // falling back on the other one. A face states its pairs once, so a table nothing
  // can be read out of is a face whose pairs are unknown, and the other table's
  // half of them moves text to a place neither Word nor the face asked for.
  if (gpos.kind === "malformed") return { kind: "unavailable", reason: "gpos-malformed" };
  if (legacy.kind === "malformed") return { kind: "unavailable", reason: "kern-malformed" };

  const inOrder: readonly (readonly [KerningSource, PairReading])[] =
    PREFERRED_KERNING_TABLE === "gpos"
      ? [
          ["gpos", gpos],
          ["kern", legacy],
        ]
      : [
          ["kern", legacy],
          ["gpos", gpos],
        ];

  for (const [source, reading] of inOrder) {
    if (reading.kind !== "pairs") continue;
    const { movementFor } = reading;
    return {
      kind: "kerning",
      source,
      subtablesLeftUnread: reading.unhonoured,
      // Asked in characters, as the advances are, and answered through the same
      // character map, so a pair measured here is the pair the face draws.
      kerningBetween: (leftCodePoint, rightCodePoint) => {
        const left = glyphFor(leftCodePoint);
        const right = glyphFor(rightCodePoint);
        if (left === 0 || right === 0) return 0;
        return movementFor(left, right) ?? 0;
      },
    };
  }

  // A face whose only pair positioning states a movement this reader will not
  // guess at has kerning and no pair of it could be read, which is a different
  // answer from a face that states none.
  if (gpos.unhonoured > 0) return { kind: "unavailable", reason: "gpos-unsupported" };
  return { kind: "unavailable", reason: "unkerned" };
}

/**
 * Which glyph a face draws each character with, read out of the same cmap the
 * advances are read through, so that a character measured at one glyph's width is
 * never written as another.
 *
 * A writer embedding the face needs this and layout does not: a font embedded
 * under Identity-H is addressed by glyph, where the document is written in
 * characters. Unmapped characters answer 0, which is .notdef and is the face's
 * own answer for one it cannot draw.
 *
 * A face whose cmap cannot be read is **refused rather than written wrongly**:
 * there is no glyph number to fall back on that would draw the right letter.
 */
export function readGlyphIndex(bytes: Uint8Array, faceName?: string): CodeToGlyph {
  const { tables } = openFontFile(bytes, faceName);
  const index = glyphIndexIn(tables);
  if (index.kind === "unavailable") {
    throw new DocxPagesError({
      code: "font-glyphs-unreadable",
      message: `the font's character map could not be read: ${index.reason}`,
      at: "core/layout/font-file.readGlyphIndex",
      context: { reason: index.reason, presentTables: [...tables.keys()].sort() },
    });
  }
  return index.glyphFor;
}

/**
 * Reads what laying text out in a face needs: its vertical metrics, the advance of
 * every glyph it maps, and whether it draws its letters without serifs.
 *
 * `faceName` picks one face out of a collection, which holds several; without one
 * the face the file is named for answers, and a file holding one face ignores it.
 */
export function readFontFile(bytes: Uint8Array, faceName?: string): ReadFontFileResult {
  const { format, tables } = openFontFile(bytes, faceName);

  const head = requireTable(tables, HEAD, 20);
  const hhea = requireTable(tables, HHEA, 10);
  const headView = viewOf(head);
  const hheaView = viewOf(hhea);

  const metricCount =
    hhea.byteLength >= HHEA_METRIC_COUNT_AT + 2 ? hheaView.getUint16(HHEA_METRIC_COUNT_AT) : 0;

  const index = glyphIndexIn(tables);

  return {
    format,
    metrics: {
      unitsPerEm: headView.getUint16(18),
      ascender: hheaView.getInt16(4),
      descender: hheaView.getInt16(6),
      lineGap: hheaView.getInt16(8),
    },
    advances: readAdvanceTable(tables, metricCount),
    // A face whose characters cannot be mapped to glyphs has no pairs either: the
    // tables state glyphs, and there is nothing to look them up by.
    kerning:
      index.kind === "glyphs"
        ? readKerning(tables, index.glyphFor)
        : { kind: "unavailable", reason: index.reason },
    sansSerif: readsSansSerif(tables),
    ...readPost(tables),
  };
}

export function readFontMetrics(bytes: Uint8Array): ReadFontMetricsResult {
  const { format, metrics } = readFontFile(bytes);
  return { format, metrics };
}

/**
 * What a font file calls the faces in it, without reading a glyph.
 *
 * A file on disk is named for whoever shipped it and not for what a document asks
 * for: `seguisb.ttf` is Segoe UI Semibold and `calibril.ttf` is Calibri Light, and
 * a collection holds several faces behind one name. So a machine's fonts can only
 * be offered to a layout by opening each one and asking it, which is what this
 * answers. Reading a whole face costs the advance of every glyph in it; this reads
 * two tables.
 *
 * The family is what a document names in `w:rFonts` and the bold and italic are
 * what it asks for beside it. `fullName` is the face's whole name, which is a
 * family of its own to a document that names it there.
 */
export function readFontFaces(bytes: Uint8Array): readonly FontFaceName[] {
  if (bytes.byteLength < 12)
    throw unreadable("the file is too short to be a font", bytes.byteLength);

  const signature = tagAt(bytes, 0);
  if (signature === "ttcf") {
    return collectionOffsets(bytes).map((at) => faceNameOf(readSfntTables(bytes, at)));
  }
  if (signature === "wOFF") return [faceNameOf(readWoffTables(bytes))];
  return [faceNameOf(readSfntTables(bytes, 0))];
}

function faceNameOf(tables: ReadonlyMap<string, Uint8Array>): FontFaceName {
  const names = namesOf(tables.get(NAME) ?? new Uint8Array(0));
  const family = names.get(FAMILY_NAME) ?? "";
  return {
    family,
    fullName: names.get(FULL_NAME) ?? family,
    ...styleOf(tables),
  };
}
