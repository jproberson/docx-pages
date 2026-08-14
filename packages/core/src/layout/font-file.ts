import { unzlibSync } from "fflate";

import { DocxPagesError } from "../errors.js";
import {
  MATH_VALUE_CONSTANTS,
  type AdvanceTable,
  type FontMetrics,
  type InkBox,
  type InkTable,
  type KerningSource,
  type KerningTable,
  type MathAssembly,
  type MathAssemblyPart,
  type MathConstants,
  type MathTable,
  type MathValueConstant,
  type MathVariant,
} from "./font-metrics.js";
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
  // What each of the face's glyphs draws, out of its outlines, and what it says
  // about setting mathematics. A face stating no MATH table is `math-missing`
  // rather than an error: all but a handful state none.
  readonly ink: InkTable;
  readonly math: MathTable;
  // Null where the face states no `post` table, which nothing can be invented
  // for: a renderer that needs a line has to say what it did instead.
  readonly underline: UnderlineMetrics | null;
  // How far the face's letters lean, in degrees, negative to the right. Zero for
  // an upright face and for one that does not say.
  readonly italicAngle: number;
};

const AT = "core/layout/font-file.readFontMetrics";

const CFF = "CFF ";
const GLYF = "glyf";
const GPOS = "GPOS";
const HEAD = "head";
const HHEA = "hhea";
const HMTX = "hmtx";
const KERN = "kern";
const LOCA = "loca";
const MATH = "MATH";
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

// What one glyph draws, or null where it draws nothing at all. A space is written
// as a glyph of no length, so nothing is not a fault.
type GlyphInkOf = (glyph: number) => InkBox | null;

// What stands in for the character map of a face that has none, so that a reader
// below can be written without asking whether there is one.
const NOTHING_MAPPED: CodeToGlyph = () => 0;

type InkReading =
  | { readonly kind: "ink"; readonly inkOfGlyph: GlyphInkOf }
  | { readonly kind: "missing" }
  | { readonly kind: "unsupported" }
  | { readonly kind: "malformed" };

const INK_MISSING: InkReading = { kind: "missing" };
const INK_UNSUPPORTED: InkReading = { kind: "unsupported" };
const INK_MALFORMED: InkReading = { kind: "malformed" };

// Where `head` says how wide the offsets in `loca` are, and how long a glyph's
// own header is: the number of contours it holds and then the four sides of what
// it draws.
const LOCA_FORMAT_AT = 50;
const GLYPH_HEADER_LENGTH = 10;

// A TrueType glyph states the box it draws in its own header, so nothing has to
// follow the outline. A composite glyph states it there too, over all the glyphs
// it is built from.
function readTrueTypeInk(tables: ReadonlyMap<string, Uint8Array>): InkReading {
  const glyf = tables.get(GLYF);
  const loca = tables.get(LOCA);
  if (glyf === undefined || loca === undefined) return INK_MISSING;

  const head = tables.get(HEAD);
  if (head === undefined || head.byteLength < LOCA_FORMAT_AT + 2) return INK_MALFORMED;

  const long = viewOf(head).getInt16(LOCA_FORMAT_AT) === 1;
  const entry = long ? 4 : 2;
  const glyphCount = Math.floor(loca.byteLength / entry) - 1;
  if (glyphCount < 1) return INK_MALFORMED;

  const locaView = viewOf(loca);
  const glyfView = viewOf(glyf);
  // A short `loca` counts in pairs of bytes, which is what lets it reach a table
  // twice as long as its own numbers.
  const offsetAt = (index: number): number =>
    long ? locaView.getUint32(index * 4) : locaView.getUint16(index * 2) * 2;

  return {
    kind: "ink",
    inkOfGlyph: (glyph) => {
      if (glyph < 0 || glyph >= glyphCount) return null;
      const from = offsetAt(glyph);
      const to = offsetAt(glyph + 1);
      if (to <= from || from + GLYPH_HEADER_LENGTH > glyf.byteLength) return null;
      return {
        left: glyfView.getInt16(from + 2),
        bottom: glyfView.getInt16(from + 4),
        right: glyfView.getInt16(from + 6),
        top: glyfView.getInt16(from + 8),
      };
    },
  };
}

type MutableBox = { left: number; right: number; bottom: number; top: number };

function extendedBy(box: MutableBox | null, x: number, y: number): MutableBox {
  if (box === null) return { left: x, right: x, bottom: y, top: y };
  return {
    left: Math.min(box.left, x),
    right: Math.max(box.right, x),
    bottom: Math.min(box.bottom, y),
    top: Math.max(box.top, y),
  };
}

const cubicAt = (from: number, first: number, second: number, to: number, at: number): number => {
  const rest = 1 - at;
  return (
    rest * rest * rest * from +
    3 * rest * rest * at * first +
    3 * rest * at * at * second +
    at * at * at * to
  );
};

// Where a curve turns back on one axis, which is where it reaches furthest. A
// PostScript face states no box of its own, so the box is what the outline comes
// to, and a curve's own ends are not the whole of it: the control points lie
// outside the ink, and taking them for it would draw the letter taller than it is.
function turningPoints(from: number, first: number, second: number, to: number): readonly number[] {
  const square = -from + 3 * first - 3 * second + to;
  const linear = 2 * from - 4 * first + 2 * second;
  const constant = first - from;

  if (Math.abs(square) < 1e-12) {
    if (Math.abs(linear) < 1e-12) return [];
    return [-constant / linear];
  }

  const discriminant = linear * linear - 4 * square * constant;
  if (discriminant < 0) return [];
  const root = Math.sqrt(discriminant);
  return [(-linear + root) / (2 * square), (-linear - root) / (2 * square)];
}

function alongCurve(
  box: MutableBox | null,
  from: readonly [number, number],
  first: readonly [number, number],
  second: readonly [number, number],
  to: readonly [number, number],
): MutableBox {
  let reached = extendedBy(box, to[0], to[1]);

  for (const at of turningPoints(from[0], first[0], second[0], to[0])) {
    if (at <= 0 || at >= 1) continue;
    const x = cubicAt(from[0], first[0], second[0], to[0], at);
    const y = cubicAt(from[1], first[1], second[1], to[1], at);
    reached = extendedBy(reached, x, y);
  }
  for (const at of turningPoints(from[1], first[1], second[1], to[1])) {
    if (at <= 0 || at >= 1) continue;
    const x = cubicAt(from[0], first[0], second[0], to[0], at);
    const y = cubicAt(from[1], first[1], second[1], to[1], at);
    reached = extendedBy(reached, x, y);
  }
  return reached;
}

type CffIndex = {
  readonly count: number;
  readonly entry: (index: number) => Uint8Array | null;
  readonly end: number;
};

function readCffIndex(cff: Uint8Array, at: number): CffIndex | null {
  if (at < 0 || at + 2 > cff.byteLength) return null;

  const count = viewOf(cff).getUint16(at);
  if (count === 0) return { count: 0, entry: () => null, end: at + 2 };

  const offsetSize = cff[at + 2] ?? 0;
  if (offsetSize < 1 || offsetSize > 4) return null;

  const offsetsAt = at + 3;
  const dataAt = offsetsAt + (count + 1) * offsetSize - 1;
  if (dataAt >= cff.byteLength) return null;

  const offsetOf = (index: number): number => {
    let value = 0;
    for (let byte = 0; byte < offsetSize; byte += 1) {
      value = value * 256 + (cff[offsetsAt + index * offsetSize + byte] ?? 0);
    }
    return value;
  };

  const end = dataAt + offsetOf(count);
  if (end > cff.byteLength || offsetOf(0) !== 1) return null;

  return {
    count,
    end,
    entry: (index) =>
      index < 0 || index >= count
        ? null
        : cff.subarray(dataAt + offsetOf(index), dataAt + offsetOf(index + 1)),
  };
}

// A dictionary writes its operands and then the operator they belong to. The
// two-byte operators are keyed by twelve hundred and their second byte, so one
// map holds both kinds.
const ESCAPED_OPERATOR = 12;
const ESCAPE_KEY = 1200;

function readCffDict(data: Uint8Array): ReadonlyMap<number, readonly number[]> {
  const found = new Map<number, readonly number[]>();
  let operands: number[] = [];
  let at = 0;

  while (at < data.byteLength) {
    const first = data[at] ?? 0;

    if (first <= 21) {
      const key = first === ESCAPED_OPERATOR ? ESCAPE_KEY + (data[at + 1] ?? 0) : first;
      at += first === ESCAPED_OPERATOR ? 2 : 1;
      found.set(key, operands);
      operands = [];
    } else if (first === 28) {
      operands.push(viewOf(data).getInt16(at + 1));
      at += 3;
    } else if (first === 29) {
      operands.push(viewOf(data).getInt32(at + 1));
      at += 5;
    } else if (first === 30) {
      // A real number, written as nibbles and ended by an f. Nothing read here is
      // one, so it is stepped over rather than worked out.
      at += 1;
      while (at < data.byteLength) {
        const nibbles = data[at] ?? 0xff;
        at += 1;
        if ((nibbles & 0x0f) === 0x0f || nibbles >> 4 === 0x0f) break;
      }
      operands.push(0);
    } else if (first >= 32 && first <= 246) {
      operands.push(first - 139);
      at += 1;
    } else if (first >= 247 && first <= 250) {
      operands.push((first - 247) * 256 + (data[at + 1] ?? 0) + 108);
      at += 2;
    } else if (first >= 251 && first <= 254) {
      operands.push(-(first - 251) * 256 - (data[at + 1] ?? 0) - 108);
      at += 2;
    } else {
      at += 1;
    }
  }
  return found;
}

const CHARSTRINGS_KEY = 17;
const PRIVATE_KEY = 18;
const LOCAL_SUBRS_KEY = 19;
const CHARSTRING_TYPE_KEY = ESCAPE_KEY + 6;
const FD_ARRAY_KEY = ESCAPE_KEY + 36;
const FD_SELECT_KEY = ESCAPE_KEY + 37;

// A subroutine is called by a number counted from the middle of its index, so the
// bias has to be worked out from how many there are.
const biasOf = (count: number): number => (count < 1240 ? 107 : count < 33900 ? 1131 : 32768);

function localSubrsOf(
  cff: Uint8Array,
  dict: ReadonlyMap<number, readonly number[]>,
): CffIndex | null {
  const stated = dict.get(PRIVATE_KEY);
  const size = stated?.[0];
  const at = stated?.[1];
  if (size === undefined || at === undefined || at + size > cff.byteLength) return null;

  const subrsAt = readCffDict(cff.subarray(at, at + size)).get(LOCAL_SUBRS_KEY)?.[0];
  return subrsAt === undefined ? null : readCffIndex(cff, at + subrsAt);
}

// Which font dictionary each glyph of a CID-keyed face belongs to, since each of
// them carries local subroutines of its own.
function fontDictOf(
  cff: Uint8Array,
  at: number,
  glyphCount: number,
): ((glyph: number) => number) | null {
  if (at + 1 > cff.byteLength) return null;
  const view = viewOf(cff);
  const format = cff[at] ?? 0;

  if (format === 0) {
    if (at + 1 + glyphCount > cff.byteLength) return null;
    return (glyph) => cff[at + 1 + glyph] ?? 0;
  }
  if (format !== 3) return null;

  const ranges = view.getUint16(at + 1);
  if (at + 3 + ranges * 3 + 2 > cff.byteLength) return null;
  return (glyph) => {
    for (let index = 0; index < ranges; index += 1) {
      const record = at + 3 + index * 3;
      const next =
        index + 1 === ranges ? view.getUint16(at + 3 + ranges * 3) : view.getUint16(record + 3);
      if (glyph >= view.getUint16(record) && glyph < next) return cff[record + 2] ?? 0;
    }
    return 0;
  };
}

const CHARSTRING_TYPE_2 = 2;

function readCffInk(cff: Uint8Array): InkReading {
  if (cff.byteLength < 4) return INK_MALFORMED;

  const headerLength = cff[2] ?? 0;
  const names = readCffIndex(cff, headerLength);
  const tops = names === null ? null : readCffIndex(cff, names.end);
  const strings = tops === null ? null : readCffIndex(cff, tops.end);
  const globalSubrs = strings === null ? null : readCffIndex(cff, strings.end);
  const top = tops?.entry(0);
  if (globalSubrs === null || top === undefined || top === null) return INK_MALFORMED;

  const dict = readCffDict(top);
  if ((dict.get(CHARSTRING_TYPE_KEY)?.[0] ?? CHARSTRING_TYPE_2) !== CHARSTRING_TYPE_2) {
    return INK_UNSUPPORTED;
  }

  const charStringsAt = dict.get(CHARSTRINGS_KEY)?.[0];
  const charStrings = charStringsAt === undefined ? null : readCffIndex(cff, charStringsAt);
  if (charStrings === null) return INK_MALFORMED;

  const shared = localSubrsOf(cff, dict);
  const fdArrayAt = dict.get(FD_ARRAY_KEY)?.[0];
  const fdSelectAt = dict.get(FD_SELECT_KEY)?.[0];
  const fdArray = fdArrayAt === undefined ? null : readCffIndex(cff, fdArrayAt);
  const fontDict = fdSelectAt === undefined ? null : fontDictOf(cff, fdSelectAt, charStrings.count);

  const subrsFor = (glyph: number): CffIndex | null => {
    if (fdArray === null || fontDict === null) return shared;
    const entry = fdArray.entry(fontDict(glyph));
    return entry === null ? shared : localSubrsOf(cff, readCffDict(entry));
  };

  return {
    kind: "ink",
    inkOfGlyph: (glyph) => {
      const charstring = charStrings.entry(glyph);
      if (charstring === null) return null;
      return outlineBoxOf(charstring, globalSubrs, subrsFor(glyph));
    },
  };
}

// How deep a charstring may call into itself before this stops following it,
// which is the depth the format itself allows.
const CALL_DEPTH = 10;

type Pen = {
  x: number;
  y: number;
  box: MutableBox | null;
  stems: number;
  width: boolean;
  stack: number[];
  broken: boolean;
};

/**
 * The box a PostScript outline comes to, worked out by following the charstring:
 * a CFF face states no box for a glyph anywhere, unlike a TrueType one.
 *
 * Only what draws is followed. The hints are counted rather than read, since a
 * hint mask is as long as there are stems and stepping over it wrongly loses the
 * rest of the glyph.
 */
function outlineBoxOf(
  charstring: Uint8Array,
  globalSubrs: CffIndex,
  localSubrs: CffIndex | null,
): InkBox | null {
  const pen: Pen = { x: 0, y: 0, box: null, stems: 0, width: false, stack: [], broken: false };
  runCharstring(charstring, pen, globalSubrs, localSubrs, 0);
  if (pen.broken || pen.box === null) return null;
  return { ...pen.box };
}

function runCharstring(
  code: Uint8Array,
  pen: Pen,
  globalSubrs: CffIndex,
  localSubrs: CffIndex | null,
  depth: number,
): void {
  if (depth > CALL_DEPTH) {
    pen.broken = true;
    return;
  }

  const view = viewOf(code);
  let at = 0;

  while (at < code.byteLength && !pen.broken) {
    const operator = code[at] ?? 0;

    if (operator >= 32 || operator === 28) {
      at = pushOperand(code, view, at, pen);
      continue;
    }
    at += 1;

    switch (operator) {
      case 1:
      case 3:
      case 18:
      case 23:
        countStems(pen);
        break;
      case 19:
      case 20:
        countStems(pen);
        at += Math.ceil(pen.stems / 8);
        break;
      case 21: {
        const moved = takeFrom(pen, 2);
        moveBy(pen, moved[0] ?? 0, moved[1] ?? 0);
        break;
      }
      case 22:
        moveBy(pen, takeFrom(pen, 1)[0] ?? 0, 0);
        break;
      case 4:
        moveBy(pen, 0, takeFrom(pen, 1)[0] ?? 0);
        break;
      case 5:
        drawLines(pen, pen.stack.splice(0));
        break;
      case 6:
        drawAlternatingLines(pen, pen.stack.splice(0), true);
        break;
      case 7:
        drawAlternatingLines(pen, pen.stack.splice(0), false);
        break;
      case 8:
        drawCurves(pen, pen.stack.splice(0));
        break;
      case 24: {
        const stack = pen.stack.splice(0);
        const curves = Math.max(0, stack.length - 2);
        drawCurves(pen, stack.slice(0, curves));
        drawLines(pen, stack.slice(curves));
        break;
      }
      case 25: {
        const stack = pen.stack.splice(0);
        const lines = Math.max(0, stack.length - 6);
        drawLines(pen, stack.slice(0, lines));
        drawCurves(pen, stack.slice(lines));
        break;
      }
      case 26:
        drawSameAxisCurves(pen, pen.stack.splice(0), true);
        break;
      case 27:
        drawSameAxisCurves(pen, pen.stack.splice(0), false);
        break;
      case 30:
        drawTurningCurves(pen, pen.stack.splice(0), false);
        break;
      case 31:
        drawTurningCurves(pen, pen.stack.splice(0), true);
        break;
      case 10:
      case 29: {
        const subrs = operator === 10 ? localSubrs : globalSubrs;
        const called = pen.stack.pop();
        const entry =
          subrs === null || called === undefined ? null : subrs.entry(called + biasOf(subrs.count));
        if (entry === null) {
          pen.broken = true;
          break;
        }
        runCharstring(entry, pen, globalSubrs, localSubrs, depth + 1);
        break;
      }
      case 11:
        return;
      case 14:
        pen.stack.length = 0;
        return;
      case ESCAPED_OPERATOR: {
        const escaped = code[at] ?? 0;
        at += 1;
        if (!drawFlex(pen, escaped)) pen.broken = true;
        break;
      }
      default:
        pen.broken = true;
    }
  }
}

function pushOperand(code: Uint8Array, view: DataView, at: number, pen: Pen): number {
  const first = code[at] ?? 0;
  if (first === 28) {
    pen.stack.push(view.getInt16(at + 1));
    return at + 3;
  }
  if (first <= 246) {
    pen.stack.push(first - 139);
    return at + 1;
  }
  if (first <= 250) {
    pen.stack.push((first - 247) * 256 + (code[at + 1] ?? 0) + 108);
    return at + 2;
  }
  if (first <= 254) {
    pen.stack.push(-(first - 251) * 256 - (code[at + 1] ?? 0) - 108);
    return at + 2;
  }
  // A sixteen-sixteen fixed point number, which a charstring may state where a
  // whole one will not do.
  pen.stack.push(view.getInt32(at + 1) / FIXED);
  return at + 5;
}

// The first operator that clears the stack may carry the glyph's width ahead of
// its own operands, which is how a face states a width other than its default.
function takeFrom(pen: Pen, wanted: number): readonly number[] {
  const stack = pen.stack.splice(0);
  const from = !pen.width && stack.length > wanted ? stack.length - wanted : 0;
  pen.width = true;
  return stack.slice(from);
}

// An odd operand ahead of the stems is the glyph's width, which changes nothing
// about how many stems there are: either way it is two numbers to a stem.
function countStems(pen: Pen): void {
  pen.stems += Math.floor(pen.stack.splice(0).length / 2);
  pen.width = true;
}

function moveBy(pen: Pen, across: number, up: number): void {
  pen.x += across;
  pen.y += up;
  pen.box = extendedBy(pen.box, pen.x, pen.y);
}

function drawLines(pen: Pen, stack: readonly number[]): void {
  for (let at = 0; at + 1 < stack.length; at += 2) {
    pen.x += stack[at] ?? 0;
    pen.y += stack[at + 1] ?? 0;
    pen.box = extendedBy(pen.box, pen.x, pen.y);
  }
}

function drawAlternatingLines(pen: Pen, stack: readonly number[], horizontal: boolean): void {
  let across = horizontal;
  for (const step of stack) {
    if (across) pen.x += step;
    else pen.y += step;
    across = !across;
    pen.box = extendedBy(pen.box, pen.x, pen.y);
  }
}

function curveTo(
  pen: Pen,
  firstX: number,
  firstY: number,
  secondX: number,
  secondY: number,
  toX: number,
  toY: number,
): void {
  const from: readonly [number, number] = [pen.x, pen.y];
  const first: readonly [number, number] = [pen.x + firstX, pen.y + firstY];
  const second: readonly [number, number] = [first[0] + secondX, first[1] + secondY];
  const to: readonly [number, number] = [second[0] + toX, second[1] + toY];

  pen.box = alongCurve(pen.box, from, first, second, to);
  pen.x = to[0];
  pen.y = to[1];
}

function drawCurves(pen: Pen, stack: readonly number[]): void {
  for (let at = 0; at + 5 < stack.length; at += 6) {
    curveTo(
      pen,
      stack[at] ?? 0,
      stack[at + 1] ?? 0,
      stack[at + 2] ?? 0,
      stack[at + 3] ?? 0,
      stack[at + 4] ?? 0,
      stack[at + 5] ?? 0,
    );
  }
}

// The two curves that run mostly one way: each states four numbers, and the first
// of them may carry one more for the axis it otherwise holds still on.
function drawSameAxisCurves(pen: Pen, stack: readonly number[], vertical: boolean): void {
  const leaning = stack.length % 4 === 1;
  let odd = leaning ? (stack[0] ?? 0) : 0;
  let at = leaning ? 1 : 0;

  while (at + 3 < stack.length) {
    if (vertical) {
      curveTo(
        pen,
        odd,
        stack[at] ?? 0,
        stack[at + 1] ?? 0,
        stack[at + 2] ?? 0,
        0,
        stack[at + 3] ?? 0,
      );
    } else {
      curveTo(
        pen,
        stack[at] ?? 0,
        odd,
        stack[at + 1] ?? 0,
        stack[at + 2] ?? 0,
        stack[at + 3] ?? 0,
        0,
      );
    }
    odd = 0;
    at += 4;
  }
}

// The two curves that turn: they start along one axis and end along the other,
// swapping over with every four numbers, and a last fifth number ends the final
// curve off its own axis.
function drawTurningCurves(pen: Pen, stack: readonly number[], startsHorizontal: boolean): void {
  let horizontal = startsHorizontal;
  let at = 0;

  while (at + 3 < stack.length) {
    const extra = stack.length - at === 5 ? (stack[at + 4] ?? 0) : 0;
    if (horizontal) {
      curveTo(
        pen,
        stack[at] ?? 0,
        0,
        stack[at + 1] ?? 0,
        stack[at + 2] ?? 0,
        extra,
        stack[at + 3] ?? 0,
      );
    } else {
      curveTo(
        pen,
        0,
        stack[at] ?? 0,
        stack[at + 1] ?? 0,
        stack[at + 2] ?? 0,
        stack[at + 3] ?? 0,
        extra,
      );
    }
    horizontal = !horizontal;
    at += 4;
  }
}

const operand = (stack: readonly number[], index: number): number => stack[index] ?? 0;

const FLEX = 35;
const HFLEX = 34;
const HFLEX_1 = 36;
const FLEX_1 = 37;

// The four ways a face writes two curves that stand in for a nearly flat line.
// Only these of the escaped operators draw; the arithmetic ones are refused,
// since a glyph whose outline is worked out rather than stated is a glyph this
// cannot measure.
function drawFlex(pen: Pen, operator: number): boolean {
  const stack = pen.stack.splice(0);

  if (operator === FLEX && stack.length >= 13) {
    drawCurves(pen, stack.slice(0, 12));
    return true;
  }
  // Every one of the three shorter forms leaves out the numbers that would say
  // where the second curve ends, because it ends back on the line it started on.
  if (operator === HFLEX && stack.length >= 7) {
    curveTo(pen, operand(stack, 0), 0, operand(stack, 1), operand(stack, 2), operand(stack, 3), 0);
    curveTo(pen, operand(stack, 4), 0, operand(stack, 5), -operand(stack, 2), operand(stack, 6), 0);
    return true;
  }
  if (operator === HFLEX_1 && stack.length >= 9) {
    const startY = pen.y;
    curveTo(
      pen,
      operand(stack, 0),
      operand(stack, 1),
      operand(stack, 2),
      operand(stack, 3),
      operand(stack, 4),
      0,
    );
    const secondY = pen.y + operand(stack, 7);
    curveTo(
      pen,
      operand(stack, 5),
      0,
      operand(stack, 6),
      operand(stack, 7),
      operand(stack, 8),
      startY - secondY,
    );
    return true;
  }
  if (operator === FLEX_1 && stack.length >= 11) {
    const startX = pen.x;
    const startY = pen.y;
    const across =
      operand(stack, 0) +
      operand(stack, 2) +
      operand(stack, 4) +
      operand(stack, 6) +
      operand(stack, 8);
    const up =
      operand(stack, 1) +
      operand(stack, 3) +
      operand(stack, 5) +
      operand(stack, 7) +
      operand(stack, 9);

    curveTo(
      pen,
      operand(stack, 0),
      operand(stack, 1),
      operand(stack, 2),
      operand(stack, 3),
      operand(stack, 4),
      operand(stack, 5),
    );
    const secondX = pen.x + operand(stack, 6) + operand(stack, 8);
    const secondY = pen.y + operand(stack, 7) + operand(stack, 9);
    const mostlyAcross = Math.abs(across) > Math.abs(up);
    curveTo(
      pen,
      operand(stack, 6),
      operand(stack, 7),
      operand(stack, 8),
      operand(stack, 9),
      mostlyAcross ? operand(stack, 10) : startX - secondX,
      mostlyAcross ? startY - secondY : operand(stack, 10),
    );
    return true;
  }
  return false;
}

// A TrueType face states the box in the glyph's own header and a PostScript one
// states none anywhere, so which of the two the face carries decides how much
// work an answer is. Where a face carries both, which no real one does, the
// stated box is taken over the followed outline.
function readInkOfGlyph(tables: ReadonlyMap<string, Uint8Array>): InkReading {
  const truetype = readTrueTypeInk(tables);
  if (truetype.kind !== "missing") return truetype;

  const cff = tables.get(CFF);
  return cff === undefined ? INK_MISSING : readCffInk(cff);
}

function readInk(tables: ReadonlyMap<string, Uint8Array>, glyphFor: CodeToGlyph): InkTable {
  const reading = readInkOfGlyph(tables);
  if (reading.kind === "missing") return { kind: "unavailable", reason: "outlines-missing" };
  if (reading.kind === "unsupported")
    return { kind: "unavailable", reason: "outlines-unsupported" };
  if (reading.kind === "malformed") return { kind: "unavailable", reason: "outlines-malformed" };

  const { inkOfGlyph } = reading;
  return {
    kind: "ink",
    inkOfGlyph,
    // A character the face has no glyph for draws nothing here rather than
    // .notdef's own box: what another face would draw for it is a question this
    // file cannot answer.
    inkOf: (codePoint) => {
      const glyph = glyphFor(codePoint);
      return glyph === 0 ? null : inkOfGlyph(glyph);
    },
  };
}

// What a glyph advances, which the advance table answers in characters and a
// math variant has none of: the taller parenthesis is reached through the MATH
// table and is in no cmap at all.
function advanceOfGlyphIn(
  tables: ReadonlyMap<string, Uint8Array>,
  metricCount: number,
): (glyph: number) => number {
  const hmtx = tables.get(HMTX);
  const long = hmtx === undefined ? 0 : Math.min(metricCount, Math.floor(hmtx.byteLength / 4));
  if (hmtx === undefined || long < 1) return () => 0;

  const view = viewOf(hmtx);
  return (glyph) => view.getUint16(Math.min(Math.max(glyph, 0), long - 1) * 4);
}

const MATH_HEADER_LENGTH = 10;
const MATH_VALUE_LENGTH = 4;
const MATH_CONSTANTS_LENGTH = 8 + MATH_VALUE_CONSTANTS.length * MATH_VALUE_LENGTH + 2;

// The table as the file states it, which is in glyphs throughout: what a caller
// asks in characters is turned into this on the way in.
type MathReading =
  | {
      readonly kind: "math";
      readonly constants: MathConstants;
      readonly minConnectorOverlap: number;
      readonly italicCorrectionOfGlyph: (glyph: number) => number;
      readonly taller: Grown;
      readonly wider: Grown;
    }
  | { readonly kind: "malformed" };

// The fifty-one the table states as a value and a device offset, which are read
// in the order the table writes them rather than one by one, so that the names
// and the layout cannot drift apart.
function valueConstantsAt(view: DataView, at: number): Record<MathValueConstant, number> {
  const values: Record<string, number> = {};
  MATH_VALUE_CONSTANTS.forEach((name, index) => {
    values[name] = view.getInt16(at + index * MATH_VALUE_LENGTH);
  });
  return values;
}

function readMathConstants(math: Uint8Array, at: number): MathConstants | null {
  if (at + MATH_CONSTANTS_LENGTH > math.byteLength) return null;

  const view = viewOf(math);
  return {
    ...valueConstantsAt(view, at + 8),
    scriptPercentScaleDown: view.getInt16(at),
    scriptScriptPercentScaleDown: view.getInt16(at + 2),
    delimitedSubFormulaMinHeight: view.getUint16(at + 4),
    displayOperatorMinHeight: view.getUint16(at + 6),
    radicalDegreeBottomRaisePercent: view.getInt16(at + MATH_CONSTANTS_LENGTH - 2),
  };
}

// What the face says a glyph leans past its own advance, which is what a
// following character has to be moved by. Read by coverage, like a kern pair.
function readItalicCorrections(math: Uint8Array, at: number): ((glyph: number) => number) | null {
  if (at + 4 > math.byteLength) return null;

  const view = viewOf(math);
  const covered = readCoverage(math, at + view.getUint16(at));
  const count = view.getUint16(at + 2);
  if (covered === null || at + 4 + count * MATH_VALUE_LENGTH > math.byteLength) return null;

  return (glyph) => {
    const index = covered.indexOf(glyph);
    return index === null || index >= count ? 0 : view.getInt16(at + 4 + index * MATH_VALUE_LENGTH);
  };
}

const CONSTRUCTION_HEADER_LENGTH = 4;
const VARIANT_LENGTH = 4;
const ASSEMBLY_PART_LENGTH = 10;
const EXTENDER_FLAG = 1;

type Grown = {
  readonly variantsOf: (glyph: number) => readonly MathVariant[];
  readonly assemblyOf: (glyph: number) => MathAssembly | null;
};

function readGrowth(
  math: Uint8Array,
  variantsAt: number,
  coverageAt: number,
  constructionsAt: number,
  count: number,
  advanceOfGlyph: (glyph: number) => number,
  inkOfGlyph: GlyphInkOf,
): Grown | null {
  const covered = readCoverage(math, coverageAt);
  if (covered === null || constructionsAt + count * 2 > math.byteLength) return null;

  const view = viewOf(math);
  // A construction is named from the start of the variants table it hangs off,
  // not from the offsets that name it.
  const constructionFor = (glyph: number): number | null => {
    const index = covered.indexOf(glyph);
    if (index === null || index >= count) return null;
    return variantsAt + view.getUint16(constructionsAt + index * 2);
  };

  const resolved = (glyph: number, measurement: number): MathVariant => ({
    glyph,
    measurement,
    advance: advanceOfGlyph(glyph),
    ink: inkOfGlyph(glyph),
  });

  return {
    variantsOf: (glyph) => {
      const construction = constructionFor(glyph);
      if (construction === null || construction + CONSTRUCTION_HEADER_LENGTH > math.byteLength) {
        return [];
      }
      const variants = view.getUint16(construction + 2);
      if (construction + CONSTRUCTION_HEADER_LENGTH + variants * VARIANT_LENGTH > math.byteLength) {
        return [];
      }
      const found: MathVariant[] = [];
      for (let index = 0; index < variants; index += 1) {
        const record = construction + CONSTRUCTION_HEADER_LENGTH + index * VARIANT_LENGTH;
        found.push(resolved(view.getUint16(record), view.getUint16(record + 2)));
      }
      return found;
    },

    assemblyOf: (glyph) => {
      const construction = constructionFor(glyph);
      if (construction === null || construction + CONSTRUCTION_HEADER_LENGTH > math.byteLength) {
        return null;
      }
      const stated = view.getUint16(construction);
      if (stated === 0) return null;

      const assembly = construction + stated;
      if (assembly + 6 > math.byteLength) return null;
      const partCount = view.getUint16(assembly + 4);
      if (assembly + 6 + partCount * ASSEMBLY_PART_LENGTH > math.byteLength) return null;

      const parts: MathAssemblyPart[] = [];
      for (let index = 0; index < partCount; index += 1) {
        const record = assembly + 6 + index * ASSEMBLY_PART_LENGTH;
        const part = view.getUint16(record);
        parts.push({
          glyph: part,
          startConnector: view.getUint16(record + 2),
          endConnector: view.getUint16(record + 4),
          fullAdvance: view.getUint16(record + 6),
          extender: (view.getUint16(record + 8) & EXTENDER_FLAG) !== 0,
          advance: advanceOfGlyph(part),
          ink: inkOfGlyph(part),
        });
      }
      return { italicCorrection: view.getInt16(assembly), parts };
    },
  };
}

const NO_GROWTH: Grown = { variantsOf: () => [], assemblyOf: () => null };

function readMathTable(
  math: Uint8Array,
  advanceOfGlyph: (glyph: number) => number,
  inkOfGlyph: GlyphInkOf,
): MathReading {
  if (math.byteLength < MATH_HEADER_LENGTH) return { kind: "malformed" };

  const view = viewOf(math);
  const constants = readMathConstants(math, view.getUint16(4));
  if (constants === null) return { kind: "malformed" };

  const glyphInfoAt = view.getUint16(6);
  const variantsAt = view.getUint16(8);

  let italicCorrectionOfGlyph: (glyph: number) => number = () => 0;
  if (glyphInfoAt !== 0) {
    if (glyphInfoAt + 2 > math.byteLength) return { kind: "malformed" };
    const italicsAt = view.getUint16(glyphInfoAt);
    if (italicsAt !== 0) {
      const read = readItalicCorrections(math, glyphInfoAt + italicsAt);
      if (read === null) return { kind: "malformed" };
      italicCorrectionOfGlyph = read;
    }
  }

  let taller = NO_GROWTH;
  let wider = NO_GROWTH;
  let minConnectorOverlap = 0;
  if (variantsAt !== 0) {
    if (variantsAt + 10 > math.byteLength) return { kind: "malformed" };
    minConnectorOverlap = view.getUint16(variantsAt);
    const verticalCount = view.getUint16(variantsAt + 6);
    const horizontalCount = view.getUint16(variantsAt + 8);
    const verticalCoverage = view.getUint16(variantsAt + 2);
    const horizontalCoverage = view.getUint16(variantsAt + 4);
    const verticalAt = variantsAt + 10;
    const horizontalAt = verticalAt + verticalCount * 2;

    const vertical =
      verticalCoverage === 0
        ? NO_GROWTH
        : readGrowth(
            math,
            variantsAt,
            variantsAt + verticalCoverage,
            verticalAt,
            verticalCount,
            advanceOfGlyph,
            inkOfGlyph,
          );
    const horizontal =
      horizontalCoverage === 0
        ? NO_GROWTH
        : readGrowth(
            math,
            variantsAt,
            variantsAt + horizontalCoverage,
            horizontalAt,
            horizontalCount,
            advanceOfGlyph,
            inkOfGlyph,
          );
    if (vertical === null || horizontal === null) return { kind: "malformed" };
    taller = vertical;
    wider = horizontal;
  }

  return { kind: "math", constants, minConnectorOverlap, italicCorrectionOfGlyph, taller, wider };
}

function readMath(
  tables: ReadonlyMap<string, Uint8Array>,
  glyphFor: CodeToGlyph,
  metricCount: number,
  ink: InkTable,
): MathTable {
  const math = tables.get(MATH);
  if (math === undefined) return { kind: "unavailable", reason: "math-missing" };

  const inkOfGlyph = ink.kind === "ink" ? ink.inkOfGlyph : () => null;
  const read = readMathTable(math, advanceOfGlyphIn(tables, metricCount), inkOfGlyph);
  if (read.kind !== "math") return { kind: "unavailable", reason: "math-malformed" };

  // The table is written in glyphs and a caller asks in characters, as it does of
  // the advances and the pairs, so every way in goes through the one cmap.
  return {
    kind: "math",
    constants: read.constants,
    minConnectorOverlap: read.minConnectorOverlap,
    italicCorrectionOf: (codePoint) => read.italicCorrectionOfGlyph(glyphFor(codePoint)),
    tallerVariantsOf: (codePoint) => read.taller.variantsOf(glyphFor(codePoint)),
    widerVariantsOf: (codePoint) => read.wider.variantsOf(glyphFor(codePoint)),
    piecesToGrowTaller: (codePoint) => read.taller.assemblyOf(glyphFor(codePoint)),
    piecesToGrowWider: (codePoint) => read.wider.assemblyOf(glyphFor(codePoint)),
  };
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
  // A face whose characters cannot be mapped to glyphs answers for none of the
  // three tables below: every one of them is written in glyphs, and there is
  // nothing to look them up by.
  const unmapped =
    index.kind === "unavailable" ? ({ kind: "unavailable", reason: index.reason } as const) : null;
  const glyphFor = index.kind === "glyphs" ? index.glyphFor : NOTHING_MAPPED;
  const ink = unmapped ?? readInk(tables, glyphFor);

  return {
    format,
    metrics: {
      unitsPerEm: headView.getUint16(18),
      ascender: hheaView.getInt16(4),
      descender: hheaView.getInt16(6),
      lineGap: hheaView.getInt16(8),
    },
    advances: readAdvanceTable(tables, metricCount),
    kerning: unmapped ?? readKerning(tables, glyphFor),
    ink,
    math: unmapped ?? readMath(tables, glyphFor, metricCount, ink),
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
