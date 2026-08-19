import { zlibSync } from "fflate";

import { readFontFile } from "../layout/font-file.js";
import {
  MATH_VALUE_CONSTANTS,
  type FontMetrics,
  type MathConstants,
  type MathValueConstant,
  type SuppliedFace,
} from "../layout/font-metrics.js";

export type FontFixture = {
  readonly unitsPerEm: number;
  readonly ascender: number;
  readonly descender: number;
  readonly lineGap: number;
  readonly advances?: Readonly<Record<string, number>>;
  readonly cmapFormat?: 0 | 2 | 4 | 6 | 12 | 13;
  // What a whole run of characters draws, which is what a format 13 cmap states
  // and nothing else does: the character named states the glyph, and the one it
  // is mapped to is the last of the run drawing it.
  readonly sharedRanges?: Readonly<Record<string, string>>;
  // A subtable in a format nothing here reads, declared at the widest encoding so
  // that it outranks the fixture's own. A face states one to ask whether the
  // subtable that can be read is passed over for it.
  readonly unreadableSubtable?: boolean;
  // What a format 2 cmap maps. Its keys are bytes rather than characters, and its
  // glyphs are stated by number rather than taken from the fixture's own, since
  // what a format 2 subtable is indexed by is a character written in one of the
  // legacy multi-byte encodings and not the character itself.
  readonly highByteMapping?: HighByteMapping;
  // Which encodings the face declares it maps. A symbol face declares the symbol
  // one, and Symbol itself declares both: its glyphs are reachable by byte through
  // its own page and by what they mean.
  readonly subtables?: readonly ("unicode" | "symbol")[];
  // What .notdef advances, which is what a symbol face answers for a character its
  // page has no glyph for.
  readonly notdefAdvance?: number;
  readonly longMetrics?: number;
  // What the face says about itself in its PANOSE classification, which is what
  // decides the face a character it has no glyph for is drawn out of. A fixture
  // that states neither writes no OS/2 table at all, as a face may.
  readonly panoseFamily?: number;
  readonly panoseSerifStyle?: number;
  // What the face calls itself, which is how one is picked out of a collection.
  readonly faceName?: string;
  // The family the face belongs to, where that is not the whole of its own name:
  // `Calibri Light` is a face of the `Calibri` family, and a document naming
  // either has to find it.
  readonly familyName?: string;
  // What the face says about its own weight and slope in `head`.
  readonly bold?: boolean;
  readonly italic?: boolean;
  // Where the face puts the line under its letters, how thick it is and how far
  // they lean, all of which it states in its `post` table. A fixture stating none
  // of them writes no table at all, as a face may.
  readonly underlinePosition?: number;
  readonly underlineThickness?: number;
  readonly italicAngle?: number;
  // What each pair of the face's characters moves, in font units. `kernPairs`
  // writes the legacy `kern` table and the other two write GPOS pair positioning,
  // and a fixture may state both to ask which of them is read.
  readonly kernPairs?: KerningPairs;
  readonly gposPairs?: KerningPairs;
  readonly gposClassPairs?: ClassKerningPairs;
  // A second pair lookup, stating its movement as a placement beside the advance
  // the way a face states the pairs of a script that runs the other way. The
  // reader will not guess at one, so a fixture states it to ask what happens to
  // the pairs beside it.
  readonly gposRightToLeftPairs?: KerningPairs;
  // What the kern subtable says it holds, for a fixture that wants one whose count
  // lies about its own pairs.
  readonly claimedKernPairs?: number;
  // What the kern subtable says it is long, for a fixture that wants one whose
  // length is too small for its own pairs, as Calibri's and Cambria's are.
  readonly claimedKernLength?: number;
  // How many bytes to cut off the end of the GPOS table, which is how a fixture
  // states one whose offsets run past what is there.
  readonly cutFromGpos?: number;
  // What the pair positioning states it moves, as the two value formats. The
  // default is an X advance on the first glyph and nothing on the second, which is
  // the only movement the reader takes; a fixture states another to ask for the
  // refusal.
  readonly gposValueFormats?: readonly [first: number, second: number];
  // Whether the pair lookup is reached through an extension, which is how a real
  // face reaches a subtable further into the file than a two-byte offset can name.
  readonly gposThroughExtension?: boolean;
  // What each character draws, as the box its outline fills. Written as a
  // TrueType glyph, which states its box in its own header; a character the
  // fixture states no box for is written as a glyph of no length, which is how a
  // face writes a space.
  readonly boxes?: Readonly<Record<string, InkFixture>>;
  // Whether `loca` counts in bytes or in pairs of them, which is what a face with
  // more outline than a two-byte offset can reach has to state.
  readonly locaFormat?: "short" | "long";
  // What each character draws, as the path its outline follows. Written as a
  // PostScript charstring, which states no box at all: the box is what the
  // outline comes to, and a curve reaches past its own ends.
  readonly outlines?: Readonly<Record<string, GlyphPath>>;
  // What the face says about setting mathematics, and how many bytes to cut off
  // the end of what it says, for a fixture that wants a table whose offsets run
  // past what is there.
  readonly math?: MathFixture;
  readonly cutFromMath?: number;
  readonly omit?: "head" | "hhea" | "cmap" | "hmtx";
};

export type InkFixture = {
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
  readonly top: number;
};

// A move to where the outline starts, and then the lines and curves that close
// it. A curve states its two control points and where it ends.
export type GlyphPath = {
  readonly from: readonly [number, number];
  readonly steps: readonly GlyphStep[];
};

export type GlyphStep =
  | { readonly line: readonly [number, number] }
  | { readonly curve: readonly [number, number, number, number, number, number] };

export type MathFixture = {
  readonly constants?: Partial<MathConstants>;
  readonly italicCorrections?: Readonly<Record<string, number>>;
  // What each character grows through, as the character whose glyph is drawn and
  // how far that glyph reaches along the axis it grows on.
  readonly tallerVariants?: Readonly<Record<string, readonly MathVariantFixture[]>>;
  readonly widerVariants?: Readonly<Record<string, readonly MathVariantFixture[]>>;
  readonly tallerPieces?: Readonly<Record<string, MathPiecesFixture>>;
  readonly minConnectorOverlap?: number;
};

export type MathVariantFixture = {
  readonly character: string;
  readonly measurement: number;
};

export type MathPiecesFixture = {
  readonly italicCorrection?: number;
  readonly parts: readonly {
    readonly character: string;
    readonly startConnector: number;
    readonly endConnector: number;
    readonly fullAdvance: number;
    readonly extender?: boolean;
  }[];
};

// A pair of the face's characters and what it moves: `{ AV: -80 }` says an A
// followed by a V closes up by 80 font units. Both characters have to be ones the
// fixture states an advance for, since a pair is stated in glyphs.
export type KerningPairs = Readonly<Record<string, number>>;

export type ClassKerningPairs = {
  // The characters in each class from 1 up; class 0 is every other glyph.
  readonly firstClasses: readonly string[];
  readonly secondClasses: readonly string[];
  // What a pair of classes moves, indexed by the first class and then the second,
  // both counting from class 0.
  readonly values: readonly (readonly number[])[];
};

type Glyph = {
  readonly codePoint: number;
  readonly id: number;
  readonly advance: number;
};

const HEAD_LENGTH = 54;
const HHEA_LENGTH = 36;

function glyphsOf(fixture: FontFixture): readonly Glyph[] {
  return Object.entries(fixture.advances ?? {})
    .map(([character, advance]) => ({ codePoint: character.codePointAt(0) ?? 0, advance }))
    .sort((left, right) => left.codePoint - right.codePoint)
    .map((entry, index) => ({ ...entry, id: index + 1 }));
}

// Glyph 0 is .notdef and carries no advance of its own, so it takes one slot ahead
// of the fixture's glyphs.
const longMetricCount = (fixture: FontFixture, glyphs: readonly Glyph[]): number =>
  fixture.longMetrics ?? glyphs.length + 1;

const MAC_STYLE_AT = 44;
const LOCA_FORMAT_AT = 50;

function headTable(fixture: FontFixture): Uint8Array {
  const table = new Uint8Array(HEAD_LENGTH);
  const view = new DataView(table.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint32(12, 0x5f0f3cf5);
  view.setUint16(18, fixture.unitsPerEm);
  view.setUint16(MAC_STYLE_AT, (fixture.bold === true ? 1 : 0) | (fixture.italic === true ? 2 : 0));
  view.setInt16(LOCA_FORMAT_AT, fixture.locaFormat === "long" ? 1 : 0);
  return table;
}

// A glyph that draws the box it states, as four points around it. Nothing here
// reads the outline itself, but a header stating a box the outline does not fill
// is not a glyph any face would write.
function glyphOutline(box: InkFixture): Uint8Array {
  const table = new Uint8Array(34);
  const view = new DataView(table.buffer);

  view.setInt16(0, 1);
  view.setInt16(2, box.left);
  view.setInt16(4, box.bottom);
  view.setInt16(6, box.right);
  view.setInt16(8, box.top);
  view.setUint16(10, 3);
  view.setUint16(12, 0);
  for (let point = 0; point < 4; point += 1) table[14 + point] = 1;

  const across = [box.left, box.right - box.left, 0, box.left - box.right];
  const up = [box.bottom, 0, box.top - box.bottom, 0];
  across.forEach((step, point) => {
    view.setInt16(18 + point * 2, step);
  });
  up.forEach((step, point) => {
    view.setInt16(26 + point * 2, step);
  });
  return table;
}

function trueTypeOutlines(
  fixture: FontFixture,
  glyphs: readonly Glyph[],
): readonly (readonly [string, Uint8Array])[] {
  const stated = fixture.boxes ?? {};
  const drawn: Uint8Array[] = [new Uint8Array(0)];
  for (const glyph of glyphs) {
    const box = Object.entries(stated).find(
      ([character]) => character.codePointAt(0) === glyph.codePoint,
    )?.[1];
    drawn.push(box === undefined ? new Uint8Array(0) : glyphOutline(box));
  }

  const glyf = concat(drawn);
  const long = fixture.locaFormat === "long";
  const loca = new Uint8Array((drawn.length + 1) * (long ? 4 : 2));
  const view = new DataView(loca.buffer);

  let at = 0;
  drawn.forEach((glyph, index) => {
    if (long) view.setUint32(index * 4, at);
    else view.setUint16(index * 2, at / 2);
    at += glyph.length;
  });
  if (long) view.setUint32(drawn.length * 4, at);
  else view.setUint16(drawn.length * 2, at / 2);

  return [
    ["loca", loca],
    ["glyf", glyf],
  ];
}

// A charstring writes its numbers before the operator that takes them. Anything
// that will not fit in one byte is written as the two the format keeps for it.
function charstringNumber(value: number): Uint8Array {
  if (value >= -107 && value <= 107) return Uint8Array.from([value + 139]);
  const written = new Uint8Array(3);
  written[0] = 28;
  new DataView(written.buffer).setInt16(1, value);
  return written;
}

const R_MOVE_TO = 21;
const R_LINE_TO = 5;
const R_CURVE_TO = 8;
const END_CHAR = 14;

function charstringOf(path: GlyphPath): Uint8Array {
  const written: Uint8Array[] = [
    charstringNumber(path.from[0]),
    charstringNumber(path.from[1]),
    Uint8Array.from([R_MOVE_TO]),
  ];

  for (const step of path.steps) {
    if ("line" in step) {
      written.push(...step.line.map(charstringNumber), Uint8Array.from([R_LINE_TO]));
    } else {
      written.push(...step.curve.map(charstringNumber), Uint8Array.from([R_CURVE_TO]));
    }
  }

  written.push(Uint8Array.from([END_CHAR]));
  return concat(written);
}

// An index of entries, which is how a PostScript font keeps every list it has.
function cffIndex(entries: readonly Uint8Array[]): Uint8Array {
  if (entries.length === 0) return new Uint8Array(2);

  const offsetSize = 4;
  const header = 3 + (entries.length + 1) * offsetSize;
  const out = new Uint8Array(header + entries.reduce((sum, each) => sum + each.length, 0));
  const view = new DataView(out.buffer);
  view.setUint16(0, entries.length);
  out[2] = offsetSize;

  let at = 1;
  entries.forEach((entry, index) => {
    view.setUint32(3 + index * offsetSize, at);
    out.set(entry, header + at - 1);
    at += entry.length;
  });
  view.setUint32(3 + entries.length * offsetSize, at);
  return out;
}

const CHARSTRINGS_OPERATOR = 17;
const CFF_HEADER_LENGTH = 4;
// The dictionary is one thirty-two bit number and the operator it belongs to, so
// its length is known before the offset inside it is.
const TOP_DICT_LENGTH = 6;

function cffTable(fixture: FontFixture, glyphs: readonly Glyph[]): Uint8Array {
  const stated = fixture.outlines ?? {};
  const charstrings: Uint8Array[] = [Uint8Array.from([END_CHAR])];
  for (const glyph of glyphs) {
    const path = Object.entries(stated).find(
      ([character]) => character.codePointAt(0) === glyph.codePoint,
    )?.[1];
    charstrings.push(path === undefined ? Uint8Array.from([END_CHAR]) : charstringOf(path));
  }

  const header = Uint8Array.from([1, 0, CFF_HEADER_LENGTH, 1]);
  const names = cffIndex([Uint8Array.from(Array.from("Meridian", (each) => each.charCodeAt(0)))]);
  const strings = cffIndex([]);
  const globalSubrs = cffIndex([]);
  const topIndexLength = 3 + 2 * 4 + TOP_DICT_LENGTH;
  const charStringsAt =
    header.length + names.length + topIndexLength + strings.length + globalSubrs.length;

  const dict = new Uint8Array(TOP_DICT_LENGTH);
  dict[0] = 29;
  new DataView(dict.buffer).setInt32(1, charStringsAt);
  dict[5] = CHARSTRINGS_OPERATOR;

  return concat([header, names, cffIndex([dict]), strings, globalSubrs, cffIndex(charstrings)]);
}

function hheaTable(fixture: FontFixture, glyphs: readonly Glyph[]): Uint8Array {
  const table = new Uint8Array(HHEA_LENGTH);
  const view = new DataView(table.buffer);
  view.setUint32(0, 0x00010000);
  view.setInt16(4, fixture.ascender);
  view.setInt16(6, fixture.descender);
  view.setInt16(8, fixture.lineGap);
  view.setUint16(34, longMetricCount(fixture, glyphs));
  return table;
}

function hmtxTable(fixture: FontFixture, glyphs: readonly Glyph[]): Uint8Array {
  const total = glyphs.length + 1;
  const long = Math.min(longMetricCount(fixture, glyphs), total);
  const table = new Uint8Array(long * 4 + (total - long) * 2);
  const view = new DataView(table.buffer);
  const advanceOf = (id: number): number =>
    id === 0
      ? (fixture.notdefAdvance ?? 0)
      : (glyphs.find((glyph) => glyph.id === id)?.advance ?? 0);

  for (let id = 0; id < long; id += 1) view.setUint16(id * 4, advanceOf(id));
  return table;
}

function cmapFormat4(glyphs: readonly Glyph[]): Uint8Array {
  const mapped = glyphs.filter((glyph) => glyph.codePoint <= 0xffff);
  const segCount = mapped.length + 1;
  const length = 16 + segCount * 8;
  const table = new Uint8Array(length);
  const view = new DataView(table.buffer);

  const endAt = 14;
  const startAt = endAt + segCount * 2 + 2;
  const deltaAt = startAt + segCount * 2;
  const rangeAt = deltaAt + segCount * 2;

  view.setUint16(0, 4);
  view.setUint16(2, length);
  view.setUint16(6, segCount * 2);

  mapped.forEach((glyph, index) => {
    view.setUint16(endAt + index * 2, glyph.codePoint);
    view.setUint16(startAt + index * 2, glyph.codePoint);
    view.setUint16(deltaAt + index * 2, (glyph.id - glyph.codePoint) & 0xffff);
    view.setUint16(rangeAt + index * 2, 0);
  });

  const last = mapped.length;
  view.setUint16(endAt + last * 2, 0xffff);
  view.setUint16(startAt + last * 2, 0xffff);
  view.setUint16(deltaAt + last * 2, 1);
  view.setUint16(rangeAt + last * 2, 0);

  return table;
}

function cmapFormat12(glyphs: readonly Glyph[]): Uint8Array {
  const length = 16 + glyphs.length * 12;
  const table = new Uint8Array(length);
  const view = new DataView(table.buffer);

  view.setUint16(0, 12);
  view.setUint32(4, length);
  view.setUint32(12, glyphs.length);

  glyphs.forEach((glyph, index) => {
    const group = 16 + index * 12;
    view.setUint32(group, glyph.codePoint);
    view.setUint32(group + 4, glyph.codePoint);
    view.setUint32(group + 8, glyph.id);
  });

  return table;
}

// The whole array, whether the face maps the character or not: a character with no
// glyph of its own holds .notdef where its byte is, which is the only way this
// format has of saying so.
function cmapFormat0(glyphs: readonly Glyph[]): Uint8Array {
  const table = new Uint8Array(262);
  const view = new DataView(table.buffer);
  view.setUint16(0, 0);
  view.setUint16(2, table.length);

  for (const glyph of glyphs) if (glyph.codePoint <= 0xff) table[6 + glyph.codePoint] = glyph.id;
  return table;
}

// One run from the lowest character the face maps to the highest, so that a
// character the fixture leaves out of the middle of it is in the run and unmapped.
function cmapFormat6(glyphs: readonly Glyph[]): Uint8Array {
  const mapped = glyphs.filter((glyph) => glyph.codePoint <= 0xffff);
  const first = mapped[0]?.codePoint ?? 0;
  const last = mapped[mapped.length - 1]?.codePoint ?? first;
  const count = mapped.length === 0 ? 0 : last - first + 1;

  const table = new Uint8Array(10 + count * 2);
  const view = new DataView(table.buffer);
  view.setUint16(0, 6);
  view.setUint16(2, table.length);
  view.setUint16(6, first);
  view.setUint16(8, count);

  for (const glyph of mapped) view.setUint16(10 + (glyph.codePoint - first) * 2, glyph.id);
  return table;
}

// Format 12's groups, meaning the opposite: a group draws the one glyph rather
// than a consecutive run of them. A character the fixture states no run for stands
// on its own, so a fixture that states none writes the map format 12 would.
function cmapFormat13(fixture: FontFixture, glyphs: readonly Glyph[]): Uint8Array {
  const table = new Uint8Array(16 + glyphs.length * 12);
  const view = new DataView(table.buffer);
  view.setUint16(0, 13);
  view.setUint32(4, table.length);
  view.setUint32(12, glyphs.length);

  glyphs.forEach((glyph, index) => {
    const shared = Object.entries(fixture.sharedRanges ?? {}).find(
      ([character]) => character.codePointAt(0) === glyph.codePoint,
    );
    const group = 16 + index * 12;
    view.setUint32(group, glyph.codePoint);
    view.setUint32(group + 4, shared?.[1].codePointAt(0) ?? glyph.codePoint);
    view.setUint32(group + 8, glyph.id);
  });

  return table;
}

// What a format 2 cmap maps: a byte that is a character on its own, and the second
// byte of each pair under the first byte that leads it. Both name the glyph by
// number.
export type HighByteMapping = {
  readonly singleBytes?: Readonly<Record<number, number>>;
  readonly leadBytes?: Readonly<Record<number, Readonly<Record<number, number>>>>;
  // What the subheaders state their glyphs are offset by, which is how a face
  // keeps the array they share short. A glyph of zero is .notdef whatever the
  // offset says, so the array holds a zero rather than the offset backwards.
  readonly delta?: number;
};

type ByteRun = {
  readonly first: number;
  readonly count: number;
  readonly glyphs: readonly number[];
};

// One run from the lowest byte the subheader maps to the highest, so that a byte
// the fixture leaves out of the middle of it is in the run and unmapped.
function byteRun(mapping: Readonly<Record<number, number>>): ByteRun {
  const bytes = Object.keys(mapping)
    .map(Number)
    .sort((left, right) => left - right);
  const first = bytes[0] ?? 0;
  const last = bytes[bytes.length - 1] ?? first;
  const count = bytes.length === 0 ? 0 : last - first + 1;

  return {
    first,
    count,
    glyphs: Array.from({ length: count }, (_, index) => mapping[first + index] ?? 0),
  };
}

const SUBHEADER_KEYS_AT = 6;
const FORMAT_2_SUBHEADERS_AT = SUBHEADER_KEYS_AT + 256 * 2;
const SUBHEADER_LENGTH = 8;

// Subheader zero is where every byte that is a character on its own is looked up,
// and there is one more for each byte that leads a pair. They share one glyph
// array, which each reaches through an offset counted from the field that states
// it. A byte the keys say nothing about goes to subheader zero, which is what
// leaves it a single-byte character and not a lead.
export function buildHighByteSubtable(mapping: HighByteMapping): Uint8Array {
  const leads = Object.keys(mapping.leadBytes ?? {})
    .map(Number)
    .sort((left, right) => left - right);
  const runs = [
    byteRun(mapping.singleBytes ?? {}),
    ...leads.map((lead) => byteRun(mapping.leadBytes?.[lead] ?? {})),
  ];
  const delta = mapping.delta ?? 0;

  const glyphsAt = FORMAT_2_SUBHEADERS_AT + runs.length * SUBHEADER_LENGTH;
  const table = new Uint8Array(glyphsAt + runs.reduce((sum, run) => sum + run.count, 0) * 2);
  const view = new DataView(table.buffer);
  view.setUint16(0, 2);
  view.setUint16(2, table.length);

  let written = 0;
  runs.forEach((run, index) => {
    const at = FORMAT_2_SUBHEADERS_AT + index * SUBHEADER_LENGTH;
    view.setUint16(at, run.first);
    view.setUint16(at + 2, run.count);
    view.setInt16(at + 4, delta);
    view.setUint16(at + 6, glyphsAt + written * 2 - (at + 6));

    run.glyphs.forEach((glyph, offset) => {
      view.setUint16(glyphsAt + (written + offset) * 2, glyph === 0 ? 0 : (glyph - delta) & 0xffff);
    });
    written += run.count;
  });

  leads.forEach((lead, index) => {
    view.setUint16(SUBHEADER_KEYS_AT + lead * 2, (index + 1) * SUBHEADER_LENGTH);
  });

  return table;
}

function cmapSubtable(fixture: FontFixture, glyphs: readonly Glyph[]): Uint8Array {
  const format = fixture.cmapFormat ?? 4;
  if (format === 0) return cmapFormat0(glyphs);
  if (format === 2) return buildHighByteSubtable(fixture.highByteMapping ?? {});
  if (format === 6) return cmapFormat6(glyphs);
  if (format === 12) return cmapFormat12(glyphs);
  if (format === 13) return cmapFormat13(fixture, glyphs);
  return cmapFormat4(glyphs);
}

const WINDOWS_PLATFORM = 3;
const SYMBOL_ENCODING = 0;
const BASIC_PLANE_ENCODING = 1;
const FULL_REPERTOIRE_ENCODING = 10;

// The formats that reach past the basic plane are declared where a real face
// declares them. The byte-indexed ones are declared on the Windows platform rather
// than the Macintosh one a real face writes them on, since what a fixture asks
// about is how a subtable is read and not which of them is picked. Format 2 goes
// there too, so that a face stating one is refused for the format it is in rather
// than for the platform it would really sit on.
const encodingFor = (format: number): number =>
  format === 12 || format === 13 ? FULL_REPERTOIRE_ENCODING : BASIC_PLANE_ENCODING;

// Every encoding the fixture declares points at the one subtable, since what a
// fixture has to say here is which encodings a face offers rather than what each
// of them maps. The unreadable one is the exception, and comes first so that a
// reader passing it over has reached past a subtable it prefers and met first.
function cmapTable(fixture: FontFixture, glyphs: readonly Glyph[]): Uint8Array {
  const mapped = cmapSubtable(fixture, glyphs);
  const format = fixture.cmapFormat ?? 4;

  const records = [
    ...(fixture.unreadableSubtable === true
      ? [{ encoding: FULL_REPERTOIRE_ENCODING, bytes: buildHighByteSubtable({}) }]
      : []),
    ...(fixture.subtables ?? ["unicode"]).map((encoding) => ({
      encoding: encoding === "symbol" ? SYMBOL_ENCODING : encodingFor(format),
      bytes: mapped,
    })),
  ];

  const offsets = new Map<Uint8Array, number>();
  let end = 4 + records.length * 8;
  for (const record of records) {
    if (offsets.has(record.bytes)) continue;
    offsets.set(record.bytes, end);
    end += record.bytes.length;
  }

  const table = new Uint8Array(end);
  const view = new DataView(table.buffer);
  view.setUint16(2, records.length);

  records.forEach((record, index) => {
    const at = 4 + index * 8;
    view.setUint16(at, WINDOWS_PLATFORM);
    view.setUint16(at + 2, record.encoding);
    view.setUint32(at + 4, offsets.get(record.bytes) ?? 0);
  });

  for (const [bytes, offset] of offsets) table.set(bytes, offset);
  return table;
}

// Long enough to hold the PANOSE bytes, which sit ten bytes in and are the only
// part of the table anything here reads.
const OS2_LENGTH = 78;
const PANOSE_AT = 32;

function os2Table(fixture: FontFixture): Uint8Array {
  const table = new Uint8Array(OS2_LENGTH);
  table[PANOSE_AT] = fixture.panoseFamily ?? 0;
  table[PANOSE_AT + 1] = fixture.panoseSerifStyle ?? 0;
  return table;
}

// The family a face belongs to and the whole of its own name, which a collection
// is searched by. Written on the Windows platform, two bytes to a character.
const FAMILY_NAME = 1;
const FULL_NAME = 4;

function nameTable(faceName: string, familyName: string): Uint8Array {
  const written = [
    [FAMILY_NAME, familyName],
    [FULL_NAME, faceName],
  ] as const;
  const stringsAt = 6 + written.length * 12;
  const strings = written.map(([, value]) => value).join("");
  const table = new Uint8Array(stringsAt + strings.length * 2);
  const view = new DataView(table.buffer);

  view.setUint16(2, written.length);
  view.setUint16(4, stringsAt);
  let at = 0;
  written.forEach(([id, value], index) => {
    const record = 6 + index * 12;
    view.setUint16(record, 3);
    view.setUint16(record + 2, 1);
    view.setUint16(record + 6, id);
    view.setUint16(record + 8, value.length * 2);
    view.setUint16(record + 10, at * 2);
    for (const [step, character] of Array.from(value).entries()) {
      view.setUint16(stringsAt + (at + step) * 2, character.charCodeAt(0));
    }
    at += value.length;
  });

  return table;
}

// Long enough to hold the angle, the underline and its thickness, which are the
// only part of the table anything here reads.
const POST_LENGTH = 32;

// Stated as a distance up from the baseline, so a line under the letters is
// negative in the file. A fixture states where it wants the line, which is down.
function postTable(fixture: FontFixture): Uint8Array {
  const table = new Uint8Array(POST_LENGTH);
  const view = new DataView(table.buffer);
  view.setUint32(0, 0x00030000);
  view.setInt32(4, Math.round((fixture.italicAngle ?? 0) * 65536));
  view.setInt16(8, -(fixture.underlinePosition ?? 0));
  view.setInt16(10, fixture.underlineThickness ?? 0);
  return table;
}

const concat = (parts: readonly Uint8Array[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
};

type GlyphPair = { readonly left: number; readonly right: number; readonly value: number };

const idOf = (codePoint: number | undefined, glyphs: readonly Glyph[]): number =>
  glyphs.find((glyph) => glyph.codePoint === (codePoint ?? -1))?.id ?? 0;

function glyphPairs(stated: KerningPairs, glyphs: readonly Glyph[]): readonly GlyphPair[] {
  return Object.entries(stated)
    .map(([pair, value]) => ({
      left: idOf(pair.codePointAt(0), glyphs),
      right: idOf(pair.codePointAt(1), glyphs),
      value,
    }))
    .sort((left, right) => left.left - right.left || left.right - right.right);
}

const KERN_HEADER = 4;
const KERN_SUBTABLE_HEADER = 6;
const KERN_FORMAT_0_HEADER = 8;
const KERN_PAIR_LENGTH = 6;

// The legacy table as a Windows font states it: version zero, one format 0
// subtable, horizontal values. The search range and its two companions are left at
// zero, since a reader that maps the pairs never consults them.
function kernTable(fixture: FontFixture, glyphs: readonly Glyph[]): Uint8Array {
  const pairs = glyphPairs(fixture.kernPairs ?? {}, glyphs);
  const subtableLength =
    KERN_SUBTABLE_HEADER + KERN_FORMAT_0_HEADER + pairs.length * KERN_PAIR_LENGTH;
  const table = new Uint8Array(KERN_HEADER + subtableLength);
  const view = new DataView(table.buffer);

  view.setUint16(2, 1);
  view.setUint16(KERN_HEADER + 2, fixture.claimedKernLength ?? subtableLength);
  // Format 0 in the high byte, horizontal values in the low one.
  view.setUint16(KERN_HEADER + 4, 1);

  const body = KERN_HEADER + KERN_SUBTABLE_HEADER;
  view.setUint16(body, fixture.claimedKernPairs ?? pairs.length);
  pairs.forEach((pair, index) => {
    const at = body + KERN_FORMAT_0_HEADER + index * KERN_PAIR_LENGTH;
    view.setUint16(at, pair.left);
    view.setUint16(at + 2, pair.right);
    view.setInt16(at + 4, pair.value);
  });

  return table;
}

const X_PLACEMENT = 0x0001;
const X_ADVANCE = 0x0004;
const FIRST_GLYPH_ADVANCE: readonly [number, number] = [X_ADVANCE, 0];

// How a face states a pair of a script that runs the other way: the glyph is
// placed by exactly what it is advanced. Every lookup of the kind measured on this
// machine states it so, and none of them covers a Latin glyph.
const PLACED_AND_ADVANCED: readonly [number, number] = [X_PLACEMENT | X_ADVANCE, 0];

const valueLength = (format: number): number => {
  let length = 0;
  for (let bit = 1; bit <= 0x80; bit <<= 1) if ((format & bit) !== 0) length += 2;
  return length;
};

// A value record holds the fields its format states, in the order of the bits that
// state them. The movement goes in the X advance and, where the format states one,
// in the X placement beside it, which is how a face states a pair for text running
// the other way; everything else is written as zero.
function writeValue(view: DataView, at: number, format: number, value: number): void {
  let slot = at;
  for (let bit = 1; bit <= 0x80; bit <<= 1) {
    if ((format & bit) === 0) continue;
    view.setInt16(slot, bit === X_ADVANCE || bit === X_PLACEMENT ? value : 0);
    slot += 2;
  }
}

function coverageTable(glyphs: readonly number[]): Uint8Array {
  const table = new Uint8Array(4 + glyphs.length * 2);
  const view = new DataView(table.buffer);
  view.setUint16(0, 1);
  view.setUint16(2, glyphs.length);
  glyphs.forEach((glyph, index) => {
    view.setUint16(4 + index * 2, glyph);
  });
  return table;
}

// The other spelling of a coverage, which names runs of glyphs rather than each of
// them. Both are written by fixtures, since a face may state either.
function rangedCoverage(glyphs: readonly number[]): Uint8Array {
  const ranges: { start: number; end: number; at: number }[] = [];
  glyphs.forEach((glyph, index) => {
    const last = ranges.at(-1);
    if (last !== undefined && glyph === last.end + 1) last.end = glyph;
    else ranges.push({ start: glyph, end: glyph, at: index });
  });

  const table = new Uint8Array(4 + ranges.length * 6);
  const view = new DataView(table.buffer);
  view.setUint16(0, 2);
  view.setUint16(2, ranges.length);
  ranges.forEach((range, index) => {
    const at = 4 + index * 6;
    view.setUint16(at, range.start);
    view.setUint16(at + 2, range.end);
    view.setUint16(at + 4, range.at);
  });
  return table;
}

function classesOf(
  classes: readonly string[],
  glyphs: readonly Glyph[],
): ReadonlyMap<number, number> {
  const found = new Map<number, number>();
  classes.forEach((characters, index) => {
    for (const character of characters)
      found.set(idOf(character.codePointAt(0), glyphs), index + 1);
  });
  return found;
}

// Class 0 is every glyph the definition does not name, so only the classes from
// one up are written.
function rangedClassDefinition(classes: ReadonlyMap<number, number>): Uint8Array {
  const ranges: { start: number; end: number; value: number }[] = [];
  for (const glyph of [...classes.keys()].sort((left, right) => left - right)) {
    const value = classes.get(glyph) ?? 0;
    const last = ranges.at(-1);
    if (last !== undefined && glyph === last.end + 1 && value === last.value) last.end = glyph;
    else ranges.push({ start: glyph, end: glyph, value });
  }

  const table = new Uint8Array(4 + ranges.length * 6);
  const view = new DataView(table.buffer);
  view.setUint16(0, 2);
  view.setUint16(2, ranges.length);
  ranges.forEach((range, index) => {
    const at = 4 + index * 6;
    view.setUint16(at, range.start);
    view.setUint16(at + 2, range.end);
    view.setUint16(at + 4, range.value);
  });
  return table;
}

// The other spelling of a class definition, which lists a class for every glyph of
// one run.
function listedClassDefinition(classes: ReadonlyMap<number, number>): Uint8Array {
  const glyphs = [...classes.keys()].sort((left, right) => left - right);
  const start = glyphs[0] ?? 0;
  const count = glyphs.length === 0 ? 0 : (glyphs.at(-1) ?? start) - start + 1;

  const table = new Uint8Array(6 + count * 2);
  const view = new DataView(table.buffer);
  view.setUint16(0, 1);
  view.setUint16(2, start);
  view.setUint16(4, count);
  for (const [glyph, value] of classes) view.setUint16(6 + (glyph - start) * 2, value);
  return table;
}

const PAIR_SET_HEADER = 10;

function pairSetSubtable(
  pairs: readonly GlyphPair[],
  formats: readonly [number, number],
): Uint8Array {
  const byFirst = new Map<number, GlyphPair[]>();
  for (const pair of pairs) {
    const held = byFirst.get(pair.left);
    if (held === undefined) byFirst.set(pair.left, [pair]);
    else held.push(pair);
  }

  const firsts = [...byFirst.keys()].sort((left, right) => left - right);
  const recordLength = 2 + valueLength(formats[0]) + valueLength(formats[1]);
  const setLengths = firsts.map((first) => 2 + (byFirst.get(first)?.length ?? 0) * recordLength);
  const setsAt = PAIR_SET_HEADER + firsts.length * 2;
  const coverageAt = setsAt + setLengths.reduce((sum, each) => sum + each, 0);
  const coverage = coverageTable(firsts);

  const table = new Uint8Array(coverageAt + coverage.length);
  const view = new DataView(table.buffer);
  view.setUint16(0, 1);
  view.setUint16(2, coverageAt);
  view.setUint16(4, formats[0]);
  view.setUint16(6, formats[1]);
  view.setUint16(8, firsts.length);

  let at = setsAt;
  firsts.forEach((first, index) => {
    view.setUint16(PAIR_SET_HEADER + index * 2, at);
    const held = byFirst.get(first) ?? [];
    view.setUint16(at, held.length);
    held.forEach((pair, place) => {
      const record = at + 2 + place * recordLength;
      view.setUint16(record, pair.right);
      writeValue(view, record + 2, formats[0], pair.value);
      writeValue(view, record + 2 + valueLength(formats[0]), formats[1], 0);
    });
    at += setLengths[index] ?? 0;
  });

  table.set(coverage, coverageAt);
  return table;
}

const CLASS_PAIR_HEADER = 16;

function classPairSubtable(
  stated: ClassKerningPairs,
  glyphs: readonly Glyph[],
  formats: readonly [number, number],
): Uint8Array {
  const firstClasses = classesOf(stated.firstClasses, glyphs);
  const secondClasses = classesOf(stated.secondClasses, glyphs);
  const firstCount = stated.firstClasses.length + 1;
  const secondCount = stated.secondClasses.length + 1;
  const recordLength = valueLength(formats[0]) + valueLength(formats[1]);

  const firstDefinition = rangedClassDefinition(firstClasses);
  const secondDefinition = listedClassDefinition(secondClasses);
  const coverage = rangedCoverage([...firstClasses.keys()].sort((left, right) => left - right));

  const firstAt = CLASS_PAIR_HEADER + firstCount * secondCount * recordLength;
  const secondAt = firstAt + firstDefinition.length;
  const coverageAt = secondAt + secondDefinition.length;

  const table = new Uint8Array(coverageAt + coverage.length);
  const view = new DataView(table.buffer);
  view.setUint16(0, 2);
  view.setUint16(2, coverageAt);
  view.setUint16(4, formats[0]);
  view.setUint16(6, formats[1]);
  view.setUint16(8, firstAt);
  view.setUint16(10, secondAt);
  view.setUint16(12, firstCount);
  view.setUint16(14, secondCount);

  for (let first = 0; first < firstCount; first += 1) {
    for (let second = 0; second < secondCount; second += 1) {
      const at = CLASS_PAIR_HEADER + (first * secondCount + second) * recordLength;
      writeValue(view, at, formats[0], stated.values[first]?.[second] ?? 0);
      writeValue(view, at + valueLength(formats[0]), formats[1], 0);
    }
  }

  table.set(firstDefinition, firstAt);
  table.set(secondDefinition, secondAt);
  table.set(coverage, coverageAt);
  return table;
}

const LOOKUP_HEADER = 8;
const EXTENSION_LENGTH = 8;
const PAIR_ADJUSTMENT = 2;
const EXTENSION_POSITIONING = 9;

function extensionSubtable(subtable: Uint8Array): Uint8Array {
  const wrapper = new Uint8Array(EXTENSION_LENGTH + subtable.length);
  const view = new DataView(wrapper.buffer);
  view.setUint16(0, 1);
  view.setUint16(2, PAIR_ADJUSTMENT);
  view.setUint32(4, EXTENSION_LENGTH);
  wrapper.set(subtable, EXTENSION_LENGTH);
  return wrapper;
}

function oneLookup(subtable: Uint8Array, throughExtension: boolean): Uint8Array {
  const lookup = new Uint8Array(LOOKUP_HEADER + subtable.length);
  const view = new DataView(lookup.buffer);
  view.setUint16(0, throughExtension ? EXTENSION_POSITIONING : PAIR_ADJUSTMENT);
  view.setUint16(4, 1);
  view.setUint16(6, LOOKUP_HEADER);
  lookup.set(subtable, LOOKUP_HEADER);
  return lookup;
}

function lookupList(lookups: readonly Uint8Array[]): Uint8Array {
  const listHeader = 2 + lookups.length * 2;
  const list = new Uint8Array(listHeader + lookups.reduce((sum, each) => sum + each.length, 0));
  const view = new DataView(list.buffer);
  view.setUint16(0, lookups.length);

  let at = listHeader;
  lookups.forEach((lookup, index) => {
    view.setUint16(2 + index * 2, at);
    list.set(lookup, at);
    at += lookup.length;
  });
  return list;
}

// One default script whose default language wants the one feature, which is the
// least a face can state and still be shaped. The reader finds the lookups through
// the feature rather than through this, but a face states it and so does a fixture.
function scriptList(): Uint8Array {
  const list = new Uint8Array(20);
  const view = new DataView(list.buffer);
  view.setUint16(0, 1);
  list.set(tagBytes("DFLT"), 2);
  view.setUint16(6, 8);
  view.setUint16(8, 4);
  view.setUint16(14, 0xffff);
  view.setUint16(16, 1);
  return list;
}

function featureList(lookupCount: number): Uint8Array {
  const featureAt = 8;
  const list = new Uint8Array(featureAt + 4 + lookupCount * 2);
  const view = new DataView(list.buffer);
  view.setUint16(0, 1);
  list.set(tagBytes("kern"), 2);
  view.setUint16(6, featureAt);
  view.setUint16(featureAt + 2, lookupCount);
  for (let index = 0; index < lookupCount; index += 1) {
    view.setUint16(featureAt + 4 + index * 2, index);
  }
  return list;
}

const GPOS_HEADER = 10;

function gposTable(fixture: FontFixture, glyphs: readonly Glyph[]): Uint8Array {
  const formats = fixture.gposValueFormats ?? FIRST_GLYPH_ADVANCE;
  const pairs =
    fixture.gposClassPairs === undefined
      ? pairSetSubtable(glyphPairs(fixture.gposPairs ?? {}, glyphs), formats)
      : classPairSubtable(fixture.gposClassPairs, glyphs, formats);

  const throughExtension = fixture.gposThroughExtension === true;
  const lookups = [
    oneLookup(throughExtension ? extensionSubtable(pairs) : pairs, throughExtension),
  ];
  if (fixture.gposRightToLeftPairs !== undefined) {
    const other = pairSetSubtable(
      glyphPairs(fixture.gposRightToLeftPairs, glyphs),
      PLACED_AND_ADVANCED,
    );
    lookups.push(oneLookup(other, false));
  }

  const scripts = scriptList();
  const features = featureList(lookups.length);

  const header = new Uint8Array(GPOS_HEADER);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint16(4, GPOS_HEADER);
  view.setUint16(6, GPOS_HEADER + scripts.length);
  view.setUint16(8, GPOS_HEADER + scripts.length + features.length);

  const whole = concat([header, scripts, features, lookupList(lookups)]);

  return fixture.cutFromGpos === undefined
    ? whole
    : whole.subarray(0, whole.length - fixture.cutFromGpos);
}

const MATH_HEADER_LENGTH = 10;
const MATH_VALUE_LENGTH = 4;
const MATH_CONSTANTS_LENGTH = 8 + MATH_VALUE_CONSTANTS.length * MATH_VALUE_LENGTH + 2;
const GLYPH_INFO_LENGTH = 8;

function mathConstantsTable(stated: Partial<MathConstants>): Uint8Array {
  const table = new Uint8Array(MATH_CONSTANTS_LENGTH);
  const view = new DataView(table.buffer);

  view.setInt16(0, stated.scriptPercentScaleDown ?? 0);
  view.setInt16(2, stated.scriptScriptPercentScaleDown ?? 0);
  view.setUint16(4, stated.delimitedSubFormulaMinHeight ?? 0);
  view.setUint16(6, stated.displayOperatorMinHeight ?? 0);

  MATH_VALUE_CONSTANTS.forEach((name: MathValueConstant, index) => {
    view.setInt16(8 + index * MATH_VALUE_LENGTH, stated[name] ?? 0);
  });

  view.setInt16(MATH_CONSTANTS_LENGTH - 2, stated.radicalDegreeBottomRaisePercent ?? 0);
  return table;
}

// A coverage names the glyphs a table answers for, in the order it answers for
// them, which is why everything written beside one is sorted by glyph.
function coverageOf(glyphIds: readonly number[]): Uint8Array {
  const table = new Uint8Array(4 + glyphIds.length * 2);
  const view = new DataView(table.buffer);
  view.setUint16(0, 1);
  view.setUint16(2, glyphIds.length);
  glyphIds.forEach((glyph, index) => {
    view.setUint16(4 + index * 2, glyph);
  });
  return table;
}

const byGlyph = <T>(
  stated: Readonly<Record<string, T>>,
  glyphs: readonly Glyph[],
): readonly { readonly glyph: number; readonly stated: T }[] =>
  Object.entries(stated)
    .map(([character, value]) => ({ glyph: idOf(character.codePointAt(0), glyphs), stated: value }))
    .sort((left, right) => left.glyph - right.glyph);

function italicCorrectionsTable(
  stated: Readonly<Record<string, number>>,
  glyphs: readonly Glyph[],
): Uint8Array {
  const corrections = byGlyph(stated, glyphs);
  const coverage = coverageOf(corrections.map((each) => each.glyph));
  const header = 4 + corrections.length * MATH_VALUE_LENGTH;

  const table = new Uint8Array(header + coverage.length);
  const view = new DataView(table.buffer);
  view.setUint16(0, header);
  view.setUint16(2, corrections.length);
  corrections.forEach((each, index) => {
    view.setInt16(4 + index * MATH_VALUE_LENGTH, each.stated);
  });
  table.set(coverage, header);
  return table;
}

const ASSEMBLY_PART_LENGTH = 10;

function assemblyTable(pieces: MathPiecesFixture, glyphs: readonly Glyph[]): Uint8Array {
  const table = new Uint8Array(6 + pieces.parts.length * ASSEMBLY_PART_LENGTH);
  const view = new DataView(table.buffer);
  view.setInt16(0, pieces.italicCorrection ?? 0);
  view.setUint16(4, pieces.parts.length);

  pieces.parts.forEach((part, index) => {
    const at = 6 + index * ASSEMBLY_PART_LENGTH;
    view.setUint16(at, idOf(part.character.codePointAt(0), glyphs));
    view.setUint16(at + 2, part.startConnector);
    view.setUint16(at + 4, part.endConnector);
    view.setUint16(at + 6, part.fullAdvance);
    view.setUint16(at + 8, part.extender === true ? 1 : 0);
  });
  return table;
}

type Growth = {
  readonly glyph: number;
  readonly variants: readonly MathVariantFixture[];
  readonly pieces: MathPiecesFixture | undefined;
};

function constructionTable(growth: Growth, glyphs: readonly Glyph[]): Uint8Array {
  const header = 4 + growth.variants.length * 4;
  const assembly = growth.pieces === undefined ? null : assemblyTable(growth.pieces, glyphs);

  const table = new Uint8Array(header + (assembly?.length ?? 0));
  const view = new DataView(table.buffer);
  view.setUint16(0, assembly === null ? 0 : header);
  view.setUint16(2, growth.variants.length);
  growth.variants.forEach((variant, index) => {
    view.setUint16(4 + index * 4, idOf(variant.character.codePointAt(0), glyphs));
    view.setUint16(6 + index * 4, variant.measurement);
  });
  if (assembly !== null) table.set(assembly, header);
  return table;
}

function growthOf(
  variants: Readonly<Record<string, readonly MathVariantFixture[]>> | undefined,
  pieces: Readonly<Record<string, MathPiecesFixture>> | undefined,
  glyphs: readonly Glyph[],
): readonly Growth[] {
  const characters = [...new Set([...Object.keys(variants ?? {}), ...Object.keys(pieces ?? {})])];
  return characters
    .map((character) => ({
      glyph: idOf(character.codePointAt(0), glyphs),
      variants: variants?.[character] ?? [],
      pieces: pieces?.[character],
    }))
    .sort((left, right) => left.glyph - right.glyph);
}

// Every offset inside the variants table counts from the table itself, so the
// constructions are laid out after the two lists of offsets that name them and
// the coverages after those.
function mathVariantsTable(math: MathFixture, glyphs: readonly Glyph[]): Uint8Array {
  const taller = growthOf(math.tallerVariants, math.tallerPieces, glyphs);
  const wider = growthOf(math.widerVariants, undefined, glyphs);
  const constructions = [...taller, ...wider].map((each) => constructionTable(each, glyphs));

  const header = MATH_HEADER_LENGTH + (taller.length + wider.length) * 2;
  const parts: Uint8Array[] = [];
  const offsets: number[] = [];
  let at = header;
  for (const construction of constructions) {
    offsets.push(at);
    parts.push(construction);
    at += construction.length;
  }

  const tallCoverage = coverageOf(taller.map((each) => each.glyph));
  const wideCoverage = coverageOf(wider.map((each) => each.glyph));
  const tallCoverageAt = at;
  const wideCoverageAt = at + tallCoverage.length;

  const table = new Uint8Array(header);
  const view = new DataView(table.buffer);
  view.setUint16(0, math.minConnectorOverlap ?? 0);
  view.setUint16(2, taller.length === 0 ? 0 : tallCoverageAt);
  view.setUint16(4, wider.length === 0 ? 0 : wideCoverageAt);
  view.setUint16(6, taller.length);
  view.setUint16(8, wider.length);
  offsets.forEach((offset, index) => {
    view.setUint16(MATH_HEADER_LENGTH + index * 2, offset);
  });

  return concat([table, ...parts, tallCoverage, wideCoverage]);
}

function mathTable(math: MathFixture, glyphs: readonly Glyph[]): Uint8Array {
  const constants = mathConstantsTable(math.constants ?? {});
  const corrections = italicCorrectionsTable(math.italicCorrections ?? {}, glyphs);
  const glyphInfo = new Uint8Array(GLYPH_INFO_LENGTH + corrections.length);
  new DataView(glyphInfo.buffer).setUint16(0, GLYPH_INFO_LENGTH);
  glyphInfo.set(corrections, GLYPH_INFO_LENGTH);

  const variants = mathVariantsTable(math, glyphs);

  const header = new Uint8Array(MATH_HEADER_LENGTH);
  const view = new DataView(header.buffer);
  view.setUint16(0, 1);
  view.setUint16(4, MATH_HEADER_LENGTH);
  view.setUint16(6, MATH_HEADER_LENGTH + constants.length);
  view.setUint16(8, MATH_HEADER_LENGTH + constants.length + glyphInfo.length);

  return concat([header, constants, glyphInfo, variants]);
}

function tablesOf(fixture: FontFixture): readonly (readonly [string, Uint8Array])[] {
  const glyphs = glyphsOf(fixture);
  const tables: (readonly [string, Uint8Array])[] = [
    ["head", headTable(fixture)],
    ["hhea", hheaTable(fixture, glyphs)],
  ];

  if (fixture.panoseFamily !== undefined || fixture.panoseSerifStyle !== undefined) {
    tables.push(["OS/2", os2Table(fixture)]);
  }
  if (fixture.faceName !== undefined) {
    tables.push(["name", nameTable(fixture.faceName, fixture.familyName ?? fixture.faceName)]);
  }
  if (
    fixture.underlinePosition !== undefined ||
    fixture.underlineThickness !== undefined ||
    fixture.italicAngle !== undefined
  ) {
    tables.push(["post", postTable(fixture)]);
  }
  if (fixture.advances !== undefined) {
    tables.push(["cmap", cmapTable(fixture, glyphs)], ["hmtx", hmtxTable(fixture, glyphs)]);
  }
  if (fixture.kernPairs !== undefined) {
    tables.push(["kern", kernTable(fixture, glyphs)]);
  }
  if (fixture.gposPairs !== undefined || fixture.gposClassPairs !== undefined) {
    tables.push(["GPOS", gposTable(fixture, glyphs)]);
  }
  if (fixture.boxes !== undefined) {
    tables.push(...trueTypeOutlines(fixture, glyphs));
  }
  if (fixture.outlines !== undefined) {
    tables.push(["CFF ", cffTable(fixture, glyphs)]);
  }
  if (fixture.math !== undefined) {
    const written = mathTable(fixture.math, glyphs);
    const cut = fixture.cutFromMath ?? 0;
    tables.push(["MATH", cut === 0 ? written : written.subarray(0, written.length - cut)]);
  }

  return tables.filter(([tag]) => tag !== fixture.omit);
}

const tagBytes = (tag: string): Uint8Array =>
  Uint8Array.from([0, 1, 2, 3].map((index) => tag.charCodeAt(index)));

const padded = (length: number): number => (length + 3) & ~3;

export function buildSfnt(fixture: FontFixture): Uint8Array {
  const tables = tablesOf(fixture);
  const directoryLength = 12 + tables.length * 16;
  const total = directoryLength + tables.reduce((sum, [, data]) => sum + padded(data.length), 0);

  const out = new Uint8Array(total);
  const view = new DataView(out.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint16(4, tables.length);

  let offset = directoryLength;
  tables.forEach(([tag, data], index) => {
    const record = 12 + index * 16;
    out.set(tagBytes(tag), record);
    view.setUint32(record + 8, offset);
    view.setUint32(record + 12, data.length);
    out.set(data, offset);
    offset += padded(data.length);
  });

  return out;
}

export function buildWoff(fixture: FontFixture, compress = false): Uint8Array {
  const tables = tablesOf(fixture).map(([tag, data]) => {
    const body = compress ? zlibSync(data) : data;
    return { tag, data, body: body.length < data.length ? body : data };
  });

  const directoryLength = 44 + tables.length * 20;
  const total = directoryLength + tables.reduce((sum, t) => sum + padded(t.body.length), 0);

  const out = new Uint8Array(total);
  const view = new DataView(out.buffer);
  out.set(tagBytes("wOFF"), 0);
  view.setUint32(4, 0x00010000);
  view.setUint32(8, total);
  view.setUint16(12, tables.length);

  let offset = directoryLength;
  tables.forEach((table, index) => {
    const record = 44 + index * 20;
    out.set(tagBytes(table.tag), record);
    view.setUint32(record + 4, offset);
    view.setUint32(record + 8, table.body.length);
    view.setUint32(record + 12, table.data.length);
    out.set(table.body, offset);
    offset += padded(table.body.length);
  });

  return out;
}

// A collection of whole sfnts behind one header, each moved to where it sits in
// the file: a table record in a collection counts from the start of the file
// rather than from the face's own directory.
export function buildCollection(fixtures: readonly FontFixture[]): Uint8Array {
  const members = fixtures.map(buildSfnt);
  const header = 12 + members.length * 4;
  const bases: number[] = [];
  let offset = header;
  for (const member of members) {
    bases.push(offset);
    offset += padded(member.length);
  }

  const out = new Uint8Array(offset);
  const view = new DataView(out.buffer);
  out.set(tagBytes("ttcf"), 0);
  view.setUint32(4, 0x00010000);
  view.setUint32(8, members.length);

  members.forEach((member, index) => {
    const base = bases[index] ?? 0;
    view.setUint32(12 + index * 4, base);
    out.set(member, base);

    const inner = new DataView(out.buffer, base);
    const count = inner.getUint16(4);
    for (let table = 0; table < count; table += 1) {
      const record = 12 + table * 16;
      inner.setUint32(record + 8, inner.getUint32(record + 8) + base);
    }
  });

  return out;
}

export const buildWoff2 = (): Uint8Array => {
  const out = new Uint8Array(48);
  out.set(tagBytes("wOF2"), 0);
  return out;
};

const MEASURABLE =
  " !\"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`" +
  "abcdefghijklmnopqrstuvwxyz{|}~";

export type FaceFixture = {
  readonly name: string;
  readonly metrics: FontMetrics;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly advance?: number;
  readonly characters?: string;
  readonly subtables?: FontFixture["subtables"];
  readonly notdefAdvance?: number;
  // Whether the face draws its letters without serifs, which it says through its
  // PANOSE classification as a real face does rather than being told.
  readonly sansSerif?: boolean;
  // What each pair of its characters moves, stated in the legacy table. A face
  // that states none supplies no kerning at all, as most faces do.
  readonly kernPairs?: KerningPairs;
  // What a character advances where that is not the one width every other
  // character of the face has. A character named here and nowhere else is added
  // to the face, which is how a face is given a letter outside the measurable
  // ASCII the default states.
  readonly advances?: Readonly<Record<string, number>>;
  // What each character draws, written as a TrueType outline, and what the face
  // says about setting mathematics. A face stating neither supplies neither, as a
  // face read off a file that states neither does.
  readonly boxes?: Readonly<Record<string, InkFixture>>;
  readonly math?: MathFixture;
};

// The plainest PANOSE classification of each kind: a Latin text face of normal
// sans, and one whose serifs are coves.
const SANS_SERIF = { panoseFamily: 2, panoseSerifStyle: 11 };
const SERIF = { panoseFamily: 2, panoseSerifStyle: 2 };

// A face whose every glyph is the same width, so a test can count characters
// instead of consulting a real font's widths.
export function buildFace(fixture: FaceFixture): SuppliedFace {
  const advance = fixture.advance ?? fixture.metrics.unitsPerEm / 2;
  // Every character the face measures at the one width, and then whatever the
  // fixture states a width of its own for, which may be a character none of the
  // others is: a face has to map a letter before it can draw it or grow it.
  const advances = {
    ...Object.fromEntries(
      Array.from(fixture.characters ?? MEASURABLE, (character) => [character, advance]),
    ),
    ...fixture.advances,
  };

  // A face mapping a character past the basic plane states the wider cmap, as a
  // real one does: Word draws a math run in the Mathematical Italic block, which
  // is all beyond it.
  const beyondBasicPlane = Object.keys(advances).some(
    (character) => (character.codePointAt(0) ?? 0) > 0xffff,
  );

  const file = buildSfnt({
    ...fixture.metrics,
    ...(beyondBasicPlane ? { cmapFormat: 12 as const } : {}),
    ...(fixture.subtables === undefined ? {} : { subtables: fixture.subtables }),
    ...(fixture.notdefAdvance === undefined ? {} : { notdefAdvance: fixture.notdefAdvance }),
    ...(fixture.sansSerif === undefined ? {} : fixture.sansSerif ? SANS_SERIF : SERIF),
    ...(fixture.kernPairs === undefined ? {} : { kernPairs: fixture.kernPairs }),
    ...(fixture.boxes === undefined ? {} : { boxes: fixture.boxes }),
    ...(fixture.math === undefined ? {} : { math: fixture.math }),
    advances,
  });
  const read = readFontFile(file);

  return {
    name: fixture.name,
    bold: fixture.bold ?? false,
    italic: fixture.italic ?? false,
    metrics: fixture.metrics,
    advances: read.advances,
    sansSerif: read.sansSerif,
    ...(read.kerning.kind === "kerning" ? { kerning: read.kerning } : {}),
    ...(read.ink.kind === "ink" ? { ink: read.ink } : {}),
    ...(read.math.kind === "math" ? { math: read.math } : {}),
  };
}
