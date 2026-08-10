import { zlibSync } from "fflate";

import { readFontFile } from "../layout/font-file.js";
import type { FontMetrics, SuppliedFace } from "../layout/font-metrics.js";

export type FontFixture = {
  readonly unitsPerEm: number;
  readonly ascender: number;
  readonly descender: number;
  readonly lineGap: number;
  readonly advances?: Readonly<Record<string, number>>;
  readonly cmapFormat?: 4 | 6 | 12;
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
  readonly omit?: "head" | "hhea" | "cmap" | "hmtx";
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

function headTable(fixture: FontFixture): Uint8Array {
  const table = new Uint8Array(HEAD_LENGTH);
  const view = new DataView(table.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint32(12, 0x5f0f3cf5);
  view.setUint16(18, fixture.unitsPerEm);
  view.setUint16(MAC_STYLE_AT, (fixture.bold === true ? 1 : 0) | (fixture.italic === true ? 2 : 0));
  return table;
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

function cmapFormat6(): Uint8Array {
  const table = new Uint8Array(10);
  new DataView(table.buffer).setUint16(0, 6);
  return table;
}

// Every encoding the fixture declares points at the one subtable, since what a
// fixture has to say here is which encodings a face offers rather than what each
// of them maps.
function cmapTable(fixture: FontFixture, glyphs: readonly Glyph[]): Uint8Array {
  const format = fixture.cmapFormat ?? 4;
  const subtable =
    format === 12 ? cmapFormat12(glyphs) : format === 6 ? cmapFormat6() : cmapFormat4(glyphs);
  const declared = fixture.subtables ?? ["unicode"];

  const at = 4 + declared.length * 8;
  const table = new Uint8Array(at + subtable.length);
  const view = new DataView(table.buffer);
  view.setUint16(2, declared.length);

  declared.forEach((encoding, index) => {
    const record = 4 + index * 8;
    view.setUint16(record, 3);
    view.setUint16(record + 2, encoding === "symbol" ? 0 : format === 12 ? 10 : 1);
    view.setUint32(record + 4, at);
  });

  table.set(subtable, at);
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
  if (fixture.advances !== undefined) {
    tables.push(["cmap", cmapTable(fixture, glyphs)], ["hmtx", hmtxTable(fixture, glyphs)]);
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
};

// The plainest PANOSE classification of each kind: a Latin text face of normal
// sans, and one whose serifs are coves.
const SANS_SERIF = { panoseFamily: 2, panoseSerifStyle: 11 };
const SERIF = { panoseFamily: 2, panoseSerifStyle: 2 };

// A face whose every glyph is the same width, so a test can count characters
// instead of consulting a real font's widths.
export function buildFace(fixture: FaceFixture): SuppliedFace {
  const advance = fixture.advance ?? fixture.metrics.unitsPerEm / 2;
  const file = buildSfnt({
    ...fixture.metrics,
    ...(fixture.subtables === undefined ? {} : { subtables: fixture.subtables }),
    ...(fixture.notdefAdvance === undefined ? {} : { notdefAdvance: fixture.notdefAdvance }),
    ...(fixture.sansSerif === undefined ? {} : fixture.sansSerif ? SANS_SERIF : SERIF),
    advances: Object.fromEntries(
      Array.from(fixture.characters ?? MEASURABLE, (character) => [character, advance]),
    ),
  });
  const read = readFontFile(file);

  return {
    name: fixture.name,
    bold: fixture.bold ?? false,
    italic: fixture.italic ?? false,
    metrics: fixture.metrics,
    advances: read.advances,
    sansSerif: read.sansSerif,
  };
}
