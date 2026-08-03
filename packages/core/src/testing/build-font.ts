import { zlibSync } from "fflate";

export type FontFixture = {
  readonly unitsPerEm: number;
  readonly ascender: number;
  readonly descender: number;
  readonly lineGap: number;
  readonly omit?: "head" | "hhea";
};

const HEAD_LENGTH = 54;
const HHEA_LENGTH = 36;

function headTable(fixture: FontFixture): Uint8Array {
  const table = new Uint8Array(HEAD_LENGTH);
  const view = new DataView(table.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint32(12, 0x5f0f3cf5);
  view.setUint16(18, fixture.unitsPerEm);
  return table;
}

function hheaTable(fixture: FontFixture): Uint8Array {
  const table = new Uint8Array(HHEA_LENGTH);
  const view = new DataView(table.buffer);
  view.setUint32(0, 0x00010000);
  view.setInt16(4, fixture.ascender);
  view.setInt16(6, fixture.descender);
  view.setInt16(8, fixture.lineGap);
  return table;
}

const tablesOf = (fixture: FontFixture): readonly (readonly [string, Uint8Array])[] =>
  [["head", headTable(fixture)] as const, ["hhea", hheaTable(fixture)] as const].filter(
    ([tag]) => tag !== fixture.omit,
  );

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

export const buildWoff2 = (): Uint8Array => {
  const out = new Uint8Array(48);
  out.set(tagBytes("wOF2"), 0);
  return out;
};
