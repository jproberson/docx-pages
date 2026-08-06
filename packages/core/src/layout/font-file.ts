import { unzlibSync } from "fflate";

import { DocxPagesError } from "../errors.js";
import type { AdvanceTable, FontMetrics } from "./font-metrics.js";
import { readAdvanceTable } from "./glyphs.js";

export type FontFileFormat = "sfnt" | "woff";

export type ReadFontMetricsResult = {
  readonly format: FontFileFormat;
  readonly metrics: FontMetrics;
};

export type ReadFontFileResult = ReadFontMetricsResult & {
  readonly advances: AdvanceTable;
  // Whether the face draws its letters without serifs, which is the half of the
  // question the file itself answers about which face Word borrows a character
  // from. See `sansSerif` in `font-metrics.ts` for what turns on it.
  readonly sansSerif: boolean;
};

const AT = "core/layout/font-file.readFontMetrics";

const HEAD = "head";
const HHEA = "hhea";
const NAME = "name";
const OS2 = "OS/2";

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

function namesOf(table: Uint8Array): readonly string[] {
  if (table.byteLength < 6) return [];
  const view = viewOf(table);
  const count = view.getUint16(2);
  const stringsAt = view.getUint16(4);

  const found: string[] = [];
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
    found.push(name);
  }
  return found;
}

const normalise = (name: string): string => name.trim().toLowerCase();

// A collection holds several faces in one file, each with a table directory of its
// own whose records still count from the start of the file. Word ships Cambria and
// the Yu Gothic family this way and no other, and the face a document wants may
// not be the one the file is named for: Cambria Math, which Word borrows a maths
// letter and a hyphen from, is the second face of `Cambria.ttc`. It is asked for
// by name rather than by index, since where a face sits in the file is the
// business of whoever shipped it.
function faceOfCollection(bytes: Uint8Array, wanted: string | undefined): number {
  const view = viewOf(bytes);
  const count = bytes.byteLength < 16 ? 0 : view.getUint32(8);
  const offsets: number[] = [];
  for (let index = 0; index < count && 16 + index * 4 <= bytes.byteLength; index += 1) {
    offsets.push(view.getUint32(12 + index * 4));
  }

  const first = offsets[0];
  if (first === undefined) throw unreadable("the font collection holds no faces", bytes.byteLength);
  if (wanted === undefined) return first;

  const key = normalise(wanted);
  for (const at of offsets) {
    const table = readSfntTables(bytes, at).get(NAME);
    if (table !== undefined && namesOf(table).some((each) => normalise(each) === key)) return at;
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

/**
 * Reads what laying text out in a face needs: its vertical metrics, the advance of
 * every glyph it maps, and whether it draws its letters without serifs.
 *
 * `faceName` picks one face out of a collection, which holds several; without one
 * the face the file is named for answers, and a file holding one face ignores it.
 */
export function readFontFile(bytes: Uint8Array, faceName?: string): ReadFontFileResult {
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
  const tables =
    format === "woff"
      ? readWoffTables(bytes)
      : readSfntTables(bytes, isCollection ? faceOfCollection(bytes, faceName) : 0);

  const head = requireTable(tables, HEAD, 20);
  const hhea = requireTable(tables, HHEA, 10);
  const headView = viewOf(head);
  const hheaView = viewOf(hhea);

  const metricCount =
    hhea.byteLength >= HHEA_METRIC_COUNT_AT + 2 ? hheaView.getUint16(HHEA_METRIC_COUNT_AT) : 0;

  return {
    format,
    metrics: {
      unitsPerEm: headView.getUint16(18),
      ascender: hheaView.getInt16(4),
      descender: hheaView.getInt16(6),
      lineGap: hheaView.getInt16(8),
    },
    advances: readAdvanceTable(tables, metricCount),
    sansSerif: readsSansSerif(tables),
  };
}

export function readFontMetrics(bytes: Uint8Array): ReadFontMetricsResult {
  const { format, metrics } = readFontFile(bytes);
  return { format, metrics };
}
