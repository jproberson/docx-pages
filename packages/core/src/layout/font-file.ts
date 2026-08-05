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
};

const AT = "core/layout/font-file.readFontMetrics";

const HEAD = "head";
const HHEA = "hhea";

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

// A collection holds several faces in one file, each with a table directory of its
// own whose records still count from the start of the file. Word ships Cambria and
// the Yu Gothic family this way and no other, so a name that resolves to one is
// unreadable without this. The first face is the one the file is named for; the
// rest are the styles and the maths cut, which are asked for by their own names
// and would need a file of their own here.
function firstOfCollection(bytes: Uint8Array): number {
  const view = viewOf(bytes);
  if (bytes.byteLength < 16 || view.getUint32(8) === 0) {
    throw unreadable("the font collection holds no faces", bytes.byteLength);
  }
  return view.getUint32(12);
}

export function readFontFile(bytes: Uint8Array): ReadFontFileResult {
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
      : readSfntTables(bytes, isCollection ? firstOfCollection(bytes) : 0);

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
  };
}

export function readFontMetrics(bytes: Uint8Array): ReadFontMetricsResult {
  const { format, metrics } = readFontFile(bytes);
  return { format, metrics };
}
