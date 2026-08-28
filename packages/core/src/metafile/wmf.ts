import { zlibSync } from "fflate";

/**
 * The older metafile, which Word writes beside the one this project plays.
 *
 * **It is not the same job as an EMF.** An EMF records the drawing, and
 * `readMetafilePicture` plays its lines and text into shapes. Every WMF in the
 * corpus records one thing instead: a bitmap, blitted into the frame. So this
 * reads the bitmap out and hands it on as a png, which both renderers already
 * draw, rather than playing anything.
 *
 * Read on 2026-08-14 over the eight WMFs the 718 documents hold, in five of them:
 * each is a placeable metafile whose drawing records are `DibStretchBlt` and
 * `DibBitBlt`, carrying one uncompressed 24-bit device-independent bitmap.
 *
 * **Whatever this cannot read is refused rather than approximated**, and the
 * report goes on naming the picture undrawable, which is what it is for.
 */

// The word every placeable metafile opens with, which is Aldus's own and not part
// of the format Windows defined: a placeable file carries its size in twips in
// front of the metafile proper.
const PLACEABLE_KEY = 0x9ac6cdd7;
const PLACEABLE_BYTES = 22;
const HEADER_BYTES = 18;

// The records that carry a bitmap. Each states its own parameters before the
// bitmap, so where the bitmap starts is the record's length less the bitmap's own.
const DIB_STRETCH_BLT = 0x0b41;
const DIB_BIT_BLT = 0x0940;
const SET_DIB_TO_DEVICE = 0x0d33;
const STRETCH_DIB = 0x0f43;

const INFO_HEADER_BYTES = 40;
const UNCOMPRESSED = 0;

export type DecodedBitmap = {
  readonly widthPixels: number;
  readonly heightPixels: number;
  // Three bytes a pixel, left to right and top to bottom, which is the order a png
  // is written in and the opposite of the order a bitmap is stored in.
  readonly rgb: Uint8Array;
};

/**
 * The bitmap a metafile blits, or nothing where it draws anything else.
 *
 * The records are walked rather than the first one taken: the corpus files open
 * with an escape carrying the printer's own comment, and the bitmap stands well
 * inside them.
 */
export function readMetafileBitmap(bytes: Uint8Array): DecodedBitmap | null {
  if (bytes.byteLength < PLACEABLE_BYTES + HEADER_BYTES) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  let at = view.getUint32(0, true) === PLACEABLE_KEY ? PLACEABLE_BYTES : 0;
  // The metafile's own header, whose length is stated in words and is 9 of them.
  const headerWords = view.getUint16(at + 2, true);
  if (headerWords * 2 !== HEADER_BYTES) return null;
  at += HEADER_BYTES;

  while (at + 6 <= bytes.byteLength) {
    const size = view.getUint32(at, true) * 2;
    const record = view.getUint16(at + 4, true);
    if (size < 6 || at + size > bytes.byteLength) return null;
    if (record === 0) return null;

    if (
      record === DIB_STRETCH_BLT ||
      record === DIB_BIT_BLT ||
      record === SET_DIB_TO_DEVICE ||
      record === STRETCH_DIB
    ) {
      const found = bitmapIn(view, bytes, at + 6, at + size);
      if (found !== null) return found;
    }
    at += size;
  }

  return null;
}

// Where the bitmap starts inside a record, found by the header it opens with
// rather than by counting the record's own parameters: those differ per record and
// per whether the record carries a raster operation, and the header is 40 bytes
// stating 40.
function bitmapIn(
  view: DataView,
  bytes: Uint8Array,
  from: number,
  to: number,
): DecodedBitmap | null {
  for (let at = from; at + INFO_HEADER_BYTES <= to; at += 2) {
    if (view.getUint32(at, true) !== INFO_HEADER_BYTES) continue;
    const decoded = decodeBitmap(view, bytes, at, to);
    if (decoded !== null) return decoded;
  }
  return null;
}

function decodeBitmap(
  view: DataView,
  bytes: Uint8Array,
  at: number,
  to: number,
): DecodedBitmap | null {
  const widthPixels = view.getInt32(at + 4, true);
  // A negative height is a bitmap stored top down, which is the one thing about
  // the order that is not fixed.
  const statedHeight = view.getInt32(at + 8, true);
  const heightPixels = Math.abs(statedHeight);
  const bits = view.getUint16(at + 14, true);
  const compression = view.getUint32(at + 16, true);
  const statedColours = view.getUint32(at + 32, true);

  if (widthPixels <= 0 || heightPixels <= 0 || compression !== UNCOMPRESSED) return null;
  if (bits !== 24 && bits !== 32 && bits !== 8) return null;

  const colours = bits === 8 ? (statedColours === 0 ? 256 : statedColours) : 0;
  const palette = at + INFO_HEADER_BYTES;
  const pixels = palette + colours * 4;
  // A row is padded out to a whole number of four-byte words, whatever its depth.
  const stride = Math.ceil((widthPixels * bits) / 32) * 4;
  if (pixels + stride * heightPixels > to) return null;

  const rgb = new Uint8Array(widthPixels * heightPixels * 3);
  for (let row = 0; row < heightPixels; row += 1) {
    // Stored bottom up unless the height said otherwise, so the last row read is
    // the first row drawn.
    const source = pixels + (statedHeight < 0 ? row : heightPixels - 1 - row) * stride;
    let out = row * widthPixels * 3;
    for (let column = 0; column < widthPixels; column += 1) {
      let blue: number;
      let green: number;
      let red: number;
      if (bits === 8) {
        const entry = palette + (bytes[source + column] ?? 0) * 4;
        blue = bytes[entry] ?? 0;
        green = bytes[entry + 1] ?? 0;
        red = bytes[entry + 2] ?? 0;
      } else {
        const from = source + column * (bits / 8);
        blue = bytes[from] ?? 0;
        green = bytes[from + 1] ?? 0;
        red = bytes[from + 2] ?? 0;
      }
      rgb[out] = red;
      rgb[out + 1] = green;
      rgb[out + 2] = blue;
      out += 3;
    }
  }

  return { widthPixels, heightPixels, rgb };
}

/**
 * The same bitmap written out as a png, which is what both renderers draw.
 *
 * Neither of them is taught anything new by this: the viewer already puts a png in
 * a data url and the pdf writer already reads one into an image, so a metafile
 * that turns out to be a bitmap becomes a picture like any other.
 */
export function pngOfBitmap(bitmap: DecodedBitmap): Uint8Array {
  const { widthPixels, heightPixels, rgb } = bitmap;
  // Every row of a png opens with the filter it was written under, and nought is
  // the filter that leaves the bytes alone.
  const raw = new Uint8Array(heightPixels * (widthPixels * 3 + 1));
  for (let row = 0; row < heightPixels; row += 1) {
    raw.set(
      rgb.subarray(row * widthPixels * 3, (row + 1) * widthPixels * 3),
      row * (widthPixels * 3 + 1) + 1,
    );
  }

  const header = new Uint8Array(13);
  const headerView = new DataView(header.buffer);
  headerView.setUint32(0, widthPixels);
  headerView.setUint32(4, heightPixels);
  header[8] = 8;
  // Colour type 2 is three channels with no alpha, which is what a bitmap this
  // deep holds.
  header[9] = 2;

  return concat([
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", header),
    chunk("IDAT", zlibSync(raw)),
    chunk("IEND", new Uint8Array(0)),
  ]);
}

function chunk(name: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(data.byteLength + 12);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.byteLength);
  for (let at = 0; at < 4; at += 1) out[4 + at] = name.charCodeAt(at);
  out.set(data, 8);
  view.setUint32(out.byteLength - 4, crc32(out.subarray(4, out.byteLength - 4)));
  return out;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let at = 0; at < 256; at += 1) {
    let value = at;
    for (let round = 0; round < 8; round += 1) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[at] = value >>> 0;
  }
  return table;
})();

function crc32(bytes: Uint8Array): number {
  let value = 0xffffffff;
  for (const byte of bytes) value = (CRC_TABLE[(value ^ byte) & 0xff] ?? 0) ^ (value >>> 8);
  return (value ^ 0xffffffff) >>> 0;
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((sum, each) => sum + each.byteLength, 0));
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.byteLength;
  }
  return out;
}

// The picture a WMF holds, as a png, or nothing where it holds something this
// cannot read.
export function pngFromMetafile(bytes: Uint8Array): Uint8Array | null {
  const bitmap = readMetafileBitmap(bytes);
  return bitmap === null ? null : pngOfBitmap(bitmap);
}

// The records of an EMF that blit a bitmap rather than drawing anything.
const EMF_BIT_BLT = 76;
const EMF_STRETCH_BLT = 77;
const EMF_SET_DIB_TO_DEVICE = 80;
const EMF_STRETCH_DIB_ITS = 81;
const EMF_RECORD_HEADER_BYTES = 8;

/**
 * The bitmap an **enhanced** metafile blits, or nothing where it draws anything else.
 *
 * `readMetafilePicture` plays an EMF's lines and text into shapes and refuses any record
 * it does not know, which is right for a drawing and wrong for the other kind of EMF
 * Word writes: one whose whole content is a photograph, wrapped as a single
 * `EMR_STRETCHDIBITS` between a header and an end. That one played as nothing at all,
 * and `38b5a0336fda` drew its second picture nowhere while saying nothing about it.
 *
 * So this is the WMF path applied to the newer format, and for the same reason: the
 * bitmap is read out and handed on as a png, which both renderers already draw. The DIB
 * inside the record is the very same structure, so `bitmapIn` finds it the same way, by
 * the 40-byte header it opens with rather than by counting parameters.
 *
 * **A metafile this cannot read is still refused rather than approximated**, and
 * `readMetafilePicture` is still asked first, so an EMF that really does record a drawing
 * is played as one.
 */
export function readEnhancedMetafileBitmap(bytes: Uint8Array): DecodedBitmap | null {
  if (bytes.byteLength < EMF_RECORD_HEADER_BYTES) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  let at = 0;
  while (at + EMF_RECORD_HEADER_BYTES <= bytes.byteLength) {
    const type = view.getUint32(at, true);
    const length = view.getUint32(at + 4, true);
    if (length < EMF_RECORD_HEADER_BYTES || length % 4 !== 0) return null;
    if (at + length > bytes.byteLength) return null;

    if (
      type === EMF_BIT_BLT ||
      type === EMF_STRETCH_BLT ||
      type === EMF_SET_DIB_TO_DEVICE ||
      type === EMF_STRETCH_DIB_ITS
    ) {
      const found = bitmapIn(view, bytes, at + EMF_RECORD_HEADER_BYTES, at + length);
      if (found !== null) return found;
    }
    at += length;
  }

  return null;
}

export function pngFromEnhancedMetafile(bytes: Uint8Array): Uint8Array | null {
  const bitmap = readEnhancedMetafileBitmap(bytes);
  return bitmap === null ? null : pngOfBitmap(bitmap);
}
