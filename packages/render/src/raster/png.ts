import { deflateSync, inflateSync } from "node:zlib";

// Reading back a drawing of a page.
//
// Both rasterisers write png and nothing here can decode one: Chrome writes no
// other format, and taking a dependency to read four bytes a pixel would be a
// larger thing than the reader. What is needed is the whole image at eight bits a
// channel and nothing else, so interlacing and the deep bit depths are refused
// rather than guessed at. `build-png.ts` writes one by hand to test this.

export type RasterImage = {
  readonly width: number;
  readonly height: number;
  // Four bytes a pixel, red green blue alpha, row by row from the top.
  readonly pixels: Uint8Array;
};

const SIGNATURE = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]);

const CHANNELS = new Map<number, number>([
  [0, 1], // grey
  [2, 3], // red green blue
  [3, 1], // an index into the palette
  [4, 2], // grey and alpha
  [6, 4], // red green blue alpha
]);

type Chunks = {
  readonly header: DataView;
  readonly palette: Uint8Array | null;
  readonly transparency: Uint8Array | null;
  readonly data: Uint8Array;
};

function chunksOf(bytes: Uint8Array): Chunks {
  for (const [at, byte] of SIGNATURE.entries()) {
    if (bytes[at] !== byte) throw new Error("not a png");
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let header: DataView | null = null;
  let palette: Uint8Array | null = null;
  let transparency: Uint8Array | null = null;
  const parts: Uint8Array[] = [];

  for (let at = SIGNATURE.length; at + 8 <= bytes.length;) {
    const length = view.getUint32(at);
    const kind = String.fromCharCode(...bytes.subarray(at + 4, at + 8));
    const body = bytes.subarray(at + 8, at + 8 + length);
    if (kind === "IHDR") header = new DataView(body.buffer, body.byteOffset, body.byteLength);
    if (kind === "PLTE") palette = body;
    if (kind === "tRNS") transparency = body;
    if (kind === "IDAT") parts.push(body);
    if (kind === "IEND") break;
    at += 12 + length;
  }

  if (header === null) throw new Error("png with no header");

  let length = 0;
  for (const part of parts) length += part.length;
  const data = new Uint8Array(length);
  let written = 0;
  for (const part of parts) {
    data.set(part, written);
    written += part.length;
  }

  return { header, palette, transparency, data };
}

const paeth = (left: number, above: number, corner: number): number => {
  const estimate = left + above - corner;
  const toLeft = Math.abs(estimate - left);
  const toAbove = Math.abs(estimate - above);
  const toCorner = Math.abs(estimate - corner);
  if (toLeft <= toAbove && toLeft <= toCorner) return left;
  return toAbove <= toCorner ? above : corner;
};

// Each row is written under one of five filters, every one of them a difference
// from what has already been undone: the row above, the pixel to the left, or
// both. So this walks forward and can undo them in place.
function unfilter(rows: Uint8Array, width: number, height: number, perPixel: number): Uint8Array {
  const stride = width * perPixel;
  const out = new Uint8Array(stride * height);

  for (let row = 0; row < height; row += 1) {
    const filter = rows[row * (stride + 1)];
    const from = row * (stride + 1) + 1;
    const to = row * stride;

    for (let at = 0; at < stride; at += 1) {
      const raw = rows[from + at] ?? 0;
      const left = at >= perPixel ? (out[to + at - perPixel] ?? 0) : 0;
      const above = row > 0 ? (out[to - stride + at] ?? 0) : 0;
      const corner = row > 0 && at >= perPixel ? (out[to - stride + at - perPixel] ?? 0) : 0;

      const undone =
        filter === 0
          ? raw
          : filter === 1
            ? raw + left
            : filter === 2
              ? raw + above
              : filter === 3
                ? raw + ((left + above) >> 1)
                : filter === 4
                  ? raw + paeth(left, above, corner)
                  : null;

      if (undone === null) throw new Error(`png row filter ${String(filter)}`);
      out[to + at] = undone & 0xff;
    }
  }

  return out;
}

export function readPng(bytes: Uint8Array): RasterImage {
  const { header, palette, transparency, data } = chunksOf(bytes);

  const width = header.getUint32(0);
  const height = header.getUint32(4);
  const depth = header.getUint8(8);
  const colour = header.getUint8(9);
  const interlaced = header.getUint8(12);

  if (depth !== 8) throw new Error(`png at ${String(depth)} bits a channel`);
  if (interlaced !== 0) throw new Error("interlaced png");

  const channels = CHANNELS.get(colour);
  if (channels === undefined) throw new Error(`png colour type ${String(colour)}`);

  const raw = unfilter(inflateSync(data), width, height, channels);
  const pixels = new Uint8Array(width * height * 4);

  for (let at = 0; at < width * height; at += 1) {
    const from = at * channels;
    const to = at * 4;
    let red: number;
    let green: number;
    let blue: number;
    let alpha = 255;

    if (colour === 3) {
      const index = raw[from] ?? 0;
      if (palette === null) throw new Error("png indexed with no palette");
      red = palette[index * 3] ?? 0;
      green = palette[index * 3 + 1] ?? 0;
      blue = palette[index * 3 + 2] ?? 0;
      alpha = transparency?.[index] ?? 255;
    } else if (colour === 0 || colour === 4) {
      red = raw[from] ?? 0;
      green = red;
      blue = red;
      if (colour === 4) alpha = raw[from + 1] ?? 255;
    } else {
      red = raw[from] ?? 0;
      green = raw[from + 1] ?? 0;
      blue = raw[from + 2] ?? 0;
      if (colour === 6) alpha = raw[from + 3] ?? 255;
    }

    pixels[to] = red;
    pixels[to + 1] = green;
    pixels[to + 2] = blue;
    pixels[to + 3] = alpha;
  }

  return { width, height, pixels };
}

const CRC = (() => {
  const table = new Uint32Array(256);
  for (let at = 0; at < 256; at += 1) {
    let value = at;
    for (let bit = 0; bit < 8; bit += 1) {
      value = (value & 1) === 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[at] = value;
  }
  return table;
})();

function chunk(kind: string, body: Uint8Array): Uint8Array {
  const named = new Uint8Array(4 + body.length);
  for (let at = 0; at < 4; at += 1) named[at] = kind.charCodeAt(at);
  named.set(body, 4);

  let crc = 0xffffffff;
  for (const byte of named) crc = (CRC[(crc ^ byte) & 0xff] ?? 0) ^ (crc >>> 8);

  const out = new Uint8Array(12 + body.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, body.length);
  out.set(named, 4);
  view.setUint32(8 + body.length, (crc ^ 0xffffffff) >>> 0);
  return out;
}

/**
 * A png of rows already filtered, each row's filter written in front of it. The
 * caller decides the filters, which is what lets the reader's undoing of them be
 * tested against rows this project chose.
 */
export function pngOf(width: number, height: number, colour: number, rows: Uint8Array): Uint8Array {
  const header = new Uint8Array(13);
  const view = new DataView(header.buffer);
  view.setUint32(0, width);
  view.setUint32(4, height);
  header[8] = 8;
  header[9] = colour;

  const parts = [
    SIGNATURE,
    chunk("IHDR", header),
    chunk("IDAT", new Uint8Array(deflateSync(rows))),
    chunk("IEND", new Uint8Array(0)),
  ];

  const out = new Uint8Array(parts.reduce((total, each) => total + each.length, 0));
  let written = 0;
  for (const part of parts) {
    out.set(part, written);
    written += part.length;
  }
  return out;
}

// A drawing back out again, so that what a comparison is reading can be looked at.
export function writePng(image: RasterImage): Uint8Array {
  const stride = image.width * 3;
  const rows = new Uint8Array((stride + 1) * image.height);

  for (let row = 0; row < image.height; row += 1) {
    for (let column = 0; column < image.width; column += 1) {
      const from = (row * image.width + column) * 4;
      const to = row * (stride + 1) + 1 + column * 3;
      for (let channel = 0; channel < 3; channel += 1) {
        rows[to + channel] = image.pixels[from + channel] ?? 255;
      }
    }
  }

  return pngOf(image.width, image.height, 2, rows);
}
