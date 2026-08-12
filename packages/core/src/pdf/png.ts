import { unzlibSync } from "fflate";

// Enough of a png to write one into a pdf, which is less than playing one back:
// nothing here draws pixels, it only reads the header and gathers the pixel data
// still compressed.
//
// **A pdf and a png compress their pixels the same way.** Both deflate, and a
// pdf's predictor 15 is the png filtering, byte for byte: a filter number in front
// of every row and the same five filters behind it. So a png with no alpha needs
// no decoding at all to become a pdf image, and its bytes are carried across
// untouched. Only alpha forces the pixels open, because a png interleaves it with
// the colour and a pdf keeps it in a separate image.

const SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

const IHDR_LENGTH = 13;

export type PngColourType = 0 | 2 | 3 | 4 | 6;

export type PngImage = {
  readonly widthPixels: number;
  readonly heightPixels: number;
  readonly bitDepth: number;
  readonly colourType: PngColourType;
  readonly interlaced: boolean;
  // Three bytes to an entry, and only an indexed png has one.
  readonly palette: Uint8Array | null;
  // One alpha byte per palette entry, for as many entries as it states.
  readonly transparency: Uint8Array | null;
  // Every IDAT chunk joined, still deflated. A png splits its pixels across as
  // many chunks as it likes and the join is the whole of the stream: taking the
  // first alone reads most files and truncates the rest.
  readonly deflated: Uint8Array;
};

// How many samples a pixel of each kind carries, which is what both the pdf's
// predictor and the split below are measured in.
const SAMPLES: Readonly<Record<PngColourType, number>> = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };

export const samplesOf = (colourType: PngColourType): number => SAMPLES[colourType];

export const hasAlpha = (colourType: PngColourType): boolean =>
  colourType === 4 || colourType === 6;

const isColourType = (value: number): value is PngColourType =>
  value === 0 || value === 2 || value === 3 || value === 4 || value === 6;

/**
 * Reads a png's header and gathers its pixel data, or answers nothing where the
 * bytes are not a png at all.
 *
 * A png this cannot make a pdf image of is still read: what to do about an
 * interlaced one is the writer's business rather than the reader's, and the reader
 * saying so is how the writer can be told apart from a file that is simply broken.
 */
export function readPng(bytes: Uint8Array): PngImage | null {
  if (bytes.byteLength < 8 + 12 + IHDR_LENGTH) return null;
  if (SIGNATURE.some((byte, at) => bytes[at] !== byte)) return null;

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let header: PngImage | null = null;
  let palette: Uint8Array | null = null;
  let transparency: Uint8Array | null = null;
  const parts: Uint8Array[] = [];

  // Every chunk is its length, its name, that many bytes and a checksum. The
  // checksum is not read: a chunk that is corrupt would have to be inflated to
  // find out, and inflating is the very thing this avoids.
  let at = 8;
  while (at + 12 <= bytes.byteLength) {
    const length = view.getUint32(at);
    const name = String.fromCharCode(...bytes.subarray(at + 4, at + 8));
    const from = at + 8;
    if (from + length > bytes.byteLength) return null;

    if (name === "IHDR") {
      if (length < IHDR_LENGTH) return null;
      const colourType = bytes[from + 9] ?? 0;
      if (!isColourType(colourType)) return null;
      header = {
        widthPixels: view.getUint32(from),
        heightPixels: view.getUint32(from + 4),
        bitDepth: bytes[from + 8] ?? 0,
        colourType,
        interlaced: (bytes[from + 12] ?? 0) !== 0,
        palette: null,
        transparency: null,
        deflated: new Uint8Array(0),
      };
    } else if (name === "PLTE") {
      palette = bytes.subarray(from, from + length);
    } else if (name === "tRNS") {
      transparency = bytes.subarray(from, from + length);
    } else if (name === "IDAT") {
      parts.push(bytes.subarray(from, from + length));
    } else if (name === "IEND") {
      break;
    }

    at = from + length + 4;
  }

  if (header === null || parts.length === 0) return null;
  return { ...header, palette, transparency, deflated: joined(parts) };
}

function joined(parts: readonly Uint8Array[]): Uint8Array {
  const total = parts.reduce((sum, part) => sum + part.byteLength, 0);
  const out = new Uint8Array(total);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.byteLength;
  }
  return out;
}

// The colour and the alpha of a png that carries both, each as its own run of
// plain pixels, which is the shape a pdf wants: the picture in one image and what
// shows through it in another.
export type SplitPng = {
  readonly colour: Uint8Array;
  readonly alpha: Uint8Array;
};

/**
 * Pulls the alpha out of a png that interleaves it with the colour.
 *
 * The only path here that opens the pixels at all, and it opens them because
 * nothing else can: a pdf keeps what shows through a picture in a separate image
 * of its own, and the two are one sample apart in every pixel of a png.
 *
 * Answers nothing where the pixels do not add up to the size the header states,
 * which is the one thing worth checking about bytes this is about to walk.
 */
export function splitAlpha(png: PngImage): SplitPng | null {
  const samples = samplesOf(png.colourType);
  const colourSamples = samples - 1;
  const rowBytes = png.widthPixels * samples;

  // A png's pixels are a zlib stream and not a bare deflate one, which is also
  // why they can be handed to a pdf untouched where there is no alpha to pull out:
  // a pdf's own FlateDecode reads exactly the same wrapper.
  let raw;
  try {
    raw = unzlibSync(png.deflated);
  } catch {
    return null;
  }
  if (raw.byteLength < png.heightPixels * (rowBytes + 1)) return null;

  unfilter(raw, png.heightPixels, rowBytes, samples);

  const pixels = png.widthPixels * png.heightPixels;
  const colour = new Uint8Array(pixels * colourSamples);
  const alpha = new Uint8Array(pixels);

  for (let row = 0; row < png.heightPixels; row += 1) {
    // Every row still carries the filter byte it was read with, which is now
    // spent: the pixels behind it have been put back.
    const from = row * (rowBytes + 1) + 1;
    for (let column = 0; column < png.widthPixels; column += 1) {
      const pixel = row * png.widthPixels + column;
      const source = from + column * samples;
      for (let sample = 0; sample < colourSamples; sample += 1) {
        colour[pixel * colourSamples + sample] = raw[source + sample] ?? 0;
      }
      alpha[pixel] = raw[source + colourSamples] ?? 0;
    }
  }

  return { colour, alpha };
}

// Puts back the pixels each row was written as a difference from. Every row states
// which of the five it used, and each looks left to the pixel before and up to the
// row above, which is why this runs forwards and in place: the row above has
// already been put back by the time the row below asks for it.
function unfilter(raw: Uint8Array, rows: number, rowBytes: number, bytesPerPixel: number): void {
  for (let row = 0; row < rows; row += 1) {
    const at = row * (rowBytes + 1);
    const filter = raw[at] ?? 0;
    const from = at + 1;
    const above = from - (rowBytes + 1);

    for (let index = 0; index < rowBytes; index += 1) {
      const left = index >= bytesPerPixel ? (raw[from + index - bytesPerPixel] ?? 0) : 0;
      const up = row > 0 ? (raw[above + index] ?? 0) : 0;
      const upLeft =
        row > 0 && index >= bytesPerPixel ? (raw[above + index - bytesPerPixel] ?? 0) : 0;
      const value = raw[from + index] ?? 0;

      raw[from + index] = (value + restored(filter, left, up, upLeft)) & 0xff;
    }
  }
}

const NONE = 0;
const SUB = 1;
const UP = 2;
const AVERAGE = 3;
const PAETH = 4;

function restored(filter: number, left: number, up: number, upLeft: number): number {
  switch (filter) {
    case NONE:
      return 0;
    case SUB:
      return left;
    case UP:
      return up;
    case AVERAGE:
      return (left + up) >> 1;
    case PAETH:
      return paeth(left, up, upLeft);
    default:
      // A filter number the format does not have. Read as no filter at all, which
      // is the one reading that leaves the row as it was written.
      return 0;
  }
}

// Whichever of the three neighbours the pixel is nearest to, guessed from their
// own straight-line estimate.
function paeth(left: number, up: number, upLeft: number): number {
  const estimate = left + up - upLeft;
  const toLeft = Math.abs(estimate - left);
  const toUp = Math.abs(estimate - up);
  const toUpLeft = Math.abs(estimate - upLeft);
  if (toLeft <= toUp && toLeft <= toUpLeft) return left;
  return toUp <= toUpLeft ? up : upLeft;
}
