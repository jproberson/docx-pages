import { pngOf } from "../raster/png.js";

// A png written by hand, so the reader that reads a drawing back is tested
// against bytes this file decides rather than against whatever Chrome happened to
// write. The same reason `build-pdf.ts` writes a pdf.

export type PngColour = "grey" | "red green blue" | "red green blue alpha";

const CHANNELS: Record<PngColour, number> = {
  grey: 1,
  "red green blue": 3,
  "red green blue alpha": 4,
};

const TYPES: Record<PngColour, number> = {
  grey: 0,
  "red green blue": 2,
  "red green blue alpha": 6,
};

/**
 * A png of the given samples, one row after another, under the row filter each
 * row names. `filters` is how the reader's undoing of them is exercised: a row
 * filtered `2` states its difference from the row above it, so this writes the
 * difference rather than the sample and the reader has to put it back.
 */
export function buildPng(
  width: number,
  height: number,
  colour: PngColour,
  samples: readonly number[],
  filters: readonly number[] = [],
): Uint8Array {
  const channels = CHANNELS[colour];
  const stride = width * channels;
  const rows = new Uint8Array((stride + 1) * height);

  for (let row = 0; row < height; row += 1) {
    const filter = filters[row] ?? 0;
    rows[row * (stride + 1)] = filter;
    for (let at = 0; at < stride; at += 1) {
      const sample = samples[row * stride + at] ?? 0;
      const above = row > 0 ? (samples[(row - 1) * stride + at] ?? 0) : 0;
      const left = at >= channels ? (samples[row * stride + at - channels] ?? 0) : 0;
      const stated =
        filter === 1 ? sample - left : filter === 2 ? sample - above : filter === 0 ? sample : null;
      if (stated === null) throw new Error(`this writer states no filter ${String(filter)}`);
      rows[row * (stride + 1) + 1 + at] = stated & 0xff;
    }
  }

  return pngOf(width, height, TYPES[colour], rows);
}
