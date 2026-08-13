import { describe, expect, it } from "vitest";

import { pngFromMetafile, readMetafileBitmap } from "./wmf.js";

// A placeable metafile holding one record, built here rather than taken from a
// document: the corpus is other people's work and none of it may be committed.
//
// Word writes the older metafile beside the one this project plays, and every one
// of the eight in the 718 documents records the same single thing: a bitmap blitted
// into the frame. So what is built here is that, and what is asserted is the
// bitmap coming back out with its rows and its channels the right way round.
const metafile = (records: readonly Uint8Array[], placeable = true): Uint8Array => {
  const header = new Uint8Array(18);
  new DataView(header.buffer).setUint16(0, 1, true);
  // Nine words of header, which is what the format states and what this checks.
  new DataView(header.buffer).setUint16(2, 9, true);

  const parts = placeable ? [placeableHeader(), header, ...records] : [header, ...records];
  const out = new Uint8Array(parts.reduce((sum, each) => sum + each.byteLength, 0));
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.byteLength;
  }
  return out;
};

const placeableHeader = (): Uint8Array => {
  const bytes = new Uint8Array(22);
  new DataView(bytes.buffer).setUint32(0, 0x9ac6cdd7, true);
  return bytes;
};

const record = (fn: number, parameters: Uint8Array): Uint8Array => {
  const bytes = new Uint8Array(6 + parameters.byteLength);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, bytes.byteLength / 2, true);
  view.setUint16(4, fn, true);
  bytes.set(parameters, 6);
  return bytes;
};

// Two pixels by two, stored bottom up as a bitmap is, in the blue-green-red order
// a bitmap keeps and a png does not.
const bitmapOf = (
  rows: readonly (readonly (readonly [number, number, number])[])[],
  compression = 0,
): Uint8Array => {
  const width = rows[0]?.length ?? 0;
  const stride = Math.ceil((width * 24) / 32) * 4;
  const bytes = new Uint8Array(40 + stride * rows.length);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 40, true);
  view.setInt32(4, width, true);
  view.setInt32(8, rows.length, true);
  view.setUint16(12, 1, true);
  view.setUint16(14, 24, true);
  view.setUint32(16, compression, true);

  for (const [at, row] of [...rows].reverse().entries()) {
    let out = 40 + at * stride;
    for (const [red, green, blue] of row) {
      bytes[out] = blue;
      bytes[out + 1] = green;
      bytes[out + 2] = red;
      out += 3;
    }
  }
  return bytes;
};

const RED = [255, 0, 0] as const;
const GREEN = [0, 255, 0] as const;
const BLUE = [0, 0, 255] as const;
const WHITE = [255, 255, 255] as const;

// The parameters a `DibStretchBlt` states before its bitmap, which this reader
// walks past by looking for the bitmap's own header rather than counting them.
const BEFORE_THE_BITMAP = new Uint8Array(20);

describe("the bitmap an older metafile blits", () => {
  it("reads it out with its rows and channels the right way round", () => {
    const dib = bitmapOf([
      [RED, GREEN],
      [BLUE, WHITE],
    ]);
    const parameters = new Uint8Array(BEFORE_THE_BITMAP.byteLength + dib.byteLength);
    parameters.set(dib, BEFORE_THE_BITMAP.byteLength);

    const found = readMetafileBitmap(metafile([record(0x0b41, parameters)]));

    expect(found?.widthPixels).toBe(2);
    expect(found?.heightPixels).toBe(2);
    expect([...(found?.rgb ?? [])]).toStrictEqual([...RED, ...GREEN, ...BLUE, ...WHITE]);
  });

  // The corpus files open with an escape carrying the printer's own comment, and
  // the bitmap stands well inside them, so the records are walked rather than the
  // first one taken.
  it("finds the bitmap behind whatever records stand in front of it", () => {
    const dib = bitmapOf([[RED]]);
    const parameters = new Uint8Array(BEFORE_THE_BITMAP.byteLength + dib.byteLength);
    parameters.set(dib, BEFORE_THE_BITMAP.byteLength);

    const found = readMetafileBitmap(
      metafile([
        record(0x0626, new Uint8Array(64)),
        record(0x0103, new Uint8Array(2)),
        record(0x0b41, parameters),
      ]),
    );

    expect(found?.widthPixels).toBe(1);
  });

  it("reads a metafile that carries no placeable header of its own", () => {
    const dib = bitmapOf([[GREEN]]);
    const parameters = new Uint8Array(BEFORE_THE_BITMAP.byteLength + dib.byteLength);
    parameters.set(dib, BEFORE_THE_BITMAP.byteLength);

    expect(readMetafileBitmap(metafile([record(0x0940, parameters)], false))?.rgb[1]).toBe(255);
  });

  // **What this cannot read is refused rather than approximated**, and the report
  // goes on naming the picture undrawable, which is what it is for.
  it("refuses a bitmap that is compressed", () => {
    const dib = bitmapOf([[RED]], 1);
    const parameters = new Uint8Array(BEFORE_THE_BITMAP.byteLength + dib.byteLength);
    parameters.set(dib, BEFORE_THE_BITMAP.byteLength);

    expect(readMetafileBitmap(metafile([record(0x0b41, parameters)]))).toBeNull();
  });

  it("refuses a metafile that blits nothing at all", () => {
    expect(readMetafileBitmap(metafile([record(0x0103, new Uint8Array(2))]))).toBeNull();
    expect(readMetafileBitmap(new Uint8Array(4))).toBeNull();
  });

  it("writes the bitmap out as a png both renderers already draw", () => {
    const dib = bitmapOf([
      [RED, GREEN],
      [BLUE, WHITE],
    ]);
    const parameters = new Uint8Array(BEFORE_THE_BITMAP.byteLength + dib.byteLength);
    parameters.set(dib, BEFORE_THE_BITMAP.byteLength);

    const png = pngFromMetafile(metafile([record(0x0b41, parameters)]));

    expect([...(png?.subarray(0, 8) ?? [])]).toStrictEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    // The header chunk states the size and the three channels it was written with.
    expect([...(png?.subarray(12, 16) ?? [])]).toStrictEqual([73, 72, 68, 82]);
    expect(png?.[19]).toBe(2);
    expect(png?.[23]).toBe(2);
    expect(png?.[25]).toBe(2);
  });
});
