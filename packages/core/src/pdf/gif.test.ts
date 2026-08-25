import { describe, expect, it } from "vitest";

import { readGif } from "./gif.js";

// **The fixtures are ImageMagick's, not this project's.** A gif hand-encoded here
// would be encoded to the rules this decoder happens to hold, which tests nothing:
// it would agree with its own bugs. Each of these was written by `magick` and its
// pixels read back with `magick ... txt:-`, so the expectations below are another
// implementation's answer about the same bytes.

// 4 by 2, three colours, no transparency.
const PLAIN = Uint8Array.from([
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x04, 0x00, 0x02, 0x00, 0xf1, 0x00, 0x00, 0xff, 0x00, 0x00,
  0x00, 0xff, 0x00, 0x00, 0x00, 0xff, 0x00, 0x00, 0x00, 0x21, 0xf9, 0x04, 0x00, 0x00, 0x00, 0x00,
  0x00, 0x2c, 0x00, 0x00, 0x00, 0x00, 0x04, 0x00, 0x02, 0x00, 0x00, 0x02, 0x04, 0x44, 0x80, 0x28,
  0x05, 0x00, 0x3b,
]);

// 4 by 4 and interlaced, which stores its rows in four passes.
const INTERLACED = Uint8Array.from([
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x04, 0x00, 0x04, 0x00, 0xf0, 0x00, 0x00, 0x11, 0x22, 0x33,
  0xdd, 0xee, 0xff, 0x21, 0xf9, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x2c, 0x00, 0x00, 0x00, 0x00,
  0x04, 0x00, 0x04, 0x00, 0x40, 0x02, 0x05, 0x8c, 0x03, 0xa9, 0x79, 0x51, 0x00, 0x3b,
]);

// 3 by 1, one drawn pixel and two the page shows through.
const TRANSPARENT = Uint8Array.from([
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x03, 0x00, 0x01, 0x00, 0xf0, 0x00, 0x00, 0x00, 0x00, 0x00,
  0xff, 0x88, 0x00, 0x21, 0xf9, 0x04, 0x01, 0x00, 0x00, 0x00, 0x00, 0x2c, 0x00, 0x00, 0x00, 0x00,
  0x03, 0x00, 0x01, 0x00, 0x00, 0x02, 0x02, 0x0c, 0x50, 0x00, 0x3b,
]);

const pixelsOf = (colour: Uint8Array): readonly (readonly number[])[] => {
  const out: number[][] = [];
  for (let at = 0; at < colour.length; at += 3) {
    out.push([colour[at] ?? 0, colour[at + 1] ?? 0, colour[at + 2] ?? 0]);
  }
  return out;
};

describe("readGif", () => {
  it("reads the pixels of a plain gif, in the order a pdf wants them", () => {
    const gif = readGif(PLAIN);

    expect(gif).not.toBeNull();
    expect([gif?.widthPixels, gif?.heightPixels]).toStrictEqual([4, 2]);
    expect(pixelsOf(gif?.colour ?? new Uint8Array())).toStrictEqual([
      [255, 0, 0],
      [0, 255, 0],
      [255, 0, 0],
      [255, 0, 0],
      [255, 0, 0],
      [255, 0, 0],
      [255, 0, 0],
      [0, 0, 255],
    ]);
  });

  it("draws no mask for a gif that marks nothing transparent", () => {
    expect(readGif(PLAIN)?.alpha).toBeNull();
  });

  // The rows of an interlaced gif are stored in four passes rather than in order,
  // so reading them straight draws the picture in stripes.
  it("weaves an interlaced gif's rows back into their own order", () => {
    const gif = readGif(INTERLACED);
    const light = [221, 238, 255];
    const dark = [17, 34, 51];

    expect([gif?.widthPixels, gif?.heightPixels]).toStrictEqual([4, 4]);
    expect(pixelsOf(gif?.colour ?? new Uint8Array())).toStrictEqual([
      light,
      light,
      light,
      light,
      dark,
      dark,
      dark,
      dark,
      dark,
      dark,
      dark,
      dark,
      light,
      light,
      light,
      light,
    ]);
  });

  it("gives a transparent index nought alpha and everything else full", () => {
    const gif = readGif(TRANSPARENT);

    expect([gif?.widthPixels, gif?.heightPixels]).toStrictEqual([3, 1]);
    expect([...(gif?.alpha ?? [])]).toStrictEqual([255, 0, 0]);
    expect(pixelsOf(gif?.colour ?? new Uint8Array())[0]).toStrictEqual([255, 136, 0]);
  });

  it("answers nothing for bytes that are not a gif at all", () => {
    expect(readGif(Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 0]))).toBeNull();
    expect(readGif(new Uint8Array(0))).toBeNull();
  });

  it("answers nothing for a gif cut off part way rather than drawing half of one", () => {
    expect(readGif(PLAIN.subarray(0, 30))).toBeNull();
  });
});
