import { describe, expect, it } from "vitest";

import { buildPng } from "../testing/build-png.js";
import { readPng, writePng } from "./png.js";

const pixelAt = (image: ReturnType<typeof readPng>, x: number, y: number): readonly number[] => {
  const at = (y * image.width + x) * 4;
  return [...image.pixels.subarray(at, at + 4)];
};

describe("reading a drawing back", () => {
  it("reads grey", () => {
    const image = readPng(buildPng(2, 2, "grey", [0, 128, 200, 255]));
    expect([image.width, image.height]).toEqual([2, 2]);
    expect(pixelAt(image, 0, 0)).toEqual([0, 0, 0, 255]);
    expect(pixelAt(image, 1, 1)).toEqual([255, 255, 255, 255]);
  });

  it("reads colour", () => {
    const image = readPng(buildPng(2, 1, "red green blue", [255, 0, 0, 0, 0, 255]));
    expect(pixelAt(image, 0, 0)).toEqual([255, 0, 0, 255]);
    expect(pixelAt(image, 1, 0)).toEqual([0, 0, 255, 255]);
  });

  it("reads what a pixel is drawn through", () => {
    const image = readPng(buildPng(1, 1, "red green blue alpha", [10, 20, 30, 40]));
    expect(pixelAt(image, 0, 0)).toEqual([10, 20, 30, 40]);
  });

  // Every row states a difference from what is already undone rather than its own
  // samples, so a reader that walks the rows in the wrong order or forgets the
  // pixel to the left comes out with different bytes rather than with none.
  it("undoes a row stated as a difference from the one above it", () => {
    const samples = [10, 20, 30, 40, 50, 60];
    expect(readPng(buildPng(3, 2, "grey", samples, [0, 2])).pixels).toEqual(
      readPng(buildPng(3, 2, "grey", samples)).pixels,
    );
  });

  it("undoes a row stated as a difference from the pixel to its left", () => {
    const samples = [10, 20, 30, 40, 50, 60];
    expect(readPng(buildPng(3, 2, "grey", samples, [1, 1])).pixels).toEqual(
      readPng(buildPng(3, 2, "grey", samples)).pixels,
    );
  });

  it("refuses what it cannot read rather than guessing", () => {
    expect(() => readPng(Uint8Array.from([1, 2, 3]))).toThrow(/not a png/);
  });

  it("writes a drawing that reads back as itself", () => {
    const drawn = readPng(buildPng(2, 2, "red green blue", [1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 1, 2]));
    expect(readPng(writePng(drawn))).toEqual(drawn);
  });
});
