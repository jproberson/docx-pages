import { describe, expect, it } from "vitest";

import { buildPng } from "../testing/build-png.js";
import { overlayOf } from "./overlay.js";
import { readPng } from "./png.js";

const dot = (grey: number): ReturnType<typeof readPng> => readPng(buildPng(1, 1, "grey", [grey]));

const colourAt = (image: ReturnType<typeof overlayOf>): readonly number[] => [
  ...image.pixels.subarray(0, 3),
];

describe("two drawings of a page in one image", () => {
  it("leaves ink both put down black", () => {
    expect(colourAt(overlayOf(dot(0), dot(0)))).toEqual([0, 0, 0]);
  });

  it("leaves paper both left alone white", () => {
    expect(colourAt(overlayOf(dot(255), dot(255)))).toEqual([255, 255, 255]);
  });

  it("draws ink only we put down in red", () => {
    expect(colourAt(overlayOf(dot(0), dot(255)))).toEqual([255, 0, 0]);
  });

  it("draws ink only Word put down in green", () => {
    expect(colourAt(overlayOf(dot(255), dot(0)))).toEqual([0, 255, 0]);
  });

  // A page one side never drew is a page of the other side's ink alone, which is
  // what a document making the wrong number of pages comes out as.
  it("draws a page one side never drew as the other side's ink", () => {
    expect(colourAt(overlayOf(null, dot(0)))).toEqual([0, 255, 0]);
  });
});
