import { describe, expect, it } from "vitest";

import { buildPng } from "../testing/build-png.js";
import { CELL_PX, differenceBetween, gridOf, shareOf } from "./difference.js";
import { readPng } from "./png.js";

// A page of one cell, drawn in whatever grey is asked for over a given share of it.
function page(covered: number, grey = 0): ReturnType<typeof readPng> {
  const samples: number[] = [];
  for (let y = 0; y < CELL_PX; y += 1) {
    for (let x = 0; x < CELL_PX; x += 1) {
      samples.push(y * CELL_PX + x < covered * CELL_PX * CELL_PX ? grey : 255);
    }
  }
  return readPng(buildPng(CELL_PX, CELL_PX, "grey", samples));
}

const between = (ours: number, theirs: number, grey = 0): number =>
  shareOf(differenceBetween(gridOf(page(ours, grey)), gridOf(page(theirs))));

describe("how different two drawings of a page look", () => {
  it("counts nothing against a page neither side drew on", () => {
    expect(differenceBetween(gridOf(page(0)), gridOf(page(0)))).toEqual({
      interesting: 0,
      differing: 0,
    });
  });

  // The whole reason nothing here counts pixels: the same glyph hinted two ways
  // covers a slightly different share of its cell, and that has to read as
  // agreement or every document scores the same.
  it("lets a cell the two drew nearly alike alone", () => {
    expect(between(0.5, 0.55)).toBe(0);
  });

  it("counts a cell one side drew in and the other did not", () => {
    expect(between(0.5, 0)).toBe(1);
  });

  it("counts a cell the two drew in different colours", () => {
    expect(between(1, 1, 200)).toBe(1);
  });

  it("counts every cell of a page one side never drew", () => {
    const drawn = gridOf(page(1));
    expect(differenceBetween(null, drawn)).toEqual({ interesting: 1, differing: 1 });
    expect(differenceBetween(drawn, null)).toEqual({ interesting: 1, differing: 1 });
  });

  it("counts what stands off the end of the smaller of two pages", () => {
    const wide = gridOf(
      readPng(buildPng(CELL_PX * 2, CELL_PX, "grey", Array(CELL_PX * 2 * CELL_PX).fill(0))),
    );
    expect(differenceBetween(gridOf(page(1)), wide)).toEqual({ interesting: 2, differing: 1 });
  });
});
