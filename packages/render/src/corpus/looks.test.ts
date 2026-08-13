import { describe, expect, it } from "vitest";

import type { Looks } from "../raster/compare.js";
import { drawnToAScaleIn, linesPlacedIn, reportOf } from "./looks.js";

const row = (id: string, differing: number, interesting = 1000): Looks => ({
  id,
  outcome: "compared",
  pagesOurs: 1,
  pagesWord: 1,
  facesStoodIn: 0,
  asks: [],
  interesting,
  differing,
  pages: [{ interesting, differing }],
  detail: "",
});

const agreement = [
  JSON.stringify({ id: "aaaaaaaaaaaa", placed: 40, lines: 44, outcome: "compared" }),
  JSON.stringify({
    id: "bbbbbbbbbbbb",
    placed: 0,
    lines: 88,
    outcome: "drawn to a scale",
    detail: "drawn at 74.9% of the stated size",
  }),
  "",
].join("\n");

describe("drawnToAScaleIn", () => {
  it("names the documents the line score says Word drew shrunk", () => {
    expect([...drawnToAScaleIn(agreement)]).toStrictEqual(["bbbbbbbbbbbb"]);
  });

  it("says nothing about a sweep that has none", () => {
    expect(drawnToAScaleIn("").size).toBe(0);
  });

  it("reads the lines of the same file without tripping over the outcome", () => {
    expect(linesPlacedIn(agreement).get("aaaaaaaaaaaa")).toStrictEqual({ placed: 40, lines: 44 });
  });
});

describe("reportOf", () => {
  // The whole point of the join: a document Word drew shrunk differs in most of its
  // cells by construction, so left in it leads the queue and is not a fault.
  it("leaves a document drawn to a scale out of the ranking", () => {
    const report = reportOf(
      [row("aaaaaaaaaaaa", 10), row("bbbbbbbbbbbb", 650)],
      linesPlacedIn(agreement),
      drawnToAScaleIn(agreement),
    );
    expect(report).toContain("aaaaaaaaaaaa");
    expect(report).not.toContain("bbbbbbbbbbbb");
  });

  it("counts it rather than dropping it in silence", () => {
    const report = reportOf(
      [row("aaaaaaaaaaaa", 10), row("bbbbbbbbbbbb", 650)],
      linesPlacedIn(agreement),
      drawnToAScaleIn(agreement),
    );
    expect(report).toContain("drawn to a scale 1");
    expect(report).toContain("1 documents, 1000 cells drawn in");
  });

  it("ranks every document where no sweep says any were scaled", () => {
    const report = reportOf([row("aaaaaaaaaaaa", 10), row("bbbbbbbbbbbb", 650)], new Map());
    expect(report).toContain("bbbbbbbbbbbb");
  });
});
