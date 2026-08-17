import { describe, expect, it } from "vitest";

import type { Looks, PageLooks } from "./compare.js";
import { floorOf, reportsAGap } from "./floor.js";

// What a ranking subtracts, and the two decisions inside it: which documents answer
// for the floor at all, and whether a page or a document is the unit.
//
// Neither is arithmetic anyone can check by eye against a real run, since a run wants
// Chrome, poppler and a manifest none of which belong in a test. So the arithmetic is
// held here against drawings stated by hand.

const looksLike = (
  id: string,
  pages: readonly PageLooks[],
  asks: readonly string[] = [],
): Looks => ({
  id,
  outcome: "compared",
  pagesOurs: pages.length,
  pagesWord: pages.length,
  facesStoodIn: 0,
  asks,
  interesting: pages.reduce((sum, each) => sum + each.interesting, 0),
  differing: pages.reduce((sum, each) => sum + each.differing, 0),
  pages,
  detail: "",
});

// A document of five pages, one of them wrong and the rest drawn exactly: the shape
// every one of the eight has, and the reason a document's own share says so little.
const ONE_BAD_PAGE = looksLike("h", [
  { interesting: 1000, differing: 100 },
  { interesting: 1000, differing: 0 },
  { interesting: 1000, differing: 0 },
  { interesting: 1000, differing: 0 },
  { interesting: 1000, differing: 0 },
]);

const EXACT = looksLike("a", [
  { interesting: 1000, differing: 0 },
  { interesting: 1000, differing: 0 },
]);

describe("which documents answer for the floor", () => {
  // A document reporting a gap holds something nobody built, so its cells are a
  // feature counted as though it were rasteriser noise. `readUnhonoured` is what
  // answers, so no list of names is kept anywhere.
  it("sets aside a document that says it holds something nobody built", () => {
    expect(reportsAGap(looksLike("h", [], ["unknown-drawing"]))).toBe(true);
    expect(reportsAGap(looksLike("a", []))).toBe(false);
  });

  // The one of the eight holds a chart and a page background; either is enough.
  it("sets one aside whatever it is short of", () => {
    expect(reportsAGap(looksLike("h", [], ["page-background", "unknown-drawing"]))).toBe(true);
  });

  /**
   * **A report is not one population.** Three of its kinds are what the layout found
   * when it went looking for a face rather than what the file asked for, and a
   * document reporting one of those is not holding anything nobody built.
   *
   * This is not a distinction on paper: on 2026-08-14 four of the eight reported a
   * borrowed character and one of the four was drawn cell for cell as Word drew it on
   * every page. Setting those aside would have left three documents answering for the
   * floor and a floor of nought, which is a floor that has measured nothing.
   */
  it("keeps a document whose report is about this machine's fonts", () => {
    expect(reportsAGap(looksLike("e", [], ["character-from-another-face"]))).toBe(false);
    expect(reportsAGap(looksLike("e", [], ["substituted-face"]))).toBe(false);
    expect(reportsAGap(looksLike("e", [], ["missing-glyph"]))).toBe(false);
    // And a document short of both is still set aside for the half that was built.
    expect(reportsAGap(looksLike("h", [], ["character-from-another-face", "equation"]))).toBe(true);
  });
});

describe("the floor as a page reads it", () => {
  // **A document's own share dilutes a long document.** One page wrong in five reads
  // as 2% of everything, where the page itself is 10% wrong: a ranking subtracting
  // the first forgives twice as much as it means to on a one-pager.
  it("answers with the worst page rather than the share of every cell", () => {
    const reading = floorOf([ONE_BAD_PAGE]);

    expect(reading.worstPage).toBeCloseTo(0.1, 10);
    expect(reading.ofEveryCell).toBeCloseTo(0.02, 10);
  });

  // **Exactly, and not nearly.** Through one rasteriser a page holding the same
  // things in the same places comes out cell for cell the same, so what is counted
  // is what is provably right rather than what is close enough to argue about.
  it("counts the pages drawn cell for cell as Word drew them", () => {
    const reading = floorOf([ONE_BAD_PAGE, EXACT]);

    expect(reading.pages).toBe(7);
    expect(reading.pagesExactlyEqual).toBe(6);
  });

  it("takes the worst page across the documents and not the worst document", () => {
    const shorter = looksLike("g", [{ interesting: 100, differing: 30 }]);

    // The shorter document is a fifth of the cells and three times the worst page.
    expect(floorOf([ONE_BAD_PAGE, shorter]).worstPage).toBeCloseTo(0.3, 10);
    expect(floorOf([ONE_BAD_PAGE, shorter]).ofEveryCell).toBeLessThan(0.03);
  });

  // A page holding almost nothing can be wholly wrong about the little it holds, so
  // `worstPageOf` needs a page to have drawn something before it can lead. What it
  // falls back on is the document's own share, which is what this states.
  it("does not let a page drawing almost nothing lead", () => {
    const nearlyEmpty = looksLike("b", [
      { interesting: 10, differing: 10 },
      { interesting: 4000, differing: 0 },
    ]);

    expect(floorOf([nearlyEmpty]).worstPage).toBeLessThan(0.01);
  });

  it("answers nothing for nothing measured", () => {
    expect(floorOf([])).toStrictEqual({
      documents: 0,
      pages: 0,
      pagesExactlyEqual: 0,
      cells: 0,
      differing: 0,
      ofEveryCell: 0,
      worstPage: 0,
    });
  });
});
