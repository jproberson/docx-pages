import { describe, expect, it } from "vitest";

import type { SectionGeometry } from "../docx/section.js";
import type { LaidOutDocument } from "./document.js";
import { drawablesOf, type PageDrawing, type PlacedGlyphs } from "./drawables.js";

// A glyph named by number is the one thing a page cannot ask for in characters,
// and where it is drawn is decided here rather than in either backend.

const LETTER: SectionGeometry = {
  widthTwips: 12240,
  heightTwips: 15840,
  margin: {
    topTwips: 720,
    rightTwips: 720,
    bottomTwips: 720,
    leftTwips: 720,
    headerTwips: 432,
    footerTwips: 144,
  },
};

const STRETCHED: PlacedGlyphs = {
  face: { name: "Meridian Math", bold: false, italic: false },
  sizePt: 11,
  color: "000000",
  ascentPt: 14,
  descentPt: 7.6,
  glyphs: [{ glyph: 3436, leftPt: 100, baselinePt: 200, advancePt: 5.44, standsFor: "(" }],
};

function pageWith(glyphRuns?: readonly PlacedGlyphs[]): PageDrawing {
  return {
    index: 0,
    geometry: LETTER,
    body: [],
    cells: [],
    floats: [],
    inlines: [],
    headerTopPt: 21.6,
    headerHeightPt: 0,
    footerTopPt: 784.8,
    bodyTopPt: 36,
    bodyBottomPt: 792,
    header: [],
    footer: [],
    headerCells: [],
    footerCells: [],
    headerFloats: [],
    footerFloats: [],
    headerInlines: [],
    footerInlines: [],
    // A page laid out before anything named a glyph states no runs at all, which
    // is not the same as stating none.
    ...(glyphRuns === undefined ? {} : { glyphRuns }),
  };
}

const layoutOf = (page: PageDrawing): LaidOutDocument => ({
  kind: "laid-out",
  page: LETTER,
  unhonoured: [],
  headerTopPt: 21.6,
  bodyTopPt: 36,
  bodyBottomPt: 792,
  pages: [page],
});

const drawn = (glyphRuns: readonly PlacedGlyphs[]): readonly string[] => {
  const page = pageWith(glyphRuns);
  return drawablesOf(layoutOf(page), page).map((drawable) => drawable.kind);
};

describe("a run of glyphs named by number", () => {
  it("comes back as a drawable of its own, carrying what the layout stated", () => {
    const page = pageWith([STRETCHED]);
    const drawables = drawablesOf(layoutOf(page), page);

    expect(drawables).toStrictEqual([{ kind: "glyphs", key: "glyphs-0", ...STRETCHED }]);
  });

  // They are text, so what stands over them and under them is what stands over and
  // under any other text.
  it("is drawn where the story's own text is drawn", () => {
    const page = { ...pageWith([STRETCHED]), cells: [] };
    const kinds = drawablesOf(layoutOf(page), page).map((drawable) => drawable.kind);

    expect(kinds).toStrictEqual(["glyphs"]);
  });

  it("draws nothing for a run holding no glyph at all", () => {
    expect(drawn([{ ...STRETCHED, glyphs: [] }])).toStrictEqual([]);
  });

  // A page laid out before anything named a glyph carries none, and draws as it
  // always did.
  it("draws nothing for a page that states no runs", () => {
    const without = pageWith();

    expect(drawablesOf(layoutOf(without), without)).toStrictEqual([]);
    expect(drawn([])).toStrictEqual([]);
  });

  it("keeps each run apart, in the order the page states them", () => {
    const second: PlacedGlyphs = {
      ...STRETCHED,
      sizePt: 22,
      glyphs: [{ glyph: 12, leftPt: 300, baselinePt: 400, advancePt: 9, standsFor: null }],
    };

    expect(
      drawn([STRETCHED, second]).length === 2 &&
        drawablesOf(layoutOf(pageWith([STRETCHED, second])), pageWith([STRETCHED, second])).map(
          (drawable) => drawable.key,
        ),
    ).toStrictEqual(["glyphs-0", "glyphs-1"]);
  });
});
