import { describe, expect, it } from "vitest";

import { NO_BORDERS } from "../docx/borders.js";
import { WHOLE_FRAME } from "../docx/anchors.js";
import { NO_PAINT } from "../docx/drawing.js";
import { UNPAINTED } from "./floats.js";
import type { SectionGeometry } from "../docx/section.js";
import type { ParagraphMark } from "../docx/styles.js";
import type { LaidOutDocument } from "./document.js";
import {
  drawablesOf,
  onTheDeviceGrid,
  type Drawable,
  type PageDrawing,
  type PlacedGlyphs,
} from "./drawables.js";
import type { ParagraphBox, ParagraphPaint, PlacedCell, PlacedLine } from "./stack.js";

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

    // Everything but the baseline, which lands on the device grid like any other
    // drawn line: 200pt is 833 units and a third, so it is drawn at 199.92.
    expect(drawables).toStrictEqual([
      {
        kind: "glyphs",
        key: "glyphs-0",
        ...STRETCHED,
        glyphs: [{ ...(STRETCHED.glyphs[0] ?? {}), baselinePt: 199.92 }],
      },
    ]);
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

// **What a page draws lands on the device Word draws in**, a three-hundredth of an
// inch down the page. The measurement is in `drawables.ts` beside the rule; what
// is held here is that the rule reaches everything drawn together and nothing
// twice.
const MARK: ParagraphMark = {
  font: { kind: "named", name: "Meridian Sans" },
  fontSizePt: 12,
  bold: false,
  italic: false,
  underline: false,
  raisePt: 0,
  lineSizePt: 12,
  lineRaisePt: 0,
  color: null,
  characterSpacingPt: 0,
  characterScale: 1,
  kernFromHalfPoints: null,
  highlight: null,
  capitals: "none",
};

const lineAt = (
  topPt: number,
  baselinePt: number,
  highlight: string | null = null,
): PlacedLine => ({
  line: {
    segments: [
      {
        kind: "text",
        mark: highlight === null ? MARK : { ...MARK, highlight },
        text: "gralm",
        widthPt: 30,
        offsetPt: 0,
      },
    ],
    widthPt: 30,
    heightPt: 14.6484375,
    ascentPt: 11.7,
    seatPt: 0,
    fontHeightPt: 14.6484375,
    heldOpenPt: null,
  },
  leftPt: 72.137,
  topPt,
  heightPt: 14.6484375,
  seatPt: 0,
  fittingHeightPt: 14.6484375,
  baselinePt,
  startsPage: false,
});

const boxOf = (
  lines: readonly PlacedLine[],
  paint: ParagraphPaint | null = null,
): ParagraphBox => ({
  index: 0,
  topPt: lines[0]?.topPt ?? 0,
  anchorTopPt: lines[0]?.topPt ?? 0,
  heightPt: 14.6484375 * lines.length,
  lines,
  marker: null,
  contentWidthPt: 30,
  markTopPt: lines[0]?.topPt ?? 0,
  contentBottomPt: (lines.at(-1)?.topPt ?? 0) + 14.6484375,
  resumesUnderPt: 0,
  widowControl: false,
  keepNext: false,
  startsPage: false,
  endsPage: false,
  endsPageAtASection: false,
  clipTo: null,
  paint,
});

const bodyPage = (
  boxes: readonly ParagraphBox[],
  cells: readonly PlacedCell[] = [],
): PageDrawing => ({
  ...pageWith(),
  body: boxes,
  cells,
});

const drawablesOfPage = (page: PageDrawing): readonly Drawable[] =>
  drawablesOf(layoutOf(page), page);

const textIn = (page: PageDrawing): Extract<Drawable, { kind: "text" }> => {
  const found = drawablesOfPage(page).find((each) => each.kind === "text");
  if (found === undefined) throw new Error("no text was drawn");
  return found;
};

const paintIn = (page: PageDrawing): Extract<Drawable, { kind: "paint" }> => {
  const found = drawablesOfPage(page).find((each) => each.kind === "paint");
  if (found === undefined) throw new Error("nothing was painted");
  return found;
};

describe("where a page's ink lands", () => {
  // 47.4258 is what this project's own arithmetic gives a first line of 12pt
  // Calibri; Word drew that line at 47.52, which is 198 units.
  it("puts a drawn baseline on the device grid", () => {
    const page = bodyPage([boxOf([lineAt(36, 47.4258)])]);

    expect(textIn(page).boxes[0]?.lines[0]?.baselinePt).toBe(47.52);
  });

  it("leaves the place across the page exactly as the layout measured it", () => {
    const page = bodyPage([boxOf([lineAt(36, 47.4258)])]);
    const line = textIn(page).boxes[0]?.lines[0];

    expect(line?.leftPt).toBe(72.137);
    expect(line?.line.segments[0]?.widthPt).toBe(30);
  });

  // **The height itself is not rounded**, which is what Word's own runs of solid
  // lines show: two lines an exact 14.6484375 apart come out 61 units apart and
  // then 62, about the exact line rather than drifting off it.
  it("keeps an exact height, so a run of lines tracks it rather than drifting", () => {
    // Twelve lines of 12pt Calibri, whose exact height is 14.6484375 and whose
    // first baseline falls a fifth of a step short of one, which is where Word's
    // own runs mix their two gaps.
    const EXACT = 14.6484375;
    const baselines = Array.from({ length: 12 }, (_, at) => 47.628 + at * EXACT);
    const page = bodyPage([boxOf(baselines.map((baseline) => lineAt(baseline - 11.7, baseline)))]);
    const drawn = textIn(page).boxes[0]?.lines.map((line) => line.baselinePt) ?? [];

    const gaps = drawn.slice(1).map((each, at) => Number((each - (drawn[at] ?? 0)).toFixed(4)));
    // Two gaps one step apart, which a height in whole steps could not draw, and a
    // span that is still the exact accumulation to within half a step.
    expect(new Set(gaps)).toStrictEqual(new Set([14.64, 14.88]));
    // However long the run, no line has drifted more than half a step from where
    // the exact height put it: the height accumulates and only the drawing rounds.
    baselines.forEach((exact, at) => {
      expect(Math.abs((drawn[at] ?? 0) - exact)).toBeLessThanOrEqual(0.12 + 1e-9);
    });
  });

  // The first thing that could go wrong: a highlight is painted from the line's
  // own box, so it has to be the same box the text was drawn in.
  it("paints a highlight over exactly the line it belongs to", () => {
    const page = bodyPage([boxOf([lineAt(36.1, 47.4258, "ffff00")])]);
    const line = textIn(page).boxes[0]?.lines[0];
    const highlight = paintIn(page).highlights[0];

    expect(highlight?.topPt).toBe(line?.topPt);
    expect((highlight?.topPt ?? 0) + (highlight?.heightPt ?? 0)).toBe(
      (line?.topPt ?? 0) + (line?.fittingHeightPt ?? 0),
    );
    expect(onTheDeviceGrid(highlight?.topPt ?? 0)).toBe(highlight?.topPt);
    expect(onTheDeviceGrid((highlight?.topPt ?? 0) + (highlight?.heightPt ?? 0))).toBe(
      (highlight?.topPt ?? 0) + (highlight?.heightPt ?? 0),
    );
  });

  it("paints a paragraph's own fill over exactly the lines it holds", () => {
    const paint: ParagraphPaint = {
      leftPt: 36,
      rightPt: 576,
      fillColor: "eeeeee",
      borders: NO_BORDERS,
    };
    const page = bodyPage([boxOf([lineAt(36.1, 47.4258)], paint)]);
    const painted = paintIn(page).paragraphs[0];
    const box = textIn(page).boxes[0];

    expect(painted?.topPt).toBe(box?.lines[0]?.topPt);
    expect(painted?.bottomPt).toBe(box?.contentBottomPt);
    expect(onTheDeviceGrid(painted?.bottomPt ?? 0)).toBe(painted?.bottomPt);
  });

  // A cell's floor is the ceiling of the cell under it, so both edges go on the
  // grid and the height is what lies between: rounding the height on its own would
  // open a gap between two rows that the layout had touching.
  it("keeps two rows touching after both land on the grid", () => {
    const cells: readonly PlacedCell[] = [
      {
        leftPt: 36,
        topPt: 100.1,
        widthPt: 200,
        heightPt: 20.3,
        fillColor: "eeeeee",
        borders: NO_BORDERS,
      },
      {
        leftPt: 36,
        topPt: 120.4,
        widthPt: 200,
        heightPt: 20.3,
        fillColor: "dddddd",
        borders: NO_BORDERS,
      },
    ];
    const painted = paintIn(bodyPage([], cells)).cells;

    const floor = (painted[0]?.topPt ?? 0) + (painted[0]?.heightPt ?? 0);
    expect(floor).toBe(painted[1]?.topPt);
    expect(onTheDeviceGrid(floor)).toBe(floor);
  });

  // The second thing that could go wrong: the grid is applied once. A page already
  // on it comes back untouched, so nothing rounds twice however many times a
  // backend asks for the list.
  it("moves nothing that is already where it will be drawn", () => {
    const page = bodyPage(
      [boxOf([lineAt(36.1, 47.4258, "ffff00")])],
      [
        {
          leftPt: 36,
          topPt: 100.1,
          widthPt: 200,
          heightPt: 20.3,
          fillColor: "eeeeee",
          borders: NO_BORDERS,
        },
      ],
    );
    const once = drawablesOfPage(page);
    const twice = drawablesOfPage({
      ...page,
      body: textIn(page).boxes,
      cells: paintIn(page).cells,
    });

    expect(twice).toStrictEqual(once);
  });
});

// A shape's own text is drawn square with the page only where the shape is: a
// turned one is drawn under a turn of its own, so the page's grid is not the grid
// its lines land on and the layout's own numbers are what a backend turns.
describe("text inside a shape", () => {
  const textBox = (turnDegrees: number): PageDrawing => ({
    ...pageWith(),
    floats: [
      {
        anchor: {
          paragraphIndex: 0,
          name: "Box",
          widthEmu: 0,
          heightEmu: 0,
          turnDegrees,
          flip: { horizontal: false, vertical: false },
          content: { kind: "shape", paint: NO_PAINT },
          horizontal: { kind: "offset", from: "column", offsetEmu: 0 },
          vertical: { kind: "offset", from: "paragraph", offsetEmu: 0 },
          wrap: "none",
          side: "bothSides",
          area: WHOLE_FRAME,
          distances: { topEmu: 0, rightEmu: 0, bottomEmu: 0, leftEmu: 0 },
          behindDoc: false,
          relativeHeight: 1,
        },
        content: {
          kind: "text-box",
          body: {
            blocks: [],
            insets: { leftEmu: 0, topEmu: 0, rightEmu: 0, bottomEmu: 0 },
            anchor: "top",
            wraps: true,
            fitsText: false,
          },
          text: {
            boxes: [boxOf([lineAt(36.1, 47.4258)])],
            cells: [],
            inlines: [],
            contentHeightPt: 20,
            contentWidthPt: 100,
          },
          paint: UNPAINTED,
        },
        leftPt: 100,
        topPt: 40,
        widthPt: 200,
        heightPt: 90,
        turnDegrees,
        flip: { horizontal: false, vertical: false },
      },
    ],
  });

  const baselineIn = (page: PageDrawing): number | undefined => {
    const found = drawablesOfPage(page).find((each) => each.kind === "text");
    return found?.kind === "text" ? found.boxes[0]?.lines[0]?.baselinePt : undefined;
  };

  it("lands on the grid where the shape stands square with the page", () => {
    expect(baselineIn(textBox(0))).toBe(47.52);
  });

  it("is left exactly where layout put it where the shape is turned", () => {
    expect(baselineIn(textBox(30))).toBe(47.4258);
  });
});
