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
  mathDrawables,
  METAFILE_PEN_OFFSET,
  onTheDeviceGrid,
  runWidthMadeUpBy,
  type Drawable,
  type DrawnRun,
  type PageDrawing,
  type PlacedGlyphs,
  type SetEquation,
} from "./drawables.js";
import type { MathPrimitive, SetMath } from "./math.js";
import type { PaintedFill } from "./painting.js";
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

    // Everything but the two the device grid takes: the baseline, as any other drawn
    // line, and the size, as any other drawn run. 200pt is 833 units and a third, so
    // it is drawn at 199.92, and 11pt is 45 and five sixths, so it is set at 11.04.
    expect(drawables).toStrictEqual([
      {
        kind: "glyphs",
        key: "glyphs-0",
        ...STRETCHED,
        sizePt: 11.04,
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

  // The shape a backend that cannot name a glyph draws instead. It is measured
  // from the glyph's own origin, so nothing that moves the glyph touches it: the
  // baseline lands on the grid and the outline is carried through as it stands.
  it("carries the face's own outline through untouched", () => {
    const outline = {
      unitsPerEm: 2048,
      contours: [
        {
          from: [0, 0] as const,
          steps: [
            { kind: "line" as const, to: [100, 0] as const },
            { kind: "quadratic" as const, control: [150, 50] as const, to: [100, 100] as const },
          ],
        },
      ],
    };
    const carved: PlacedGlyphs = {
      ...STRETCHED,
      glyphs: STRETCHED.glyphs.map((glyph) => ({ ...glyph, outline })),
    };
    const page = pageWith([carved]);
    const first = drawablesOf(layoutOf(page), page)[0];

    expect(first?.kind === "glyphs" && first.glyphs[0]?.outline).toBe(outline);
    expect(first?.kind === "glyphs" && first.glyphs[0]?.baselinePt).toBe(199.92);
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

type PaintDrawable = Extract<Drawable, { kind: "paint" }>;

const paintIn = (page: PageDrawing): PaintDrawable => {
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
    // The paragraph's own fill, which reaches a backend as the rectangle to fill.
    const filled = paintIn(page).painted[0]?.fills[0];
    const box = textIn(page).boxes[0];

    expect(filled?.topPt).toBe(box?.lines[0]?.topPt);
    expect((filled?.topPt ?? 0) + (filled?.heightPt ?? 0)).toBe(box?.contentBottomPt);
    expect(onTheDeviceGrid(filled?.heightPt ?? 0)).toBe(filled?.heightPt);
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
    const painted = paintIn(bodyPage([], cells)).painted;
    const first = painted[0]?.fills[0];
    const second = painted[1]?.fills[0];

    const floor = (first?.topPt ?? 0) + (first?.heightPt ?? 0);
    expect(floor).toBe(second?.topPt);
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
    const twice = drawablesOfPage({ ...page, body: textIn(page).boxes });

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

// Four rules a renderer used to decide for itself, moved here so that neither
// backend can answer them differently and a third cannot miss them.
describe("what a page draws rather than what a renderer decides", () => {
  const wingdings = (text: string): ParagraphBox =>
    boxOf([
      {
        ...lineAt(36, 47.4258),
        line: {
          ...lineAt(36, 47.4258).line,
          segments: [
            {
              kind: "text",
              mark: { ...MARK, font: { kind: "named", name: "Wingdings" } },
              text,
              widthPt: 12,
              offsetPt: 0,
            },
          ],
        },
      },
    ]);

  // A run in a symbol face that was stood in for holds positions in that face's
  // own page; the stand-in would draw them as its own letters.
  it("shows a stood-in symbol face's positions as what they mean, when told", () => {
    const page = bodyPage([wingdings("l")]);
    const shown = (aliasSymbolFaces: ReadonlySet<string> | null): string | undefined => {
      const drawn = drawablesOf(layoutOf(page), page, { aliasSymbolFaces });
      const first = drawn.find((each) => each.kind === "text");
      const segment = first?.kind === "text" ? first.boxes[0]?.lines[0]?.line.segments[0] : null;
      return segment?.kind === "text" ? segment.text : undefined;
    };

    expect(shown(new Set(["wingdings"]))).toBe("●");
    expect(shown(null)).toBe("l");
  });

  // A pdf has no such thing as an underline and neither has Word: the line is
  // filled, where the drawn face's own `post` table says to put it.
  it("states the rectangle an underline is drawn as, out of the face's own metrics", () => {
    const underlined = boxOf([
      {
        ...lineAt(36, 47.4258),
        line: {
          ...lineAt(36, 47.4258).line,
          segments: [
            {
              kind: "text",
              mark: { ...MARK, underline: true },
              text: "gralm",
              widthPt: 30,
              offsetPt: 0,
            },
          ],
        },
      },
    ]);
    const page = bodyPage([underlined]);
    const drawn = drawablesOf(layoutOf(page), page, {
      underlineFor: () => ({ belowBaselinePt: 1.45, thicknessPt: 0.83 }),
    });
    const text = drawn.find((each) => each.kind === "text");

    expect(text?.kind === "text" && text.underlines).toStrictEqual([
      { leftPt: 72.137, topPt: 47.52 + 1.45, widthPt: 30, heightPt: 0.83, color: "#000000" },
    ]);
  });

  // A backend holding no faces cannot be told where the line goes, and gets no
  // rectangle rather than one in a made-up place.
  it("states none where nothing can say where the line goes", () => {
    const page = bodyPage([boxOf([lineAt(36, 47.4258)])]);
    const text = drawablesOf(layoutOf(page), page).find((each) => each.kind === "text");

    expect(text?.kind === "text" && text.underlines).toStrictEqual([]);
  });

  // A dashed line runs four widths on and four off, where a dotted one runs one
  // and one, measured at a width of a point and a half.
  it("states how the dashes of a border fall, at the width it is drawn", () => {
    const dashed: ParagraphPaint = {
      leftPt: 36,
      rightPt: 576,
      fillColor: null,
      borders: {
        ...NO_BORDERS,
        top: { style: "dashed", color: "000000", widthPt: 1.5, spacePt: 0 },
      },
    };
    const page = bodyPage([boxOf([lineAt(36.1, 47.4258)], dashed)]);
    const lines = paintIn(page).painted.flatMap((each) => each.lines);

    expect(lines[0]?.dashes).toStrictEqual([6, 6]);
  });

  it("states no dashes for a line Word draws solid", () => {
    const solid: ParagraphPaint = {
      leftPt: 36,
      rightPt: 576,
      fillColor: null,
      borders: { ...NO_BORDERS, top: { style: "single", color: "000000", widthPt: 1, spacePt: 0 } },
    };
    const page = bodyPage([boxOf([lineAt(36.1, 47.4258)], solid)]);
    const lines = paintIn(page).painted.flatMap((each) => each.lines);

    expect(lines[0]?.dashes).toBeNull();
  });

  // Black is what Word draws text and borders it was told nothing about. The two
  // backends disagreed until this moved: the pdf writer drew black and the viewer
  // inherited whatever colour the page around it was set in.
  it("resolves a run that states no colour of its own to the black Word draws", () => {
    const page = bodyPage([boxOf([lineAt(36, 47.4258)])]);
    const text = textIn(page).boxes[0]?.lines[0]?.line.segments[0];

    expect(text?.kind === "text" && text.mark.color).toBe("#000000");
  });

  it("resolves a border stating no colour to the same black", () => {
    const unstated: ParagraphPaint = {
      leftPt: 36,
      rightPt: 576,
      fillColor: null,
      borders: { ...NO_BORDERS, top: { style: "single", color: null, widthPt: 1, spacePt: 0 } },
    };
    const page = bodyPage([boxOf([lineAt(36.1, 47.4258)], unstated)]);

    expect(paintIn(page).painted.flatMap((each) => each.lines)[0]?.color).toBe("#000000");
  });

  // A run held to the width it was measured at is stretched glyph by glyph only
  // where the file scaled it; otherwise the gaps between its glyphs are opened.
  it("says how a run's measured width is made up", () => {
    expect(runWidthMadeUpBy(MARK)).toBe("spacing");
    expect(runWidthMadeUpBy({ ...MARK, characterScale: 1.5 })).toBe("glyphs");
  });

  // Where a metafile's own pen stands, which both backends carried until this did.
  it("states how far a metafile's pen stands from the line it is told to draw", () => {
    expect(METAFILE_PEN_OFFSET).toBe(0.5);
  });
});

// The face a document sets its mathematics in, which is not the face the text
// around it is in: on the authored equation probes Word drew the markers in the body
// face and every piece of the equations in a second one.
const MATH_MARK: ParagraphMark = {
  ...MARK,
  font: { kind: "named", name: "Meridian Math" },
  fontSizePt: 11,
};

// A fraction as `mathPrimitivesOf` hands it over: a numerator, the bar, a
// denominator, at places no two of which land on the grid on their own. The numbers
// are the shape of a maths face's own answer at 11pt and are not measured off
// anything: what is held here is what the drawing does to them, and the measurement
// that settled that is in `drawables.ts` beside `drawnMathRun`.
const NUMERATOR: MathPrimitive = {
  kind: "text",
  text: "gralm",
  mark: MATH_MARK,
  sizePt: 7.92,
  leftPt: 290.5,
  baselinePt: 100.1,
};

const BAR: MathPrimitive = {
  kind: "fill",
  mark: MATH_MARK,
  leftPt: 290.4,
  topPt: 104.9,
  widthPt: 20.96,
  heightPt: 0.7223,
};

const DENOMINATOR: MathPrimitive = {
  kind: "text",
  text: "presk",
  mark: MATH_MARK,
  sizePt: 7.92,
  leftPt: 291.85,
  baselinePt: 115.87,
};

const FRACTION: readonly MathPrimitive[] = [NUMERATOR, BAR, DENOMINATOR];

// The same fraction as the line holds it, which is what `setMath` answers with: the
// geometry, and what is drawn hung on it. A line hands this to `mathPrimitivesOf`
// and the pieces above are what comes back.
const half = (text: string): SetMath => ({
  kind: "run",
  text,
  mark: MATH_MARK,
  sizePt: 7.92,
  box: { widthPt: 20.5, ascentPt: 5.6, descentPt: 0, insetPt: 0 },
});

const SET_FRACTION: SetMath = {
  kind: "fraction",
  mark: MATH_MARK,
  box: {
    widthPt: 20.96,
    ascentPt: 8,
    descentPt: 5,
    insetPt: 0.23,
    numerator: {
      widthPt: 20.5,
      ascentPt: 5.6,
      descentPt: 0,
      insetPt: 0,
      leftPt: 0.23,
      baselinePt: 5.02,
    },
    denominator: {
      widthPt: 20.5,
      ascentPt: 5.6,
      descentPt: 0,
      insetPt: 0,
      leftPt: 0.23,
      baselinePt: -5.6,
    },
    bar: { leftPt: 0, widthPt: 20.96, topPt: 3.1, thicknessPt: 0.7223 },
  },
  numerator: [half("gralm")],
  denominator: [half("presk")],
};

const equationLineAt = (topPt: number, baselinePt: number, offsetPt = 0): PlacedLine => ({
  line: {
    segments: [
      {
        kind: "equation",
        pieces: [SET_FRACTION],
        widthPt: 20.96,
        ascentPt: 8,
        descentPt: 5,
        offsetPt,
      },
    ],
    widthPt: 20.96 + offsetPt,
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

// A parenthesis the face grew to fit what it stands round, which is a glyph of its
// own with no character to ask for it by.
const RUNG: MathPrimitive = {
  kind: "glyph",
  glyph: 3436,
  mark: MATH_MARK,
  sizePt: 11,
  leftPt: 281.28,
  baselinePt: 100.1,
};

// How far the equation reaches either side of the line's own baseline, which is
// what the line measured and the only thing a glyph run needs that a piece does not
// carry.
const equationOf = (primitives: readonly MathPrimitive[]): SetEquation => ({
  primitives,
  ascentPt: 8,
  descentPt: 5,
});

const runsOf = (drawables: readonly Drawable[]): readonly DrawnRun[] =>
  drawables.flatMap((drawable) => (drawable.kind === "text" ? drawable.runs : []));

const fillsOf = (drawables: readonly Drawable[]): readonly PaintedFill[] =>
  drawables.flatMap((drawable) =>
    drawable.kind === "paint" ? drawable.painted.flatMap((each) => each.fills) : [],
  );

describe("a set equation reaching the page", () => {
  /**
   * **Every piece lands on the grid on its own, which is what Word does.** Measured
   * on 2026-08-14 off Word's own pdf of the two authored equation probes: all 301
   * text baselines there are a whole device unit and so are both edges of all 72
   * fills. The full statement is in `drawables.ts`.
   *
   * So the bar keeps its place against its halves because all three land on the one
   * grid, and **not** because the fraction is moved as one thing with its offsets
   * kept: 100.1 goes to 100.08 and 104.9 to 104.88, which are different distances.
   * A bar drawn from a snapped top at the thickness it was handed would be the fault
   * the other way, its foot a fraction of a unit off the grid its head sits on, so
   * **both edges are snapped and the bar is filled between them**: 104.88 down to
   * 105.6, where the 0.7223 it was handed would have ended at 105.6223.
   */
  it("puts each of a fraction's pieces on the grid, the bar between two snapped edges", () => {
    const drawables = mathDrawables(equationOf(FRACTION), "e");
    const [bar] = fillsOf(drawables);

    expect(runsOf(drawables).map((run) => run.baselinePt)).toStrictEqual([100.08, 115.92]);
    expect(bar?.topPt).toBe(104.88);
    expect((bar?.topPt ?? 0) + (bar?.heightPt ?? 0)).toBe(105.6);
    expect(bar).toMatchObject({ color: "#000000", leftPt: 290.4, widthPt: 20.96 });
  });

  // Across the page nothing is snapped: 45.2% of the lefts of those same 301
  // placements are whole units, which is the exact arithmetic the rest of the page
  // keeps and the same answer `onTheDeviceGrid` gives every other drawing.
  it("leaves where a piece stands across the page exactly where it was set", () => {
    const drawables = mathDrawables(equationOf(FRACTION), "e");

    expect(runsOf(drawables).map((run) => run.leftPt)).toStrictEqual([290.5, 291.85]);
    expect(fillsOf(drawables)[0]?.leftPt).toBe(290.4);
  });

  // A piece is drawn at the size the flattener set it at, which is not the size the
  // equation's own mark states: Word shrinks the halves of a fraction that shares
  // its line with ordinary text and leaves the bar at the stated size.
  it("draws a piece at its own size, in the equation's own face and colour", () => {
    const [numerator] = runsOf(mathDrawables(equationOf(FRACTION), "e"));

    expect(numerator?.mark.fontSizePt).toBe(7.92);
    expect(numerator?.mark.font).toStrictEqual({ kind: "named", name: "Meridian Math" });
    expect(numerator?.text).toBe("gralm");
    // Held to no width, because a piece states none: a set equation is drawn in the
    // very face it was measured in, so its own advances are the ones that draw it.
    expect(numerator?.widthPt).toBe(0);
  });

  // What a run states about spacing and scale would move the glyphs off the places
  // the arithmetic measured with the face's plain advances, and the bar's own width
  // is made of those same advances. A raise is already in the baseline it was given.
  it("draws no piece with the spacing, the scale or the raise a run may state", () => {
    const stated: ParagraphMark = {
      ...MATH_MARK,
      characterSpacingPt: 1.5,
      characterScale: 1.5,
      raisePt: 4,
      underline: true,
      highlight: "yellow",
    };
    const [numerator] = runsOf(mathDrawables(equationOf([{ ...NUMERATOR, mark: stated }]), "e"));

    expect(numerator?.mark.characterSpacingPt).toBe(0);
    expect(numerator?.mark.characterScale).toBe(1);
    expect(numerator?.mark.raisePt).toBe(0);
    expect(numerator?.mark.underline).toBe(false);
    expect(numerator?.mark.highlight).toBeNull();
  });

  /**
   * **A stretched delimiter needs nothing new.** It is a rung of the face's own
   * ladder, which is a glyph with no character at all, and that is the one thing the
   * glyph run was built for: the parenthesis Word stretches round a fraction reached
   * Word's own pdf as a glyph its ToUnicode calls `!`.
   */
  it("draws a stretched delimiter as the glyph run built for a shape with no character", () => {
    const drawables = mathDrawables(equationOf([RUNG]), "e");

    expect(drawables).toStrictEqual([
      {
        kind: "glyphs",
        key: "e-0",
        face: { name: "Meridian Math", bold: false, italic: false },
        // Set on the device grid as a run of text is: 11pt is 45 and five sixths of a
        // unit, and Word sets it at 46, which is 11.04.
        sizePt: 11.04,
        color: "#000000",
        // The equation's own reach, since a piece states no ink of its own.
        ascentPt: 8,
        descentPt: 5,
        glyphs: [
          {
            glyph: 3436,
            leftPt: 281.28,
            baselinePt: 100.08,
            // **Three things a piece does not carry yet**, each stated here so that
            // the day it does, this test says what changed: what the glyph advances,
            // what it stands for, and the outline a browser needs to draw it at all.
            advancePt: 0,
            standsFor: null,
          },
        ],
      },
    ]);
  });

  // **The order the flattener set the pieces in is the order they are painted.**
  // Gathering every fill of an equation into one layer would put a rule under a half
  // it was drawn over, so a change of kind opens another drawable and the two stay
  // where they were put.
  it("keeps the order the pieces were set in", () => {
    const drawables = mathDrawables(equationOf([BAR, NUMERATOR, BAR]), "e");

    expect(drawables.map((each) => each.kind)).toStrictEqual(["paint", "text", "paint"]);
  });

  // Pieces of one kind standing together are drawn together, so a fraction of two
  // halves is one text drawable and not two.
  it("draws pieces of one kind standing together in one drawable", () => {
    const halves = [NUMERATOR, DENOMINATOR];

    expect(mathDrawables(equationOf(halves), "e")).toHaveLength(1);
    expect(runsOf(mathDrawables(equationOf(halves), "e"))).toHaveLength(2);
  });

  // A run of glyphs is drawn at one size, so two rungs set at different sizes are
  // two runs however they stand.
  it("parts a run of glyphs at a change of size", () => {
    const rung = (sizePt: number): MathPrimitive => ({ ...RUNG, sizePt });

    expect(mathDrawables(equationOf([rung(11), rung(11)]), "e")).toHaveLength(1);
    expect(mathDrawables(equationOf([rung(11), rung(7.92)]), "e")).toHaveLength(2);
  });

  // An equation is text, and what stands over it and under it is what stands over
  // and under any other text.
  it("is drawn where the story's own text is drawn", () => {
    const page = bodyPage([boxOf([equationLineAt(36.1, 47.4258)])]);
    const kinds = drawablesOf(layoutOf(page), page).map((drawable) => drawable.kind);

    expect(kinds).toStrictEqual(["text", "text", "paint", "text"]);
  });

  /**
   * **An equation is found on the line that holds it**, and the line says where its
   * baseline is: the segment carries the pieces placed about that baseline and an
   * offset along the line, and nothing hangs off the page.
   *
   * The whole chain is in these four numbers. The line's own baseline lands on the
   * grid at 47.52; the numerator stands 5.02 above it and is drawn at 42.48, the bar
   * 3.1 above it and is filled from 44.4 down to 45.12, the denominator 5.6 below it
   * and is drawn at 53.04. Across the page the offset and the half's own place are
   * added and nothing is rounded: 72.137 + 10 + 0.23.
   */
  it("stands where the line put it, on the grid, piece by piece", () => {
    const page = bodyPage([boxOf([equationLineAt(36.1, 47.4258, 10)])]);
    const drawables = drawablesOf(layoutOf(page), page);

    expect(runsOf(drawables).map((run) => run.baselinePt)).toStrictEqual([42.48, 53.04]);
    expect(runsOf(drawables).map((run) => run.leftPt)).toStrictEqual([82.367, 82.367]);

    const [bar] = fillsOf(drawables);
    expect(bar?.topPt).toBe(44.4);
    expect((bar?.topPt ?? 0) + (bar?.heightPt ?? 0)).toBe(45.12);
    expect(bar?.leftPt).toBe(82.137);
  });

  // A line that set no equation draws none, and a page of plain text draws as it
  // always did.
  it("draws nothing for a line that holds no equation", () => {
    const page = bodyPage([boxOf([lineAt(36.1, 47.4258)])]);

    expect(drawablesOf(layoutOf(page), page).map((each) => each.kind)).toStrictEqual(["text"]);
  });
});
