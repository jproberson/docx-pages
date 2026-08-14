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
  METAFILE_PEN_OFFSET,
  onTheDeviceGrid,
  runWidthMadeUpBy,
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
