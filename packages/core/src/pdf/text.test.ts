import { strFromU8, unzlibSync } from "fflate";
import { describe, expect, it } from "vitest";

import type { SectionGeometry } from "../docx/section.js";
import type { LaidOutDocument } from "../layout/document.js";
import type { PageDrawing, PlacedGlyphs } from "../layout/drawables.js";
import type { SetMath } from "../layout/math.js";
import type { ParagraphBox, PlacedLine } from "../layout/stack.js";
import type { ParagraphMark } from "../docx/styles.js";
import { buildSfnt } from "../testing/build-font.js";

import { contentOf } from "./content.js";
import type { PdfFont } from "./document.js";
import { pdfFonts } from "./fonts.js";
import { pdfImages } from "./images.js";
import { pdfObjects } from "./objects.js";

// A glyph the drawing names by number, which is the only way to ask for a shape
// with no character: the parenthesis Word stretches round a fraction is a variant
// glyph of Cambria Math and is in no character map at all.
//
// A pdf can draw one exactly, because it already embeds the face whole and
// addresses it by glyph. What has to be got right is what is written beside the
// drawing: the width of a glyph nothing can look up, and what it stands for.

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

const UNITS_PER_EM = 1000;

// Glyphs are numbered from one in code point order, so `(` is the first of these
// three and `A` the second.
const FILE = buildSfnt({
  unitsPerEm: UNITS_PER_EM,
  ascender: 800,
  descender: -200,
  lineGap: 0,
  advances: { "(": 400, A: 660, B: 640 },
});

const OPENING_GLYPH = 1;
const A_GLYPH = 2;

const SUPPLIED: PdfFont = { name: "Meridian Math", bytes: FILE };

const GROWN: PlacedGlyphs = {
  face: { name: "Meridian Math", bold: false, italic: false },
  sizePt: 22,
  color: "112233",
  ascentPt: 14,
  descentPt: 7.6,
  glyphs: [{ glyph: OPENING_GLYPH, leftPt: 100, baselinePt: 200, advancePt: 11, standsFor: "(" }],
};

function pageWith(
  glyphRuns: readonly PlacedGlyphs[],
  body: readonly ParagraphBox[] = [],
): PageDrawing {
  return {
    index: 0,
    geometry: LETTER,
    body,
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
    glyphRuns,
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

// The syntax is latin1, so a stream reads back one character to a byte.
const textOf = (bytes: Uint8Array): string => strFromU8(bytes, true);

function streamOf(
  glyphRuns: readonly PlacedGlyphs[],
  fonts: readonly PdfFont[] = [SUPPLIED],
  body: readonly ParagraphBox[] = [],
) {
  const page = pageWith(glyphRuns, body);
  const objects = pdfObjects();
  const written = pdfFonts(fonts);
  const images = pdfImages({
    imageBytes: () => undefined,
    metricsFor: () => ({ kind: "missing", fontName: "" }),
    fonts: written,
    objects,
  });

  const content = contentOf(layoutOf(page), page, {
    fonts: written,
    images,
    aliasSymbolFaces: null,
  });
  return { drawn: textOf(content.bytes), fonts: written, objects };
}

// Every object of the file, as one string, with the streams inflated so that the
// map of glyphs to characters can be read back out of it.
function filedAway(fonts: ReturnType<typeof pdfFonts>, objects: ReturnType<typeof pdfObjects>) {
  const root = objects.add(fonts.resources(objects));
  const bytes = objects.bytes({ root });
  const text = textOf(bytes);

  const streams = [...text.matchAll(/stream\n([\s\S]*?)\nendstream/g)].flatMap((found) => {
    const body = found[1] ?? "";
    try {
      return [textOf(unzlibSync(Uint8Array.from(body, (each) => each.charCodeAt(0))))];
    } catch {
      return [];
    }
  });
  return { text, streams };
}

describe("a glyph the drawing names by number", () => {
  // The baseline lands on the device grid before it reaches here, which
  // `drawables.ts` does once for everything a page draws: 200pt down the page is
  // 833 units and a third, so what is drawn is 199.92 and the flip of it is 592.08.
  it("is shown as its own two bytes, at the place the layout put it", () => {
    const { drawn } = streamOf([GROWN]);

    expect(drawn).toContain("/F0 22 Tf");
    expect(drawn).toContain("1 0 0 1 100 592.08 Tm");
    expect(drawn).toContain("<0001> Tj");
  });

  // Across the page nothing is rounded, which is Word's own arithmetic: only the
  // baseline lands on the grid.
  it("leaves the place across the page exactly where the layout put it", () => {
    const { drawn } = streamOf([
      {
        ...GROWN,
        glyphs: [
          {
            ...(GROWN.glyphs[0] ?? {
              glyph: 1,
              leftPt: 0,
              baselinePt: 0,
              advancePt: 0,
              standsFor: null,
            }),
            leftPt: 100.137,
          },
        ],
      },
    ]);

    expect(drawn).toContain("1 0 0 1 100.137 ");
  });

  it("is drawn in the colour the run states", () => {
    const { drawn } = streamOf([GROWN]);

    expect(drawn).toContain("0.066667 0.133333 0.2 rg");
  });

  // A stretched delimiter takes neither spacing nor scaling, and a page whose last
  // run asked for either would carry it into this one: a pdf reader holds both
  // until something says otherwise.
  it("states the spacing and the scale rather than inheriting a run's", () => {
    const { drawn } = streamOf([GROWN]);

    expect(drawn).toContain("0 Tc");
    expect(drawn).toContain("100 Tz");
  });

  it("writes each glyph of a run at its own place", () => {
    const { drawn } = streamOf([
      {
        ...GROWN,
        glyphs: [
          ...GROWN.glyphs,
          { glyph: A_GLYPH, leftPt: 140, baselinePt: 260, advancePt: 14.52, standsFor: null },
        ],
      },
    ]);

    expect(drawn).toContain("1 0 0 1 140 532.08 Tm");
    expect(drawn).toContain("<0002> Tj");
  });

  // A backend that embeds the face names the glyph rather than drawing its
  // outline: the embedded one is hinted, it is selectable, and it is the same
  // shape by construction. The outline is there for a backend that cannot.
  it("names the glyph even where the run carries the shape as well", () => {
    const outlined: PlacedGlyphs = {
      ...GROWN,
      glyphs: [
        {
          ...(GROWN.glyphs[0] ?? {
            glyph: OPENING_GLYPH,
            leftPt: 0,
            baselinePt: 0,
            advancePt: 0,
            standsFor: null,
          }),
          outline: { unitsPerEm: 1000, contours: [] },
        },
      ],
    };
    const { drawn } = streamOf([outlined]);

    expect(drawn).toContain("<0001> Tj");
    expect(drawn).not.toContain(" re");
  });

  it("draws nothing at all for a page that names no glyph", () => {
    expect(streamOf([]).drawn).not.toContain("Tj");
  });

  // **The width is the one the layout measured**, and it is the one width nothing
  // here could look up afterwards: a glyph with no character has no advance to ask
  // the advance table for.
  it("writes the width the run stated, in glyph space", () => {
    const { fonts, objects } = streamOf([GROWN]);
    const { text } = filedAway(fonts, objects);

    // 11pt drawn at 22pt is half the em, which is 500 thousandths of it.
    expect(text).toContain("/W [1 [500]]");
  });

  // The plain `(` of this face advances 400 units. The stretched one is drawn at
  // half the em, and it is the drawn width that has to be written: taking the
  // character's own would say the shape is the size it is not.
  it("keeps the stated width ahead of what the character it stands for advances", () => {
    const { fonts, objects } = streamOf([GROWN]);
    const { text } = filedAway(fonts, objects);

    expect(text).not.toContain("[400]");
  });

  /**
   * A glyph named by number is mapped to the character it stands for, so the text
   * can still be selected and searched. **Word's own pdf does not do this**: its
   * grown delimiters came back out of the text layer as `.` and `/`, whatever its
   * subset happened to number them.
   */
  it("says which character the glyph stands for", () => {
    const { fonts, objects } = streamOf([GROWN]);
    const { streams } = filedAway(fonts, objects);

    expect(streams.some((each) => each.includes("<0001> <0028>"))).toBe(true);
  });

  it("leaves a glyph standing for nothing out of the map, and draws it all the same", () => {
    const nameless: PlacedGlyphs = {
      ...GROWN,
      glyphs: [{ glyph: A_GLYPH, leftPt: 10, baselinePt: 20, advancePt: 11, standsFor: null }],
    };
    const { drawn, fonts, objects } = streamOf([nameless]);
    const { text, streams } = filedAway(fonts, objects);

    expect(drawn).toContain("<0002> Tj");
    expect(text).toContain("/W [2 [500]]");
    expect(streams.some((each) => each.includes("<0002>"))).toBe(false);
  });

  // The face is embedded whole and addressed by glyph whatever is drawn, so naming
  // a glyph costs nothing a letter does not.
  it("draws a named glyph and a character out of the one embedded face", () => {
    const { fonts, objects } = streamOf([GROWN]);
    const face = fonts.faceFor({ name: "Meridian Math", bold: false, italic: false });
    face.glyphsFor("A");
    const { text } = filedAway(fonts, objects);

    // Consecutive glyphs are written as one run of widths, which is what the two
    // are here: the stretched shape at the width it was drawn, and the letter at
    // the width its own character advances.
    expect(face.resource).toBe("F0");
    expect(text).toContain("/W [1 [500 660]]");
  });
});

// A set equation, which reaches a backend as the three things `drawables.ts` makes
// of it and not as a fourth kind of drawing: a piece of text at a place, a fill for
// the bar, and a glyph named by number for a stretched delimiter.
//
// What is held here is the narrower thing the writer is answerable for, that each
// of the three is written where the drawing put it. Where the drawing puts them,
// and the measurement behind that, is `drawables.ts`.

const EQUATION_MARK: ParagraphMark = {
  font: { kind: "named", name: "Meridian Math" },
  fontSizePt: 11,
  bold: false,
  italic: false,
  underline: false,
  raisePt: 0,
  lineSizePt: 11,
  lineRaisePt: 0,
  color: null,
  characterSpacingPt: 0,
  characterScale: 1,
  kernFromHalfPoints: null,
  highlight: null,
  capitals: "none",
};

// A fraction as the line holds it: two halves of a face the file embeds, and the
// bar between them. `math.ts` places the pieces about the line's own baseline and
// `drawables.ts` turns them into the three things drawn below.
const half = (text: string): SetMath => ({
  kind: "run",
  text,
  mark: EQUATION_MARK,
  sizePt: 7.92,
  box: { widthPt: 5.2, ascentPt: 5.6, descentPt: 0 },
});

const SET_FRACTION: SetMath = {
  kind: "fraction",
  mark: EQUATION_MARK,
  box: {
    widthPt: 5.2,
    ascentPt: 8,
    descentPt: 5,
    numerator: { widthPt: 5.2, ascentPt: 5.6, descentPt: 0, leftPt: 0, baselinePt: 5.02 },
    denominator: { widthPt: 5.2, ascentPt: 5.6, descentPt: 0, leftPt: 1.35, baselinePt: -5.6 },
    bar: { leftPt: 0, widthPt: 20.96, topPt: 3.1, thicknessPt: 0.7223 },
  },
  numerator: [half("A")],
  denominator: [half("B")],
};

const EQUATION_LINE: PlacedLine = {
  line: {
    segments: [
      {
        kind: "equation",
        pieces: [SET_FRACTION],
        widthPt: 5.2,
        ascentPt: 8,
        descentPt: 5,
        offsetPt: 0,
      },
    ],
    widthPt: 5.2,
    heightPt: 14.6484375,
    ascentPt: 11.7,
    seatPt: 0,
    fontHeightPt: 14.6484375,
    heldOpenPt: null,
  },
  leftPt: 290.5,
  topPt: 90,
  heightPt: 14.6484375,
  seatPt: 0,
  fittingHeightPt: 14.6484375,
  baselinePt: 100.1,
  startsPage: false,
};

const EQUATION_BOX: ParagraphBox = {
  index: 0,
  topPt: 90,
  anchorTopPt: 90,
  heightPt: 14.6484375,
  lines: [EQUATION_LINE],
  marker: null,
  contentWidthPt: 5.2,
  markTopPt: 90,
  contentBottomPt: 104.6484375,
  resumesUnderPt: 0,
  widowControl: false,
  keepNext: false,
  startsPage: false,
  endsPage: false,
  endsPageAtASection: false,
  clipTo: null,
  paint: null,
};

describe("a fraction reaching the page", () => {
  // Each half at the size the flattener set it at, on the baseline that size was set
  // on. The line's own baseline of 100.1 lands on the grid at 100.08; the numerator
  // stands 5.02 above it, is drawn at 95.04 and flips to 696.96, and the denominator
  // stands 5.6 below it, is drawn at 105.6 and flips to 686.4.
  it("writes each half at its own size on its own snapped baseline", () => {
    const { drawn } = streamOf([], [SUPPLIED], [EQUATION_BOX]);

    expect(drawn).toContain("/F0 7.92 Tf");
    expect(drawn).toContain("1 0 0 1 290.5 696.96 Tm");
    expect(drawn).toContain("<0002> Tj");
    expect(drawn).toContain("1 0 0 1 291.85 686.4 Tm");
    expect(drawn).toContain("<0003> Tj");
  });

  // The bar is filled, as Word fills it, between the two snapped edges: 3.1 above the
  // line's baseline is 96.96 and the foot of it 97.68, which the writer flips into a
  // rectangle standing at 694.32 and exactly 0.72 tall.
  it("fills the bar between the two edges the drawing snapped", () => {
    const { drawn } = streamOf([], [SUPPLIED], [EQUATION_BOX]);

    expect(drawn).toContain("290.5 694.32 20.96 0.72 re");
  });

  // A stretched delimiter is the glyph run, which this file already draws: what is
  // held here is that an equation reaches it, and at the size the piece was set at.
  it("writes a stretched delimiter as a glyph named by number", () => {
    const grown: SetMath = {
      kind: "delimiter",
      mark: EQUATION_MARK,
      sizePt: 22,
      box: {
        widthPt: 11,
        ascentPt: 14,
        descentPt: 7.6,
        opening: {
          codePoint: 0x28,
          variant: { glyph: OPENING_GLYPH, measurement: 4047, advance: 500, ink: null },
          grown: true,
          widthPt: 11,
          ascentPt: 14,
          descentPt: 7.6,
          leftPt: 0,
          baselinePt: 0,
        },
        closing: null,
        content: { widthPt: 0, ascentPt: 0, descentPt: 0, leftPt: 11, baselinePt: 0 },
        setAsASubFormula: true,
        grownShort: false,
      },
      content: [],
    };
    const line: PlacedLine = {
      ...EQUATION_LINE,
      line: {
        ...EQUATION_LINE.line,
        segments: [
          {
            kind: "equation",
            pieces: [grown],
            widthPt: 11,
            ascentPt: 14,
            descentPt: 7.6,
            offsetPt: 0,
          },
        ],
      },
    };
    const { drawn } = streamOf([], [SUPPLIED], [{ ...EQUATION_BOX, lines: [line] }]);

    expect(drawn).toContain("/F0 22 Tf");
    expect(drawn).toContain("1 0 0 1 290.5 691.92 Tm");
    expect(drawn).toContain("<0001> Tj");
  });
});
