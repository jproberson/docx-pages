import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import { NO_PAINT, UNPAINTED, WHOLE_FRAME } from "@docx-pages/core";
import type {
  CropInsets,
  LaidOutDocument,
  LaidOutPage,
  MetafilePicture,
  ParagraphBox,
  ParagraphMark,
  PlacedCell,
  PlacedContent,
  PlacedPaint,
  SectionGeometry,
  TextBoxBody,
} from "@docx-pages/core";

import type { DrawableImage } from "./images.js";
import { Page } from "./page.js";

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

const EMPTY_BODY: TextBoxBody = {
  blocks: [],
  insets: { leftEmu: 0, topEmu: 0, rightEmu: 0, bottomEmu: 0 },
  anchor: "top",
  wraps: true,
  fitsText: false,
};

const textBox = (boxes: readonly ParagraphBox[] = []): PlacedContent => ({
  kind: "text-box",
  body: EMPTY_BODY,
  text:
    boxes.length === 0
      ? null
      : { boxes, cells: [], inlines: [], contentHeightPt: 0, contentWidthPt: 0 },
  paint: UNPAINTED,
});

const paragraphOf = (
  text: string,
  mark: ParagraphMark = MARK,
  options: { leftPt?: number; baselinePt?: number; widthPt?: number; offsetPt?: number } = {},
): ParagraphBox => ({
  index: 0,
  topPt: 0,
  anchorTopPt: 0,
  resumesUnderPt: 0,
  heightPt: 14,
  lines: [
    {
      line: {
        segments: [
          {
            kind: "text",
            mark,
            text,
            widthPt: options.widthPt ?? 40,
            offsetPt: options.offsetPt ?? 0,
          },
        ],
        widthPt: options.widthPt ?? 40,
        heightPt: 14,
        ascentPt: 11,
        seatPt: 0,
        fontHeightPt: 14,
        heldOpenPt: null,
      },
      leftPt: options.leftPt ?? 120,
      topPt: 30,
      heightPt: 14,
      seatPt: 0,
      fittingHeightPt: 14,
      baselinePt: options.baselinePt ?? 41,
      startsPage: false,
    },
  ],
  marker: null,
  markTopPt: 30,
  contentBottomPt: 44,
  widowControl: true,
  keepNext: false,
  startsPage: false,
  endsPage: false,
  endsPageAtASection: false,
  contentWidthPt: options.widthPt ?? 40,
  clipTo: null,
  paint: null,
});

const LETTER: SectionGeometry = {
  widthTwips: 12240,
  heightTwips: 15840,
  margin: {
    topTwips: 720,
    rightTwips: 720,
    bottomTwips: 0,
    leftTwips: 720,
    headerTwips: 432,
    footerTwips: 144,
  },
};

const NO_CROP: CropInsets = { left: 0, top: 0, right: 0, bottom: 0 };

const float = (content: PlacedContent, options: { behindDoc?: boolean; height?: number } = {}) => ({
  anchor: {
    paragraphIndex: 0,
    name: "Object",
    widthEmu: 0,
    heightEmu: 0,
    turnDegrees: 0,
    flip: { horizontal: false, vertical: false } as const,
    content: { kind: "shape", paint: NO_PAINT } as const,
    horizontal: { kind: "offset", from: "column", offsetEmu: 0 } as const,
    vertical: { kind: "offset", from: "paragraph", offsetEmu: 0 } as const,
    wrap: "none" as const,
    side: "bothSides" as const,
    area: WHOLE_FRAME,
    distances: { topEmu: 0, rightEmu: 0, bottomEmu: 0, leftEmu: 0 },
    behindDoc: options.behindDoc ?? false,
    relativeHeight: options.height ?? 0,
  },
  content,
  leftPt: 100,
  topPt: 200,
  widthPt: 180,
  heightPt: 90,
  turnDegrees: 0,
  flip: { horizontal: false, vertical: false } as const,
});

const layoutWith = (
  floats: readonly ReturnType<typeof float>[],
  body: readonly ParagraphBox[] = [],
  footerFloats: readonly ReturnType<typeof float>[] = [],
  cells: readonly PlacedCell[] = [],
): LaidOutDocument => ({
  kind: "laid-out",
  page: LETTER,
  unhonoured: [],
  headerTopPt: 21.6,
  bodyTopPt: 36,
  bodyBottomPt: 792,
  pages: [
    {
      index: 0,
      geometry: LETTER,
      body,
      cells,
      floats,
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
      footerFloats,
      headerInlines: [],
      footerInlines: [],
    },
  ],
});

// A block of colour cut to a rectangle, a rule and a run of text, which is what a
// diagram recorded as a metafile is made of.
const PICTURE: MetafilePicture = {
  widthUnits: 40,
  heightUnits: 20,
  shapes: [
    {
      kind: "fill",
      rect: { leftUnits: 2, topUnits: 3, widthUnits: 10, heightUnits: 4 },
      color: "#92d050",
      clipTo: { leftUnits: 0, topUnits: 0, widthUnits: 30, heightUnits: 15 },
    },
    {
      kind: "text",
      text: "ab",
      xUnits: [5, 16],
      baselineUnits: 17,
      emUnits: 8,
      color: "#000000",
      face: { name: "Meridian Sans", bold: false, italic: false },
      clipTo: null,
    },
  ],
};

const DRAWABLE: ReadonlyMap<string, DrawableImage> = new Map([
  ["word/media/image1.png", { kind: "bitmap", url: "data:image/png;base64,AA" }],
  ["word/media/image2.emf", { kind: "metafile", picture: PICTURE }],
]);

const firstPage = (layout: LaidOutDocument): LaidOutPage => {
  const [page] = layout.pages;
  if (page === undefined) throw new Error("the layout has no pages");
  return page;
};

const markup = (
  layout: LaidOutDocument,
  options: Partial<Parameters<typeof Page>[0]> | null = null,
) =>
  renderToStaticMarkup(
    <Page
      layout={layout}
      page={firstPage(layout)}
      imageUrl={(part) => DRAWABLE.get(part)}
      {...(options === null ? {} : options)}
    />,
  );

describe("Page", () => {
  it("draws a stood-in symbol face's positions as what they mean, only when told to", () => {
    const wingdings: ParagraphMark = { ...MARK, font: { kind: "named", name: "Wingdings" } };
    const layout = layoutWith([], [paragraphOf("l", wingdings)]);

    // Told the face was stood in for, the position paints as its meaning; not
    // told, the face is really there and the run paints as written.
    expect(markup(layout, { aliasSymbolFaces: new Set(["wingdings"]) })).toContain("●");
    expect(markup(layout)).toContain(">l</tspan>");
  });

  it("sizes the page in points so everything below stays in Word's coordinates", () => {
    const html = markup(layoutWith([]));
    expect(html).toContain("width:612pt");
    expect(html).toContain("height:792pt");
  });

  it("sizes the text layer in points too, so a glyph is drawn at the size it was measured at", () => {
    const html = markup(layoutWith([], [paragraphOf("Hello", MARK)]));
    expect(html).toContain('width="612pt"');
    expect(html).toContain('height="792pt"');
    expect(html).toContain('viewBox="0 0 612 792"');
  });

  it("scales the whole page rather than each object in it", () => {
    const html = markup(layoutWith([]), { scale: 0.5 });
    expect(html).toContain("transform:scale(0.5)");
    expect(html).toContain("transform-origin:top left");
  });

  it("puts a picture at the point position layout computed", () => {
    const html = markup(
      layoutWith([
        float({ kind: "picture", part: "word/media/image1.png", crop: NO_CROP, paint: UNPAINTED }),
      ]),
    );
    expect(html).toContain("left:100pt");
    expect(html).toContain("top:200pt");
    expect(html).toContain('src="data:image/png;base64,AA"');
  });

  it("draws the whole bitmap behind a window when srcRect crops it", () => {
    const crop: CropInsets = { left: 0, top: 0.07272, right: 0.293, bottom: 0 };
    const html = markup(
      layoutWith([
        float({ kind: "picture", part: "word/media/image1.png", crop, paint: UNPAINTED }),
      ]),
    );
    // The window stays at the placed size; the bitmap behind it is 180 / (1 - 0.293)
    // wide and rides 7.272% of its own height above the top edge.
    expect(html).toContain("width:180pt;height:90pt;overflow:hidden");
    expect(html).toMatch(/width:254\.596\d*pt/);
    expect(html).toMatch(/top:-7\.058\d*pt/);
  });

  // A metafile is a recording of the drawing, so its shapes go straight into the
  // frame the document gave it and the frame decides the scale on its own.
  it("draws a metafile's own shapes in the frame the document gave it", () => {
    const html = markup(
      layoutWith([
        float({ kind: "picture", part: "word/media/image2.emf", crop: NO_CROP, paint: UNPAINTED }),
      ]),
    );
    expect(html).toContain('viewBox="0 0 40 20" preserveAspectRatio="none"');
    expect(html).toContain("width:180pt;height:90pt");
    expect(html).toContain('<rect x="2" y="3" width="10" height="4" fill="#92d050"');
    expect(html).toContain('<text x="5 16" y="17"');
    expect(html).toContain('font-size="8"');
  });

  it("cuts a metafile's shape to the rectangle the recording clipped it to", () => {
    const html = markup(
      layoutWith([
        float({ kind: "picture", part: "word/media/image2.emf", crop: NO_CROP, paint: UNPAINTED }),
      ]),
    );
    expect(html).toContain('<clipPath id="float-0-clip-0">');
    expect(html).toContain('clip-path="url(#float-0-clip-0)"');
  });

  // A source rectangle hides a fraction of each edge, which for a recording is a
  // narrower window onto the same coordinates rather than a larger picture behind
  // a smaller frame.
  it("narrows the window onto a metafile when srcRect crops it", () => {
    const crop: CropInsets = { left: 0.25, top: 0, right: 0, bottom: 0.5 };
    const html = markup(
      layoutWith([
        float({ kind: "picture", part: "word/media/image2.emf", crop, paint: UNPAINTED }),
      ]),
    );
    expect(html).toContain('viewBox="10 0 30 10"');
  });

  it("shows nothing for a text box holding no text", () => {
    expect(markup(layoutWith([float(textBox())]))).not.toContain("<svg");
  });

  it("fills a shape at the size it was placed at", () => {
    const paint: PlacedPaint = { ...UNPAINTED, fillColor: "#F2F2F2" };
    const html = markup(layoutWith([float({ kind: "shape", paint })]));
    expect(html).toContain('fill="#F2F2F2"');
    expect(html).toContain('width="180" height="90"');
  });

  // A line shape is stored with no height at all, so the layer it draws in has to
  // be grown by the stroke to have any room to draw the line in.
  it("draws a line shape as a line, which a rectangle of no height could not be", () => {
    const paint: PlacedPaint = {
      geometry: "line",
      path: null,
      fillColor: null,
      outline: { color: "#BFBFBF", widthPt: 0.75 },
    };
    const html = markup(
      layoutWith([{ ...float({ kind: "shape", paint }), heightPt: 0, widthPt: 400 }]),
    );
    expect(html).toContain('stroke="#BFBFBF"');
    expect(html).toContain('x2="400"');
    expect(html).toContain("height:1.5pt");
  });

  // A path the file drew point by point, in shares of its own box, which the viewer
  // puts into the box the object was given. The shares are core's, so this and the
  // pdf writer draw one shape rather than two.
  it("draws a path the file drew point by point", () => {
    const paint: PlacedPaint = {
      geometry: "custom",
      fillColor: "#F2F2F2",
      outline: null,
      path: [
        { kind: "move", to: { x: 0.5, y: 0 } },
        { kind: "line", to: { x: 1, y: 1 } },
        { kind: "curve", first: { x: 0.5, y: 1 }, second: { x: 0, y: 0.5 }, to: { x: 0, y: 0 } },
        { kind: "close" },
      ],
    };
    const html = markup(layoutWith([float({ kind: "shape", paint })]));

    expect(html).toContain('d="M 90 0 L 180 90 C 90 90 0 45 0 0 Z"');
    expect(html).toContain('fill="#F2F2F2"');
  });

  it("draws a path it cannot play as nothing at all, not as the box it fits in", () => {
    const paint: PlacedPaint = { ...UNPAINTED, geometry: "custom", fillColor: "#F2F2F2" };
    const html = markup(layoutWith([float({ kind: "shape", paint })]));

    expect(html).not.toContain("#F2F2F2");
  });

  // **A run the file scaled is stretched rather than spaced out.** Every other run
  // is held to its measured width by the gaps between its glyphs, which is what
  // keeps Word's break points under a substituted face; a run stating `w:w` is drawn
  // wider or narrower glyph by glyph, as the pdf writer's `Tz` draws it.
  it("stretches a scaled run's glyphs and spaces every other run's", () => {
    const scaled: ParagraphMark = { ...MARK, characterScale: 1.5 };
    const html = markup(layoutWith([], [paragraphOf("wide", scaled), paragraphOf("plain")]));

    expect(html).toContain('lengthAdjust="spacingAndGlyphs"');
    expect(html).toContain('lengthAdjust="spacing"');
  });

  // Measured against Word, which cuts a line off mid-glyph where it runs past the
  // box holding it rather than moving it or letting it overrun.
  it("cuts a shape's text off at the frame, and lets a story's own text run on", () => {
    const html = markup(layoutWith([float(textBox([paragraphOf("held")]))], [paragraphOf("free")]));
    expect(html).toContain('width="180pt" height="90pt" viewBox="100 200 180 90"');
    expect(html).toContain('width="612pt" height="792pt" viewBox="0 0 612 792"');
    expect(html).toContain("overflow:visible");
  });

  it("outlines frames on request so placement can be checked without content", () => {
    const html = markup(layoutWith([float(textBox())]), { frames: "outlined" });
    expect(html).toContain('data-kind="text-box"');
    expect(html).toContain("left:100pt");
  });

  it("marks a picture whose part never resolved rather than drawing a broken image", () => {
    const html = markup(layoutWith([float({ kind: "missing-picture", relationshipId: "rId7" })]), {
      frames: "outlined",
    });
    expect(html).toContain('data-kind="missing-picture"');
  });

  // Measured against Word, which draws a panel it was told to send to the back of
  // the stack under a text box marked behindDoc: within one story the height a
  // shape was given is the whole of the order.
  it("stacks a story's objects by the height each was given, whatever behindDoc says", () => {
    const html = markup(
      layoutWith([
        float({ kind: "shape", paint: UNPAINTED }, { height: 5 }),
        float(textBox(), { behindDoc: true, height: 9 }),
      ]),
      { frames: "outlined" },
    );
    expect(html.indexOf('data-kind="shape"')).toBeLessThan(html.indexOf('data-kind="text-box"'));
  });

  // A panel anchored in the body covers the footer's own classification line on
  // the first page, which is why Word shows that line on the second page alone.
  it("draws the footer's own objects under the body's, whatever heights they carry", () => {
    const html = markup(
      layoutWith(
        [float({ kind: "shape", paint: UNPAINTED }, { height: 1 })],
        [],
        [float(textBox(), { height: 99 })],
      ),
      { frames: "outlined" },
    );
    expect(html.indexOf('data-kind="text-box"')).toBeLessThan(html.indexOf('data-kind="shape"'));
  });
});

describe("Page drawing text", () => {
  it("draws each line at the baseline layout gave it", () => {
    const html = markup(layoutWith([], [paragraphOf("Hello")]));

    expect(html).toContain('data-kind="text"');
    expect(html).toContain('x="120"');
    // 41pt down the page is 170 device units and five sixths, so the line is drawn
    // at 41.04: `drawables.ts` puts every drawn baseline on Word's own grid.
    expect(html).toContain('y="41.04"');
    expect(html).toContain("Hello");
  });

  it("underlines a run that asks for it, and leaves the rest alone", () => {
    const linked: ParagraphMark = { ...MARK, underline: true };
    expect(markup(layoutWith([], [paragraphOf("a link", linked)]))).toContain(
      'text-decoration="underline"',
    );
    expect(markup(layoutWith([], [paragraphOf("plain")]))).not.toContain("text-decoration");
  });

  it("names the authored face first and lets the page fall back behind it", () => {
    const html = markup(layoutWith([], [paragraphOf("Hello")]), {
      fallbackFonts: "Open Sans, sans-serif",
    });

    expect(html).toContain("&quot;Meridian Sans&quot;, Open Sans, sans-serif");
  });

  // Holding the drawn run to the width it was measured at is what keeps a
  // substituted face from wrapping somewhere Word did not.
  it("holds each run to the width it was measured at", () => {
    const html = markup(layoutWith([], [paragraphOf("Hello", MARK, { widthPt: 40 })]));
    expect(html).toContain('textLength="40"');
    expect(html).toContain('lengthAdjust="spacing"');
  });

  // A tab leaves a gap no run after it can find by adding up widths, so each run
  // is drawn at the place layout gave it rather than after the one before.
  it("starts each run where layout put it, not after the run before it", () => {
    const html = markup(layoutWith([], [paragraphOf("Hello", MARK, { offsetPt: 25 })]));
    expect(html).toContain('x="145"');
  });

  it("lifts a run off the line's baseline when its mark asks to be raised", () => {
    const raised: ParagraphMark = { ...MARK, raisePt: 4 };
    const html = markup(layoutWith([], [paragraphOf("2", raised, { baselinePt: 41 })]));

    expect(html).toContain('y="37.04"');
  });

  it("carries a run's own weight, slant, size and colour onto the page", () => {
    const bold: ParagraphMark = { ...MARK, bold: true, italic: true, color: "#FF0000" };
    const html = markup(layoutWith([], [paragraphOf("Hello", bold)]));

    expect(html).toContain('font-weight="bold"');
    expect(html).toContain('font-style="italic"');
    expect(html).toContain('font-size="12"');
    expect(html).toContain('fill="#FF0000"');
  });

  it("draws a text box's own text as well as the text that flows on the page", () => {
    const html = markup(layoutWith([float(textBox([paragraphOf("Inside")]))]));
    expect(html).toContain("Inside");
  });

  it("draws a text box's text after the frame it belongs to", () => {
    const html = markup(layoutWith([float(textBox([paragraphOf("Inside")]))]), {
      frames: "outlined",
    });
    expect(html.indexOf('data-kind="text-box"')).toBeLessThan(html.indexOf("Inside"));
  });

  it("draws no text layer for a document that has none", () => {
    expect(markup(layoutWith([]))).not.toContain('data-kind="text"');
  });
});

const CELL: PlacedCell = {
  leftPt: 100,
  topPt: 200,
  widthPt: 72,
  heightPt: 20,
  fillColor: "#DEEBF7",
  borders: {
    top: { style: "single", widthPt: 1.5, color: "#FF0000", spacePt: 0 },
    left: null,
    bottom: null,
    right: { style: "dashed", widthPt: 1, color: null, spacePt: 0 },
  },
};

describe("Page drawing a table's own lines", () => {
  it("lays a cell's colour and draws its lines", () => {
    const html = markup(layoutWith([], [], [], [CELL]));
    expect(html).toContain('data-kind="paint"');
    expect(html).toContain('fill="#DEEBF7"');
    expect(html).toContain('stroke="#FF0000"');
  });

  it("cuts a dashed line into the dashes Word draws it with", () => {
    expect(markup(layoutWith([], [], [], [CELL]))).toContain('stroke-dasharray="4 4"');
  });

  it("draws the whole of it behind the text", () => {
    const html = markup(layoutWith([], [paragraphOf("Hello")], [], [CELL]));
    expect(html.indexOf('data-kind="paint"')).toBeLessThan(html.indexOf('data-kind="text"'));
  });

  it("draws no layer at all for a page with nothing painted on it", () => {
    expect(markup(layoutWith([], [paragraphOf("Hello")]))).not.toContain('data-kind="paint"');
  });
});

const numbered = (text: string, markerText: string): ParagraphBox => ({
  ...paragraphOf(text),
  marker: {
    text: markerText,
    mark: { ...MARK, font: { kind: "named", name: "Wingdings" } },
    widthPt: 8,
    leftPt: 102,
    baselinePt: 41,
  },
});

describe("Page drawing a list number", () => {
  it("draws the number at the position layout gave it", () => {
    const html = markup(layoutWith([], [numbered("Item", "1.")]));
    expect(html).toContain('x="102" y="41.04"');
    expect(html).toContain(">1.<");
  });

  it("draws the number in its own face, ahead of the line it belongs to", () => {
    const html = markup(layoutWith([], [numbered("Item", "1.")]));
    expect(html).toContain('font-family="&quot;Wingdings&quot;, sans-serif"');
    expect(html.indexOf(">1.<")).toBeLessThan(html.indexOf(">Item<"));
  });

  it("draws nothing for a level that asks for no number at all", () => {
    const html = markup(layoutWith([], [numbered("Item", "")]));
    expect(html).not.toContain('x="102"');
  });
});

// A glyph the drawing names by number, which is the one thing a page can ask for
// that a browser cannot draw: css names a face and the browser picks the glyph,
// and no attribute of an svg `text` asks for glyph 3436 of a family.
const GROWN = {
  face: { name: "Meridian Math", bold: false, italic: false },
  sizePt: 22,
  color: "112233",
  ascentPt: 14,
  descentPt: 7.5,
  glyphs: [{ glyph: 3436, leftPt: 100, baselinePt: 200, advancePt: 11, standsFor: "(" }],
};

// The runs hang off the page, which the layout will state for itself once the
// seam that lays an equation out is built; a page carrying none draws none.
const withGlyphs = (): LaidOutDocument => {
  const layout = layoutWith([]);
  const page = { ...firstPage(layout), glyphRuns: [GROWN] };
  return { ...layout, pages: [page] };
};

describe("Page meeting a glyph it cannot draw", () => {
  it("marks the room the glyphs take and says which they were", () => {
    const html = markup(withGlyphs());

    expect(html).toContain('data-kind="glyphs"');
    expect(html).toContain('data-glyphs="3436"');
    expect(html).toContain('data-undrawn-glyphs="3436"');
  });

  // **Not the character it stands for.** A stretched parenthesis drawn as a plain
  // one is the right character at the wrong height, which is a page that looks
  // finished and is wrong.
  it("draws nothing in that room, the character the glyph stands for least of all", () => {
    const html = markup(withGlyphs());

    expect(html).not.toContain("<path");
    expect(html).not.toContain("Meridian Math");
    // The character it stands for is written for a reader to search by and
    // painted nowhere, which is the pdf writer's answer in the other notation.
    expect(html).toContain('fill="none"');
  });

  // The ink the layout stated, which is what says where the shape would have been:
  // 11pt of advance across, and the ascent and the descent about the baseline.
  it("puts the mark where the glyphs would have stood", () => {
    const html = markup(withGlyphs());

    expect(html).toContain("left:100pt");
    // The glyphs' own baseline lands on the grid like any other drawn line, so the
    // room marked for them starts from 199.92 rather than 200.
    expect(html).toContain("top:185.92pt");
    expect(html).toContain('width="11pt"');
    expect(html).toContain('height="21.5pt"');
  });

  it("outlines the room only where the page is asked to outline what it cannot draw", () => {
    expect(markup(withGlyphs(), { frames: "outlined" })).toContain('stroke-dasharray="3 3"');
    expect(markup(withGlyphs())).not.toContain('stroke-dasharray="3 3"');
  });
});

// A glyph reaching the page with the shape the face draws it as. **This is the one
// way a browser can draw a glyph with no character**: it addresses a face by
// character and by nothing else, and an svg draws a path as well as it draws a
// letter.
const OUTLINED = {
  ...GROWN,
  sizePt: 1000,
  glyphs: [
    {
      ...(GROWN.glyphs[0] ?? { glyph: 0, leftPt: 0, baselinePt: 0, advancePt: 0, standsFor: null }),
      outline: {
        unitsPerEm: 1000,
        contours: [
          {
            from: [0, 0] as const,
            steps: [
              { kind: "line" as const, to: [10, 0] as const },
              { kind: "quadratic" as const, control: [20, 5] as const, to: [10, 10] as const },
              {
                kind: "cubic" as const,
                first: [5, 12] as const,
                second: [2, 12] as const,
                to: [0, 10] as const,
              },
            ],
          },
        ],
      },
    },
  ],
};

const withOutlines = (): LaidOutDocument => {
  const layout = layoutWith([]);
  const page = { ...firstPage(layout), glyphRuns: [OUTLINED] };
  return { ...layout, pages: [page] };
};

describe("Page drawing a glyph from the face's own outline", () => {
  it("draws the shape the face states, as a path", () => {
    const html = markup(withOutlines());

    expect(html).toContain('data-glyph="3436"');
    expect(html).toContain("<path");
    expect(html).not.toContain("data-undrawn-glyphs");
  });

  // The face counts up from the baseline and a page counts down from its top, so
  // the outline is turned over about the glyph's own origin: the glyph starts at
  // 100pt across and 199.92 down, and its first point is the origin itself.
  it("puts the outline where the glyph stands, the right way up", () => {
    const html = markup(withOutlines());

    expect(html).toContain("M 0 14 L 10 14");
    expect(html).toContain("Q 20 9 10 4");
  });

  it("fills it in the colour the run states", () => {
    expect(markup(withOutlines())).toContain('fill="#112233"');
  });

  // A run whose glyphs are not all readable draws the ones that are and names the
  // ones it could not, rather than drawing none or pretending it drew all.
  it("draws what it can of a run and names what it could not", () => {
    const layout = layoutWith([]);
    const mixed = {
      ...OUTLINED,
      glyphs: [
        ...OUTLINED.glyphs,
        { glyph: 12, leftPt: 140, baselinePt: 200, advancePt: 9, standsFor: null },
      ],
    };
    const page = { ...firstPage(layout), glyphRuns: [mixed] };
    const html = markup({ ...layout, pages: [page] });

    expect(html).toContain('data-glyph="3436"');
    expect(html).toContain('data-undrawn-glyphs="12"');
  });

  // Nothing else about the run moves it: the outline is measured from the glyph's
  // own origin, which the baseline and the left already place.
  it("draws nothing at all for a glyph that reaches it without one", () => {
    const html = markup(withGlyphs());

    expect(html).not.toContain("<path");
    expect(html).toContain('data-undrawn-glyphs="3436"');
  });
});

// **A run stating no colour is drawn black, which is what Word draws.** The page
// used to inherit whatever colour it stood in, so a themed container drew text
// Word would have drawn black; `drawables.ts` resolves it now and both backends
// draw the same thing.
describe("Page drawing text that states no colour", () => {
  it("draws it in the black Word draws it in rather than inheriting the page's", () => {
    const html = markup(layoutWith([], [paragraphOf("Hello")]));

    expect(html).toContain('fill="#000000"');
  });
});

// A glyph named by number has no character of its own, so a page drawing one holds
// no text to select or search unless it says what the glyph stands for. The pdf
// writer maps each to its character; this is the same answer in the other
// notation.
describe("Page saying what a glyph it drew stands for", () => {
  it("writes the character beside the shape, painted nowhere", () => {
    const html = markup(withOutlines());

    expect(html).toContain('fill="none"');
    expect(html).toContain(">(</tspan>");
  });

  it("says so even where it could not draw the glyph at all", () => {
    const html = markup(withGlyphs());

    expect(html).toContain(">(</tspan>");
    expect(html).not.toContain("<path");
  });

  it("says nothing for a glyph that stands for nothing", () => {
    const layout = layoutWith([]);
    const nameless = {
      ...OUTLINED,
      glyphs: [{ ...(OUTLINED.glyphs[0] ?? {}), standsFor: null }],
    };
    const page = { ...firstPage(layout), glyphRuns: [nameless] };

    expect(markup({ ...layout, pages: [page] })).not.toContain("<tspan");
  });
});

// A fraction as `drawables.ts` hands it over: two halves at their own size and the
// bar between them, each already on the device grid. It hangs off the page exactly
// as a glyph run does, and for the same reason.
const FRACTION = [
  { kind: "text", text: "gralm", sizePt: 7.92, widthPt: 20.5, leftPt: 290.5, baselinePt: 100.08 },
  { kind: "fill", leftPt: 290.4, topPt: 104.88, widthPt: 20.96, heightPt: 0.72 },
  { kind: "text", text: "presk", sizePt: 7.92, widthPt: 19.4, leftPt: 291.85, baselinePt: 115.92 },
];

const withEquation = (): LaidOutDocument => {
  const layout = layoutWith([]);
  const page = { ...firstPage(layout), equations: [{ mark: MARK, primitives: FRACTION }] };
  return { ...layout, pages: [page] };
};

describe("Page drawing a set equation", () => {
  // A piece of an equation is a string at a place at a size, which is the shape a
  // list's number already had, so the viewer draws it the way it draws one.
  it("draws each half at its own size where the drawing placed it", () => {
    const html = markup(withEquation());

    expect(html).toContain('x="290.5" y="100.08"');
    expect(html).toContain(">gralm<");
    expect(html).toContain('x="291.85" y="115.92"');
    expect(html).toContain(">presk<");
    expect(html).toContain('font-size="7.92"');
  });

  // The bar is a fill like any other, so it is the rectangle the drawing states and
  // nothing here works anything out about it.
  it("draws the bar as the rectangle it was handed", () => {
    const html = markup(withEquation());

    expect(html).toContain('x="290.4" y="104.88" width="20.96"');
  });

  // Held to the width it was set at, as every other run is: the halves were
  // measured off the face's own advances and the bar's width is made of the same,
  // so a browser drawing them in another face keeps the two together.
  it("holds each half to the width it was set at", () => {
    expect(markup(withEquation())).toContain('textLength="20.5"');
  });
});
