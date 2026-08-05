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
  color: null,
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
  text: boxes.length === 0 ? null : { boxes, cells: [], contentHeightPt: 0, contentWidthPt: 0 },
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
      },
      leftPt: options.leftPt ?? 120,
      topPt: 30,
      heightPt: 14,
      seatPt: 0,
      baselinePt: options.baselinePt ?? 41,
      startsPage: false,
    },
  ],
  marker: null,
  markTopPt: 30,
  contentBottomPt: 44,
  widowControl: true,
  startsPage: false,
  endsPage: false,
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
    content: { kind: "shape", paint: NO_PAINT } as const,
    horizontal: { kind: "offset", from: "column", offsetEmu: 0 } as const,
    vertical: { kind: "offset", from: "paragraph", offsetEmu: 0 } as const,
    wrap: "none" as const,
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
});

const layoutWith = (
  floats: readonly ReturnType<typeof float>[],
  body: readonly ParagraphBox[] = [],
  footerFloats: readonly ReturnType<typeof float>[] = [],
  cells: readonly PlacedCell[] = [],
): LaidOutDocument => ({
  kind: "laid-out",
  page: LETTER,
  headerTopPt: 21.6,
  headerHeightPt: 0,
  bodyTopPt: 36,
  bodyBottomPt: 792,
  footerTopPt: 784.8,
  header: [],
  footer: [],
  headerCells: [],
  footerCells: [],
  headerFloats: [],
  footerFloats,
  headerInlines: [],
  footerInlines: [],
  pages: [{ index: 0, body, cells, floats, inlines: [] }],
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
    expect(html).toContain('y="41"');
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

    expect(html).toContain('y="37"');
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
    expect(html).toContain('x="102" y="41"');
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
