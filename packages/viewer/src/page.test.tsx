import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import type {
  CropInsets,
  LaidOutDocument,
  ParagraphBox,
  ParagraphMark,
  PlacedContent,
  SectionGeometry,
  TextBoxBody,
} from "@onepager/core";

import { OnePagerPage } from "./page.js";

const MARK: ParagraphMark = {
  font: { kind: "named", name: "Meridian Sans" },
  fontSizePt: 12,
  bold: false,
  italic: false,
  color: null,
};

const EMPTY_BODY: TextBoxBody = {
  blocks: [],
  insets: { leftEmu: 0, topEmu: 0, rightEmu: 0, bottomEmu: 0 },
  anchor: "top",
  wraps: true,
};

const textBox = (boxes: readonly ParagraphBox[] = []): PlacedContent => ({
  kind: "text-box",
  body: EMPTY_BODY,
  text: boxes.length === 0 ? null : { boxes, contentHeightPt: 0 },
});

const paragraphOf = (
  text: string,
  mark: ParagraphMark = MARK,
  options: { leftPt?: number; baselinePt?: number; widthPt?: number } = {},
): ParagraphBox => ({
  index: 0,
  topPt: 0,
  heightPt: 14,
  lines: [
    {
      line: {
        segments: [{ kind: "text", mark, text, widthPt: options.widthPt ?? 40 }],
        widthPt: options.widthPt ?? 40,
        heightPt: 14,
        ascentPt: 11,
      },
      leftPt: options.leftPt ?? 120,
      topPt: 30,
      baselinePt: options.baselinePt ?? 41,
    },
  ],
  marker: null,
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
    content: { kind: "shape" } as const,
    horizontal: { kind: "offset", from: "column", offsetEmu: 0 } as const,
    vertical: { kind: "offset", from: "paragraph", offsetEmu: 0 } as const,
    wrap: "none" as const,
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
): LaidOutDocument => ({
  kind: "laid-out",
  page: LETTER,
  headerTopPt: 21.6,
  headerHeightPt: 0,
  bodyTopPt: 36,
  bodyBottomPt: 792,
  footerTopPt: 784.8,
  header: [],
  body,
  footer: [],
  headerFloats: [],
  bodyFloats: floats,
  footerFloats: [],
  headerInlines: [],
  bodyInlines: [],
  footerInlines: [],
});

const markup = (
  layout: LaidOutDocument,
  options: Parameters<typeof OnePagerPage>[0] | null = null,
) =>
  renderToStaticMarkup(
    <OnePagerPage
      layout={layout}
      imageUrl={(part) =>
        part === "word/media/image1.png" ? "data:image/png;base64,AA" : undefined
      }
      {...(options === null ? {} : options)}
    />,
  );

describe("OnePagerPage", () => {
  it("sizes the page in points so everything below stays in Word's coordinates", () => {
    const html = markup(layoutWith([]));
    expect(html).toContain("width:612pt");
    expect(html).toContain("height:792pt");
  });

  it("scales the whole page rather than each object in it", () => {
    const html = renderToStaticMarkup(
      <OnePagerPage layout={layoutWith([])} imageUrl={() => undefined} scale={0.5} />,
    );
    expect(html).toContain("transform:scale(0.5)");
    expect(html).toContain("transform-origin:top left");
  });

  it("puts a picture at the point position layout computed", () => {
    const html = markup(
      layoutWith([float({ kind: "picture", part: "word/media/image1.png", crop: NO_CROP })]),
    );
    expect(html).toContain("left:100pt");
    expect(html).toContain("top:200pt");
    expect(html).toContain('src="data:image/png;base64,AA"');
  });

  it("draws the whole bitmap behind a window when srcRect crops it", () => {
    const crop: CropInsets = { left: 0, top: 0.07272, right: 0.293, bottom: 0 };
    const html = markup(
      layoutWith([float({ kind: "picture", part: "word/media/image1.png", crop })]),
    );
    // The window stays at the placed size; the bitmap behind it is 180 / (1 - 0.293)
    // wide and rides 7.272% of its own height above the top edge.
    expect(html).toContain("width:180pt;height:90pt;overflow:hidden");
    expect(html).toMatch(/width:254\.596\d*pt/);
    expect(html).toMatch(/top:-7\.058\d*pt/);
  });

  it("shows nothing for a text box holding no text", () => {
    expect(markup(layoutWith([float(textBox())]))).not.toContain("<svg");
  });

  it("outlines frames on request so placement can be checked without content", () => {
    const html = renderToStaticMarkup(
      <OnePagerPage
        layout={layoutWith([float(textBox())])}
        imageUrl={() => undefined}
        frames="outlined"
      />,
    );
    expect(html).toContain('data-kind="text-box"');
    expect(html).toContain("left:100pt");
  });

  it("marks a picture whose part never resolved rather than drawing a broken image", () => {
    const html = renderToStaticMarkup(
      <OnePagerPage
        layout={layoutWith([float({ kind: "missing-picture", relationshipId: "rId7" })])}
        imageUrl={() => undefined}
        frames="outlined"
      />,
    );
    expect(html).toContain('data-kind="missing-picture"');
  });

  it("stacks objects behind the text below the ones in front of it", () => {
    const html = renderToStaticMarkup(
      <OnePagerPage
        layout={layoutWith([
          float({ kind: "shape" }, { height: 5 }),
          float(textBox(), { behindDoc: true, height: 9 }),
        ])}
        imageUrl={() => undefined}
        frames="outlined"
      />,
    );
    expect(html.indexOf('data-kind="text-box"')).toBeLessThan(html.indexOf('data-kind="shape"'));
  });
});

describe("OnePagerPage drawing text", () => {
  it("draws each line at the baseline layout gave it", () => {
    const html = markup(layoutWith([], [paragraphOf("Hello")]));

    expect(html).toContain('data-kind="text"');
    expect(html).toContain('x="120"');
    expect(html).toContain('y="41"');
    expect(html).toContain("Hello");
  });

  it("names the authored face first and lets the page fall back behind it", () => {
    const html = markup(layoutWith([], [paragraphOf("Hello")]), {
      layout: layoutWith([], [paragraphOf("Hello")]),
      imageUrl: () => undefined,
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
    const html = renderToStaticMarkup(
      <OnePagerPage
        layout={layoutWith([float(textBox([paragraphOf("Inside")]))])}
        imageUrl={() => undefined}
        frames="outlined"
      />,
    );
    expect(html.indexOf('data-kind="text-box"')).toBeLessThan(html.indexOf("Inside"));
  });

  it("draws no text layer for a document that has none", () => {
    expect(markup(layoutWith([]))).not.toContain('data-kind="text"');
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

describe("OnePagerPage drawing a list number", () => {
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
