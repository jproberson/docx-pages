import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import type { CropInsets, LaidOutDocument, PlacedContent, SectionGeometry } from "@onepager/core";

import { OnePagerPage } from "./page.js";

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

const layoutWith = (floats: readonly ReturnType<typeof float>[]): LaidOutDocument => ({
  kind: "laid-out",
  page: LETTER,
  headerTopPt: 21.6,
  headerHeightPt: 0,
  bodyTopPt: 36,
  header: [],
  body: [],
  headerFloats: [],
  bodyFloats: floats,
  headerInlines: [],
  bodyInlines: [],
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

  it("shows nothing for a text box by default, since its text is not laid out yet", () => {
    expect(markup(layoutWith([float({ kind: "text-box" })]))).not.toContain("<svg");
  });

  it("outlines frames on request so placement can be checked without content", () => {
    const html = renderToStaticMarkup(
      <OnePagerPage
        layout={layoutWith([float({ kind: "text-box" })])}
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
          float({ kind: "text-box" }, { behindDoc: true, height: 9 }),
        ])}
        imageUrl={() => undefined}
        frames="outlined"
      />,
    );
    expect(html.indexOf('data-kind="text-box"')).toBeLessThan(html.indexOf('data-kind="shape"'));
  });
});
