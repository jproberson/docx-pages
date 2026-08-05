import { readAnchors, type FloatingAnchor } from "../docx/anchors.js";
import { readInlines } from "../docx/inlines.js";
import { blockParagraphs, readBlocks, type Block } from "../docx/blocks.js";
import { defaultFooterPart, defaultHeaderPart, readRelationships } from "../docx/relationships.js";
import { MAIN_DOCUMENT_PART, type DocxPackage } from "../docx/package.js";
import { readSectionGeometry, type SectionGeometry } from "../docx/section.js";
import { readStyleTable, type StyleTable } from "../docx/styles.js";
import { readTheme, type Theme } from "../docx/theme.js";
import {
  measureStack,
  shiftBoxes,
  type BandResolver,
  type LayoutBlocker,
  type MetricsResolver,
  type ParagraphBox,
} from "./stack.js";
import { breakStack, type PageStack } from "./pages.js";
import { placeFloat, type FloatSize, type PartResolver, type PlacedFloat } from "./floats.js";
import { placeInlines, type PlacedInline } from "./inlines.js";
import { layOutTextBox } from "./text-boxes.js";
import { emuToPoints, twipsToPoints } from "./units.js";
import type { OutlinePoint, WrapBand } from "./wrapping.js";

// The header and the footer are drawn again on every page, so only the body is
// broken up: a page is the run of it that fitted between the two.
export type LaidOutPage = {
  readonly index: number;
  readonly body: readonly ParagraphBox[];
  readonly floats: readonly PlacedFloat[];
  readonly inlines: readonly PlacedInline[];
};

export type LaidOutDocument = {
  readonly kind: "laid-out";
  readonly page: SectionGeometry;
  readonly headerTopPt: number;
  readonly headerHeightPt: number;
  readonly bodyTopPt: number;
  readonly bodyBottomPt: number;
  readonly footerTopPt: number;
  readonly header: readonly ParagraphBox[];
  readonly footer: readonly ParagraphBox[];
  readonly headerFloats: readonly PlacedFloat[];
  readonly footerFloats: readonly PlacedFloat[];
  readonly headerInlines: readonly PlacedInline[];
  readonly footerInlines: readonly PlacedInline[];
  readonly pages: readonly LaidOutPage[];
};

export type DocumentLayout =
  LaidOutDocument | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

type FloatsInPart = {
  readonly floats: readonly PlacedFloat[];
  readonly part: string;
};

type FilledFloats =
  | { readonly kind: "filled"; readonly floats: readonly (readonly PlacedFloat[])[] }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

// A text box is placed as a frame first and only then holds text, since its own
// content is laid out against the rectangle the anchor resolved to.
function fillTextBoxes(
  parts: readonly FloatsInPart[],
  styles: StyleTable,
  metricsFor: MetricsResolver,
): FilledFloats {
  const filled: (readonly PlacedFloat[])[] = [];

  for (const { floats, part } of parts) {
    const placed: PlacedFloat[] = [];
    for (const float of floats) {
      if (float.content.kind !== "text-box") {
        placed.push(float);
        continue;
      }

      const laid = layOutTextBox({
        body: float.content.body,
        rect: float,
        styles,
        metricsFor,
        part,
      });
      if (laid.kind === "blocked") return { kind: "blocked", blocker: laid.blocker };
      placed.push({ ...float, content: { ...float.content, text: laid.text } });
    }
    filled.push(placed);
  }

  return { kind: "filled", floats: filled };
}

type FloatFrame = {
  readonly page: SectionGeometry;
  readonly styles: StyleTable;
  readonly metricsFor: MetricsResolver;
  readonly theme: Theme;
  readonly part: string;
  // Where the story's own text column starts, which is what a column-relative
  // offset is measured from.
  readonly columnTopPt: number;
  readonly marginTopPt: number;
};

// "Resize shape to fit text": the box is as tall as its text, and as wide as it
// too when the text does not wrap inside it. A box holding no text still fits
// itself to the paragraph mark standing in it, which Word makes one pilcrow wide
// and one line tall: measured against Word, a 270 x 0.05pt box holding one empty
// 10pt paragraph came back 5.9 x 22.2pt, and the band it wraps text out of moved
// with it.
function fittedSizePt(anchor: FloatingAnchor, frame: FloatFrame): FloatSize {
  const stored = {
    widthPt: emuToPoints(anchor.widthEmu),
    heightPt: emuToPoints(anchor.heightEmu),
  };
  const { content } = anchor;
  if (content.kind !== "text-box" || !content.body.fitsText) return stored;

  const laid = layOutTextBox({
    body: content.body,
    rect: { leftPt: 0, topPt: 0, ...stored },
    styles: frame.styles,
    metricsFor: frame.metricsFor,
    part: frame.part,
  });
  if (laid.kind === "blocked") return stored;

  // The box is as tall as its text, its insets, and the outline that runs round the
  // whole of it, and as wide as the same three when the text does not wrap. Only a
  // width the file states counts: an outline that states none is still drawn, as a
  // hairline, but the box does not grow for it. Measured against Word over outlines
  // of nothing, three quarters of a point, one and a half, three and six, in both
  // wrap modes: every stated width grew the box by the whole of itself on each axis
  // being fitted, and the unstated one grew it on neither.
  const { insets, wraps } = content.body;
  const outline = content.paint.outline;
  const outlinePt = outline === null || !outline.widthStated ? 0 : outline.widthPt;
  return {
    widthPt: wraps
      ? stored.widthPt
      : laid.text.contentWidthPt +
        emuToPoints(insets.leftEmu) +
        emuToPoints(insets.rightEmu) +
        outlinePt,
    heightPt:
      laid.text.contentHeightPt +
      emuToPoints(insets.topEmu) +
      emuToPoints(insets.bottomEmu) +
      outlinePt,
  };
}

const placeFloatIn = (
  anchor: FloatingAnchor,
  paragraphTopPt: number,
  frame: FloatFrame,
  resolvePart: PartResolver,
): PlacedFloat =>
  placeFloat({
    anchor,
    page: frame.page,
    paragraphTopPt,
    bodyTopPt: frame.columnTopPt,
    marginTopPt: frame.marginTopPt,
    resolvePart,
    theme: frame.theme,
    sizePt: fittedSizePt(anchor, frame),
  });

// Text stays off the part of an object its wrap covers by the distances its
// anchor asks for; an object wrapped top and bottom takes the whole width of the
// page with it.
function bandFor(float: PlacedFloat, frame: FloatFrame): WrapBand {
  const { area, distances, wrap } = float.anchor;
  const spansPage = wrap === "topAndBottom";
  const outline = spansPage ? undefined : outlineOf(float);
  return {
    leftPt: spansPage
      ? 0
      : float.leftPt + float.widthPt * area.left - emuToPoints(distances.leftEmu),
    rightPt: spansPage
      ? twipsToPoints(frame.page.widthTwips)
      : float.leftPt + float.widthPt * area.right + emuToPoints(distances.rightEmu),
    topPt: float.topPt + float.heightPt * area.top - emuToPoints(distances.topEmu),
    bottomPt: float.topPt + float.heightPt * area.bottom + emuToPoints(distances.bottomEmu),
    ...(outline === undefined ? {} : { outline }),
    ...(wrap === "tight" || wrap === "through" ? { outlined: true } : {}),
  };
}

// A polygon that is its own rectangle says nothing the band does not already, so
// only a shape narrower than that somewhere is carried into the layout.
function outlineOf(float: PlacedFloat): readonly OutlinePoint[] | undefined {
  const { corners } = float.anchor.area;
  const rectangular = corners.every(
    (corner) =>
      (corner.x === float.anchor.area.left || corner.x === float.anchor.area.right) &&
      (corner.y === float.anchor.area.top || corner.y === float.anchor.area.bottom),
  );
  if (rectangular) return undefined;

  return corners.map((corner) => ({
    xPt: float.leftPt + float.widthPt * corner.x,
    yPt: float.topPt + float.heightPt * corner.y,
  }));
}

// Only geometry matters here, so a picture's part is left unresolved.
const bandsIn =
  (frame: FloatFrame): BandResolver =>
  (paragraph, topPt) =>
    readAnchors(paragraph)
      .filter((anchor) => anchor.wrap !== "none")
      .map((anchor) =>
        bandFor(
          placeFloatIn(anchor, topPt, frame, () => null),
          frame,
        ),
      );

type Story = {
  readonly kind: "measured";
  readonly part: string | null;
  readonly blocks: readonly Block[];
  readonly boxes: readonly ParagraphBox[];
  readonly heightPt: number;
};

type StoryMeasurement = Story | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

type StoryFrame = {
  readonly styles: StyleTable;
  readonly metricsFor: MetricsResolver;
  readonly leftPt: number;
  readonly widthPt: number;
};

const EMPTY_STORY: Story = { kind: "measured", part: null, blocks: [], boxes: [], heightPt: 0 };

function measureStory(
  pkg: DocxPackage,
  part: string | null,
  frame: StoryFrame,
  originPt: number,
  bandsFor: BandResolver,
): StoryMeasurement {
  if (part === null) return EMPTY_STORY;

  const blocks = readBlocks(pkg, part);
  const measured = measureStack({ ...frame, blocks, part, originPt, bandsFor });
  if (measured.kind === "blocked") return measured;
  return { kind: "measured", part, blocks, boxes: measured.boxes, heightPt: measured.heightPt };
}

export function layOutDocument(pkg: DocxPackage, metricsFor: MetricsResolver): DocumentLayout {
  const page = readSectionGeometry(pkg);
  const styles = readStyleTable(pkg);
  const theme = readTheme(pkg);
  const headerTopPt = twipsToPoints(page.margin.headerTwips);
  const pageHeightPt = twipsToPoints(page.heightTwips);
  const leftPt = twipsToPoints(page.margin.leftTwips);
  const widthPt = twipsToPoints(page.widthTwips - page.margin.leftTwips - page.margin.rightTwips);
  const frame: StoryFrame = { styles, metricsFor, leftPt, widthPt };

  const floatFrame = (
    part: string | null,
    columnTopPt: number,
    marginTopPt: number,
  ): FloatFrame => ({
    page,
    styles,
    metricsFor,
    theme,
    part: part ?? MAIN_DOCUMENT_PART,
    columnTopPt,
    marginTopPt,
  });

  // A header's own objects are placed against the top of the body, which is not
  // known until the header has been measured. So it is measured once against the
  // margin the page asks for, and again against the body top that came of it.
  const headerPart = defaultHeaderPart(pkg);
  const measureHeader = (marginTopPt: number): StoryMeasurement =>
    measureStory(
      pkg,
      headerPart,
      frame,
      headerTopPt,
      bandsIn(floatFrame(headerPart, headerTopPt, marginTopPt)),
    );
  const bodyTopUnder = (story: StoryMeasurement): number =>
    Math.max(
      twipsToPoints(page.margin.topTwips),
      headerTopPt + (story.kind === "blocked" ? 0 : story.heightPt),
    );

  const firstPass = measureHeader(twipsToPoints(page.margin.topTwips));
  if (firstPass.kind === "blocked") return firstPass;
  const bodyTopPt = bodyTopUnder(firstPass);
  const header = measureHeader(bodyTopPt);
  if (header.kind === "blocked") return header;

  const headerHeightPt = header.heightPt;

  // The footer hangs from the bottom edge, so it is measured at the origin and then
  // dropped to the height it turned out to need.
  const footerPart = defaultFooterPart(pkg);
  const measuredFooter = measureStory(
    pkg,
    footerPart,
    frame,
    0,
    bandsIn(floatFrame(footerPart, 0, bodyTopPt)),
  );
  if (measuredFooter.kind === "blocked") return measuredFooter;
  const footerTopPt =
    pageHeightPt - twipsToPoints(page.margin.footerTwips) - measuredFooter.heightPt;
  const footer: Story = {
    ...measuredFooter,
    boxes: shiftBoxes(measuredFooter.boxes, footerTopPt),
  };

  const marginBottomPt = pageHeightPt - twipsToPoints(page.margin.bottomTwips);
  const bodyBottomPt =
    footer.part === null ? marginBottomPt : Math.min(marginBottomPt, footerTopPt);

  const bodyBlocks = readBlocks(pkg);
  const bodyStack = measureStack({
    ...frame,
    blocks: bodyBlocks,
    part: MAIN_DOCUMENT_PART,
    originPt: bodyTopPt,
    bandsFor: bandsIn(floatFrame(MAIN_DOCUMENT_PART, bodyTopPt, bodyTopPt)),
  });

  if (bodyStack.kind === "blocked") return { kind: "blocked", blocker: bodyStack.blocker };

  // A drawing belongs to the page its anchoring paragraph landed on, and is placed
  // against the top that paragraph has there.
  const drawingsFor = (
    blocks: readonly Block[],
    boxOf: ReadonlyMap<number, ParagraphBox>,
    floats: FloatFrame,
  ): { readonly floats: readonly PlacedFloat[]; readonly inlines: readonly PlacedInline[] } => {
    const part = floats.part;
    const relationships = readRelationships(pkg, part);
    const resolvePart = (relationshipId: string): string | null => {
      const target = relationships.get(relationshipId)?.part;
      return target !== undefined && pkg.parts.has(target) ? target : null;
    };

    const anchored = blockParagraphs(blocks).flatMap((paragraph) => {
      const box = boxOf.get(paragraph.index);
      return box === undefined ? [] : [{ paragraph, box }];
    });

    return {
      floats: anchored.flatMap(({ paragraph, box }) =>
        readAnchors(paragraph).map((anchor) =>
          placeFloatIn(anchor, box.topPt, floats, resolvePart),
        ),
      ),
      inlines: anchored.flatMap(({ paragraph, box }) =>
        placeInlines({
          drawings: readInlines(paragraph),
          box,
          resolvePart,
          theme: floats.theme,
        }),
      ),
    };
  };

  const drawingsIn = (story: Story, columnTopPt: number) =>
    story.part === null
      ? { floats: [], inlines: [] }
      : drawingsFor(
          story.blocks,
          boxesOf(story.boxes),
          floatFrame(story.part, columnTopPt, bodyTopPt),
        );

  const headerDrawings = drawingsIn(header, headerTopPt);
  const footerDrawings = drawingsIn(footer, footerTopPt);
  const broken = breakStack({ boxes: bodyStack.boxes, topPt: bodyTopPt, bottomPt: bodyBottomPt });
  const bodyDrawings = pageBoxes(broken).map((boxOf) =>
    drawingsFor(bodyBlocks, boxOf, floatFrame(MAIN_DOCUMENT_PART, bodyTopPt, bodyTopPt)),
  );

  const filled = fillTextBoxes(
    [
      { floats: headerDrawings.floats, part: header.part ?? MAIN_DOCUMENT_PART },
      { floats: footerDrawings.floats, part: footer.part ?? MAIN_DOCUMENT_PART },
      ...bodyDrawings.map((drawings) => ({ floats: drawings.floats, part: MAIN_DOCUMENT_PART })),
    ],
    styles,
    metricsFor,
  );
  if (filled.kind === "blocked") return filled;
  const [headerFloats = [], footerFloats = [], ...pageFloats] = filled.floats;

  return {
    kind: "laid-out",
    page,
    headerFloats,
    footerFloats,
    headerInlines: headerDrawings.inlines,
    footerInlines: footerDrawings.inlines,
    headerTopPt,
    headerHeightPt,
    bodyTopPt,
    bodyBottomPt,
    footerTopPt,
    header: header.boxes,
    footer: footer.boxes,
    pages: broken.map((each) => ({
      index: each.index,
      body: each.boxes,
      floats: pageFloats[each.index] ?? [],
      inlines: bodyDrawings[each.index]?.inlines ?? [],
    })),
  };
}

const boxesOf = (boxes: readonly ParagraphBox[]): ReadonlyMap<number, ParagraphBox> =>
  new Map(boxes.map((box) => [box.index, box]));

// A paragraph the break ran through is on two pages; its drawings belong to the
// first of them, where the paragraph starts.
function pageBoxes(pages: readonly PageStack[]): readonly ReadonlyMap<number, ParagraphBox>[] {
  const seen = new Set<number>();
  return pages.map((page) => {
    const boxes = new Map<number, ParagraphBox>();
    for (const box of page.boxes) {
      if (seen.has(box.index)) continue;
      seen.add(box.index);
      boxes.set(box.index, box);
    }
    return boxes;
  });
}
