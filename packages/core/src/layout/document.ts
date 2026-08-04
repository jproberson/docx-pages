import { readAnchors, type FloatingAnchor } from "../docx/anchors.js";
import { readInlines } from "../docx/inlines.js";
import { blockParagraphs, readBlocks, type Block } from "../docx/blocks.js";
import { defaultFooterPart, defaultHeaderPart, readRelationships } from "../docx/relationships.js";
import { MAIN_DOCUMENT_PART, type DocxPackage } from "../docx/package.js";
import { readSectionGeometry, type SectionGeometry } from "../docx/section.js";
import { readStyleTable, resolveParagraphFrame, type StyleTable } from "../docx/styles.js";
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
import type { WrapBand } from "./wrapping.js";

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
  readonly part: string;
  // Where the story's own text column starts, which is what a column-relative
  // offset is measured from.
  readonly columnTopPt: number;
};

// "Resize shape to fit text": the box is as tall as its text, and as wide as it
// too when the text does not wrap inside it. A box holding no text has nothing to
// fit itself to, so it keeps the size the file stored for it.
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
  if (laid.kind === "blocked" || laid.text.boxes.every((box) => box.lines.length === 0)) {
    return stored;
  }

  const { insets, wraps } = content.body;
  return {
    widthPt: wraps
      ? stored.widthPt
      : laid.text.contentWidthPt + emuToPoints(insets.leftEmu) + emuToPoints(insets.rightEmu),
    heightPt:
      laid.text.contentHeightPt + emuToPoints(insets.topEmu) + emuToPoints(insets.bottomEmu),
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
    resolvePart,
    sizePt: fittedSizePt(anchor, frame),
  });

// Text stays off an object by the distances its anchor asks for; an object wrapped
// top and bottom takes the whole width of the page with it.
function bandFor(float: PlacedFloat, frame: FloatFrame): WrapBand {
  const { distances, wrap } = float.anchor;
  const spansPage = wrap === "topAndBottom";
  return {
    leftPt: spansPage ? 0 : float.leftPt - emuToPoints(distances.leftEmu),
    rightPt: spansPage
      ? twipsToPoints(frame.page.widthTwips)
      : float.leftPt + float.widthPt + emuToPoints(distances.rightEmu),
    topPt: float.topPt - emuToPoints(distances.topEmu),
    bottomPt: float.topPt + float.heightPt + emuToPoints(distances.bottomEmu),
  };
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
  const headerTopPt = twipsToPoints(page.margin.headerTwips);
  const pageHeightPt = twipsToPoints(page.heightTwips);
  const leftPt = twipsToPoints(page.margin.leftTwips);
  const widthPt = twipsToPoints(page.widthTwips - page.margin.leftTwips - page.margin.rightTwips);
  const frame: StoryFrame = { styles, metricsFor, leftPt, widthPt };

  const floatFrame = (part: string | null, columnTopPt: number): FloatFrame => ({
    page,
    styles,
    metricsFor,
    part: part ?? MAIN_DOCUMENT_PART,
    columnTopPt,
  });

  const headerPart = defaultHeaderPart(pkg);
  const header = measureStory(
    pkg,
    headerPart,
    frame,
    headerTopPt,
    bandsIn(floatFrame(headerPart, headerTopPt)),
  );
  if (header.kind === "blocked") return header;

  const headerHeightPt = header.heightPt;
  const bodyTopPt = Math.max(twipsToPoints(page.margin.topTwips), headerTopPt + headerHeightPt);

  // The footer hangs from the bottom edge, so it is measured at the origin and then
  // dropped to the height it turned out to need.
  const footerPart = defaultFooterPart(pkg);
  const measuredFooter = measureStory(
    pkg,
    footerPart,
    frame,
    0,
    bandsIn(floatFrame(footerPart, 0)),
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
    bandsFor: bandsIn(floatFrame(MAIN_DOCUMENT_PART, bodyTopPt)),
  });

  if (bodyStack.kind === "blocked") return { kind: "blocked", blocker: bodyStack.blocker };

  // A drawing belongs to the page its anchoring paragraph landed on, and is placed
  // against the top that paragraph has there.
  const drawingsFor = (
    blocks: readonly Block[],
    topOf: ReadonlyMap<number, number>,
    floats: FloatFrame,
  ): { readonly floats: readonly PlacedFloat[]; readonly inlines: readonly PlacedInline[] } => {
    const part = floats.part;
    const relationships = readRelationships(pkg, part);
    const resolvePart = (relationshipId: string): string | null => {
      const target = relationships.get(relationshipId)?.part;
      return target !== undefined && pkg.parts.has(target) ? target : null;
    };

    const anchored = blockParagraphs(blocks).flatMap((paragraph) => {
      const paragraphTopPt = topOf.get(paragraph.index);
      return paragraphTopPt === undefined ? [] : [{ paragraph, paragraphTopPt }];
    });

    return {
      floats: anchored.flatMap(({ paragraph, paragraphTopPt }) =>
        readAnchors(paragraph).map((anchor) =>
          placeFloatIn(anchor, paragraphTopPt, floats, resolvePart),
        ),
      ),
      inlines: anchored.flatMap(({ paragraph, paragraphTopPt }) =>
        placeInlines({
          drawings: readInlines(paragraph),
          page,
          frame: resolveParagraphFrame(paragraph, styles),
          paragraphTopPt,
          resolvePart,
        }),
      ),
    };
  };

  const drawingsIn = (story: Story, columnTopPt: number) =>
    story.part === null
      ? { floats: [], inlines: [] }
      : drawingsFor(story.blocks, topsOf(story.boxes), floatFrame(story.part, columnTopPt));

  const headerDrawings = drawingsIn(header, headerTopPt);
  const footerDrawings = drawingsIn(footer, footerTopPt);
  const broken = breakStack({ boxes: bodyStack.boxes, topPt: bodyTopPt, bottomPt: bodyBottomPt });
  const bodyDrawings = pageTops(broken).map((topOf) =>
    drawingsFor(bodyBlocks, topOf, floatFrame(MAIN_DOCUMENT_PART, bodyTopPt)),
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

const topsOf = (boxes: readonly ParagraphBox[]): ReadonlyMap<number, number> =>
  new Map(boxes.map((box) => [box.index, box.topPt]));

// A paragraph the break ran through is on two pages; its drawings belong to the
// first of them, where the paragraph starts.
function pageTops(pages: readonly PageStack[]): readonly ReadonlyMap<number, number>[] {
  const seen = new Set<number>();
  return pages.map((page) => {
    const tops = new Map<number, number>();
    for (const box of page.boxes) {
      if (seen.has(box.index)) continue;
      seen.add(box.index);
      tops.set(box.index, box.topPt);
    }
    return tops;
  });
}
