import { readAnchors } from "../docx/anchors.js";
import { readInlines } from "../docx/inlines.js";
import { blockParagraphs, readBlocks, type Block } from "../docx/blocks.js";
import { defaultFooterPart, defaultHeaderPart, readRelationships } from "../docx/relationships.js";
import { MAIN_DOCUMENT_PART, type DocxPackage } from "../docx/package.js";
import { readSectionGeometry, type SectionGeometry } from "../docx/section.js";
import { readStyleTable, resolveParagraphFrame, type StyleTable } from "../docx/styles.js";
import {
  measureStack,
  shiftBoxes,
  type LayoutBlocker,
  type MetricsResolver,
  type ParagraphBox,
} from "./stack.js";
import { placeFloat, type PlacedFloat } from "./floats.js";
import { placeInlines, type PlacedInline } from "./inlines.js";
import { layOutTextBox } from "./text-boxes.js";
import { twipsToPoints } from "./units.js";

export type LaidOutDocument = {
  readonly kind: "laid-out";
  readonly page: SectionGeometry;
  readonly headerTopPt: number;
  readonly headerHeightPt: number;
  readonly bodyTopPt: number;
  readonly bodyBottomPt: number;
  readonly footerTopPt: number;
  readonly header: readonly ParagraphBox[];
  readonly body: readonly ParagraphBox[];
  readonly footer: readonly ParagraphBox[];
  readonly headerFloats: readonly PlacedFloat[];
  readonly bodyFloats: readonly PlacedFloat[];
  readonly footerFloats: readonly PlacedFloat[];
  readonly headerInlines: readonly PlacedInline[];
  readonly bodyInlines: readonly PlacedInline[];
  readonly footerInlines: readonly PlacedInline[];
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
): StoryMeasurement {
  if (part === null) return EMPTY_STORY;

  const blocks = readBlocks(pkg, part);
  const measured = measureStack({ ...frame, blocks, part, originPt });
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

  const header = measureStory(pkg, defaultHeaderPart(pkg), frame, headerTopPt);
  if (header.kind === "blocked") return header;

  const headerHeightPt = header.heightPt;
  const bodyTopPt = Math.max(twipsToPoints(page.margin.topTwips), headerTopPt + headerHeightPt);

  // The footer hangs from the bottom edge, so it is measured at the origin and then
  // dropped to the height it turned out to need.
  const measuredFooter = measureStory(pkg, defaultFooterPart(pkg), frame, 0);
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
  });

  if (bodyStack.kind === "blocked") return { kind: "blocked", blocker: bodyStack.blocker };

  const drawingsFor = (
    blocks: readonly Block[],
    boxes: readonly ParagraphBox[],
    part: string,
  ): { readonly floats: readonly PlacedFloat[]; readonly inlines: readonly PlacedInline[] } => {
    const topOf = new Map(boxes.map((box) => [box.index, box.topPt]));
    const relationships = readRelationships(pkg, part);
    const resolvePart = (relationshipId: string): string | null => {
      const target = relationships.get(relationshipId)?.part;
      return target !== undefined && pkg.parts.has(target) ? target : null;
    };

    const paragraphs = blockParagraphs(blocks);
    return {
      floats: paragraphs.flatMap((paragraph) =>
        readAnchors(paragraph).map((anchor) =>
          placeFloat({
            anchor,
            page,
            paragraphTopPt: topOf.get(paragraph.index) ?? bodyTopPt,
            bodyTopPt,
            resolvePart,
          }),
        ),
      ),
      inlines: paragraphs.flatMap((paragraph) =>
        placeInlines({
          drawings: readInlines(paragraph),
          page,
          frame: resolveParagraphFrame(paragraph, styles),
          paragraphTopPt: topOf.get(paragraph.index) ?? bodyTopPt,
          resolvePart,
        }),
      ),
    };
  };

  const drawingsIn = (story: Story) =>
    story.part === null
      ? { floats: [], inlines: [] }
      : drawingsFor(story.blocks, story.boxes, story.part);

  const headerDrawings = drawingsIn(header);
  const footerDrawings = drawingsIn(footer);
  const bodyDrawings = drawingsFor(bodyBlocks, bodyStack.boxes, MAIN_DOCUMENT_PART);

  const filled = fillTextBoxes(
    [
      { floats: headerDrawings.floats, part: header.part ?? MAIN_DOCUMENT_PART },
      { floats: bodyDrawings.floats, part: MAIN_DOCUMENT_PART },
      { floats: footerDrawings.floats, part: footer.part ?? MAIN_DOCUMENT_PART },
    ],
    styles,
    metricsFor,
  );
  if (filled.kind === "blocked") return filled;
  const [headerFloats = [], bodyFloats = [], footerFloats = []] = filled.floats;

  return {
    kind: "laid-out",
    page,
    headerFloats,
    bodyFloats,
    footerFloats,
    headerInlines: headerDrawings.inlines,
    bodyInlines: bodyDrawings.inlines,
    footerInlines: footerDrawings.inlines,
    headerTopPt,
    headerHeightPt,
    bodyTopPt,
    bodyBottomPt,
    footerTopPt,
    header: header.boxes,
    body: bodyStack.boxes,
    footer: footer.boxes,
  };
}
