import { readAnchors } from "../docx/anchors.js";
import { readInlines } from "../docx/inlines.js";
import { blockParagraphs, readBlocks, type Block } from "../docx/blocks.js";
import { defaultHeaderPart, readRelationships } from "../docx/relationships.js";
import { MAIN_DOCUMENT_PART, type DocxPackage } from "../docx/package.js";
import { readSectionGeometry, type SectionGeometry } from "../docx/section.js";
import { readStyleTable, resolveParagraphFrame, type StyleTable } from "../docx/styles.js";
import {
  measureStack,
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
  readonly header: readonly ParagraphBox[];
  readonly body: readonly ParagraphBox[];
  readonly headerFloats: readonly PlacedFloat[];
  readonly bodyFloats: readonly PlacedFloat[];
  readonly headerInlines: readonly PlacedInline[];
  readonly bodyInlines: readonly PlacedInline[];
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

export function layOutDocument(pkg: DocxPackage, metricsFor: MetricsResolver): DocumentLayout {
  const page = readSectionGeometry(pkg);
  const styles = readStyleTable(pkg);
  const headerTopPt = twipsToPoints(page.margin.headerTwips);
  const leftPt = twipsToPoints(page.margin.leftTwips);
  const widthPt = twipsToPoints(page.widthTwips - page.margin.leftTwips - page.margin.rightTwips);

  const headerPart = defaultHeaderPart(pkg);
  const headerBlocks = headerPart === null ? [] : readBlocks(pkg, headerPart);
  const headerStack =
    headerPart === null
      ? null
      : measureStack({
          blocks: headerBlocks,
          styles,
          metricsFor,
          part: headerPart,
          originPt: headerTopPt,
          leftPt,
          widthPt,
        });

  if (headerStack !== null && headerStack.kind === "blocked") {
    return { kind: "blocked", blocker: headerStack.blocker };
  }

  const headerHeightPt = headerStack === null ? 0 : headerStack.heightPt;
  const bodyTopPt = Math.max(twipsToPoints(page.margin.topTwips), headerTopPt + headerHeightPt);

  const bodyBlocks = readBlocks(pkg);
  const bodyStack = measureStack({
    blocks: bodyBlocks,
    styles,
    metricsFor,
    part: MAIN_DOCUMENT_PART,
    originPt: bodyTopPt,
    leftPt,
    widthPt,
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

  const headerBoxes = headerStack === null ? [] : headerStack.boxes;
  const headerDrawings =
    headerPart === null
      ? { floats: [], inlines: [] }
      : drawingsFor(headerBlocks, headerBoxes, headerPart);
  const bodyDrawings = drawingsFor(bodyBlocks, bodyStack.boxes, MAIN_DOCUMENT_PART);

  const filled = fillTextBoxes(
    [
      { floats: headerDrawings.floats, part: headerPart ?? MAIN_DOCUMENT_PART },
      { floats: bodyDrawings.floats, part: MAIN_DOCUMENT_PART },
    ],
    styles,
    metricsFor,
  );
  if (filled.kind === "blocked") return filled;
  const [headerFloats = [], bodyFloats = []] = filled.floats;

  return {
    kind: "laid-out",
    page,
    headerFloats,
    bodyFloats,
    headerInlines: headerDrawings.inlines,
    bodyInlines: bodyDrawings.inlines,
    headerTopPt,
    headerHeightPt,
    bodyTopPt,
    header: headerBoxes,
    body: bodyStack.boxes,
  };
}
