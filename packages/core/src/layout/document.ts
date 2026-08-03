import { readAnchors } from "../docx/anchors.js";
import { blockParagraphs, readBlocks, type Block } from "../docx/blocks.js";
import { defaultHeaderPart } from "../docx/relationships.js";
import { MAIN_DOCUMENT_PART, type DocxPackage } from "../docx/package.js";
import { readSectionGeometry, type SectionGeometry } from "../docx/section.js";
import { readStyleTable } from "../docx/styles.js";
import {
  measureStack,
  type LayoutBlocker,
  type MetricsResolver,
  type ParagraphBox,
} from "./stack.js";
import { placeFloat, type PlacedFloat } from "./floats.js";
import { twipsToPoints } from "./units.js";

export type DocumentLayout =
  | {
      readonly kind: "laid-out";
      readonly page: SectionGeometry;
      readonly headerTopPt: number;
      readonly headerHeightPt: number;
      readonly bodyTopPt: number;
      readonly header: readonly ParagraphBox[];
      readonly body: readonly ParagraphBox[];
      readonly headerFloats: readonly PlacedFloat[];
      readonly bodyFloats: readonly PlacedFloat[];
    }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

export function layOutDocument(pkg: DocxPackage, metricsFor: MetricsResolver): DocumentLayout {
  const page = readSectionGeometry(pkg);
  const styles = readStyleTable(pkg);
  const headerTopPt = twipsToPoints(page.margin.headerTwips);

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
  });

  if (bodyStack.kind === "blocked") return { kind: "blocked", blocker: bodyStack.blocker };

  const floatsFor = (
    blocks: readonly Block[],
    boxes: readonly ParagraphBox[],
  ): readonly PlacedFloat[] => {
    const topOf = new Map(boxes.map((box) => [box.index, box.topPt]));
    return blockParagraphs(blocks).flatMap((paragraph) =>
      readAnchors(paragraph).map((anchor) =>
        placeFloat({
          anchor,
          page,
          paragraphTopPt: topOf.get(paragraph.index) ?? bodyTopPt,
          bodyTopPt,
        }),
      ),
    );
  };

  const headerBoxes = headerStack === null ? [] : headerStack.boxes;

  return {
    kind: "laid-out",
    page,
    headerFloats: floatsFor(headerBlocks, headerBoxes),
    bodyFloats: floatsFor(bodyBlocks, bodyStack.boxes),
    headerTopPt,
    headerHeightPt,
    bodyTopPt,
    header: headerBoxes,
    body: bodyStack.boxes,
  };
}
