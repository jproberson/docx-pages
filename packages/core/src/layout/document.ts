import { readParagraphs } from "../docx/paragraphs.js";
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

export const TWIPS_PER_POINT = 20;

export const twipsToPoints = (twips: number): number => twips / TWIPS_PER_POINT;

export type DocumentLayout =
  | {
      readonly kind: "laid-out";
      readonly page: SectionGeometry;
      readonly headerTopPt: number;
      readonly headerHeightPt: number;
      readonly bodyTopPt: number;
      readonly header: readonly ParagraphBox[];
      readonly body: readonly ParagraphBox[];
    }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

export function layOutDocument(pkg: DocxPackage, metricsFor: MetricsResolver): DocumentLayout {
  const page = readSectionGeometry(pkg);
  const styles = readStyleTable(pkg);
  const headerTopPt = twipsToPoints(page.margin.headerTwips);

  const headerPart = defaultHeaderPart(pkg);
  const headerStack =
    headerPart === null
      ? null
      : measureStack({
          paragraphs: readParagraphs(pkg, headerPart),
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

  const bodyStack = measureStack({
    paragraphs: readParagraphs(pkg),
    styles,
    metricsFor,
    part: MAIN_DOCUMENT_PART,
    originPt: bodyTopPt,
  });

  if (bodyStack.kind === "blocked") return { kind: "blocked", blocker: bodyStack.blocker };

  return {
    kind: "laid-out",
    page,
    headerTopPt,
    headerHeightPt,
    bodyTopPt,
    header: headerStack === null ? [] : headerStack.boxes,
    body: bodyStack.boxes,
  };
}
