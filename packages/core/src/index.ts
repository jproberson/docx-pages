export { OnePagerError, isOnePagerError } from "./errors.js";
export type { ContextValue, ErrorContext, OnePagerErrorInit } from "./errors.js";

export { openDocx, partText, partXml, MAIN_DOCUMENT_PART } from "./docx/package.js";
export type { DocxPackage } from "./docx/package.js";

export { readSectionGeometry, W_NS } from "./docx/section.js";
export type { PageMargin, SectionGeometry } from "./docx/section.js";

export { ascentPt, lineHeightPt, lookupFontMetrics } from "./layout/font-metrics.js";
export type { FontMetrics, MetricsLookup } from "./layout/font-metrics.js";

export { readFontMetrics } from "./layout/font-file.js";
export type { FontFileFormat, ReadFontMetricsResult } from "./layout/font-file.js";

export { readParagraphs, paragraphText, paragraphDescendants, MC_NS } from "./docx/paragraphs.js";
export type { Paragraph } from "./docx/paragraphs.js";

export {
  readStyleTable,
  resolveParagraphMark,
  A_NS,
  WORD_DEFAULT_FONT_SIZE_PT,
} from "./docx/styles.js";
export type { FontChoice, ParagraphMark, StyleTable } from "./docx/styles.js";

export {
  readRelationships,
  relationshipsPartFor,
  defaultHeaderPart,
  R_NS,
  PKG_REL_NS,
} from "./docx/relationships.js";
export type { Relationship } from "./docx/relationships.js";

export { measureStack, WP_NS } from "./layout/stack.js";
export type {
  LayoutBlocker,
  MeasureStackInput,
  MetricsResolver,
  ParagraphBox,
  StackMeasurement,
} from "./layout/stack.js";

export { layOutDocument } from "./layout/document.js";
export type { DocumentLayout } from "./layout/document.js";

export { readAnchors } from "./docx/anchors.js";
export type { AnchorOrigin, AnchorPosition, FloatingAnchor, WrapMode } from "./docx/anchors.js";

export { placeFloat } from "./layout/floats.js";
export type { PlacedFloat, PlaceFloatInput } from "./layout/floats.js";

export { emuToPoints, twipsToPoints, EMU_PER_POINT, TWIPS_PER_POINT } from "./layout/units.js";
