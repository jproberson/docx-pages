export { OnePagerError, isOnePagerError } from "./errors.js";
export type { ContextValue, ErrorContext, OnePagerErrorInit } from "./errors.js";

export { openDocx, partText, partXml, MAIN_DOCUMENT_PART } from "./docx/package.js";
export type { DocxPackage } from "./docx/package.js";

export { readSectionGeometry, W_NS } from "./docx/section.js";
export type { PageMargin, SectionGeometry } from "./docx/section.js";

export {
  advanceWidthPt,
  ascentPt,
  lineHeightPt,
  lookupFontMetrics,
  NO_ADVANCES,
} from "./layout/font-metrics.js";
export type {
  AdvanceTable,
  AdvancesUnavailable,
  FaceRequest,
  FontMetrics,
  GlyphAdvances,
  MetricsLookup,
  SuppliedFace,
} from "./layout/font-metrics.js";

export { readFontFile, readFontMetrics } from "./layout/font-file.js";
export type {
  FontFileFormat,
  ReadFontFileResult,
  ReadFontMetricsResult,
} from "./layout/font-file.js";

export { readBlocks, readParagraphs, blocksIn, blockParagraphs, MC_NS } from "./docx/blocks.js";
export type { Block, CellVerticalAlign, Paragraph, TableCell, TableRow } from "./docx/blocks.js";

export { paragraphText, paragraphDescendants } from "./docx/paragraphs.js";

export { numberParagraphs } from "./docx/list-numbers.js";
export type { ParagraphNumber, ParagraphNumbers } from "./docx/list-numbers.js";

export { readNumberingTable, numberingLevel, NUMBERING_PART } from "./docx/numbering.js";
export type {
  LevelRestart,
  NumberFormat,
  NumberingLevel,
  NumberingTable,
  NumberSuffix,
} from "./docx/numbering.js";

export { readRuns } from "./docx/runs.js";
export type { RunPiece, TextRun } from "./docx/runs.js";

export {
  readStyleTable,
  resolveParagraphMark,
  resolveRuns,
  A_NS,
  WORD_DEFAULT_FONT_SIZE_PT,
} from "./docx/styles.js";
export type { FontChoice, LineRule, MarkedRun, ParagraphMark, StyleTable } from "./docx/styles.js";

export {
  readRelationships,
  relationshipsPartFor,
  defaultHeaderPart,
  R_NS,
  PKG_REL_NS,
} from "./docx/relationships.js";
export type { Relationship } from "./docx/relationships.js";

export { breakLines, faceRequestFor } from "./layout/lines.js";
export type {
  BreakLinesInput,
  LineBreaking,
  LineSegment,
  LineTabs,
  MeasureFailure,
  TextLine,
} from "./layout/lines.js";

export { nextTabStopPt, tabStopsPt, DEFAULT_TAB_STOP_PT } from "./layout/tab-stops.js";

export { measureStack, WP_NS } from "./layout/stack.js";
export type {
  LayoutBlocker,
  MeasureStackInput,
  MetricsResolver,
  ParagraphBox,
  ParagraphMarker,
  PlacedLine,
  StackMeasurement,
} from "./layout/stack.js";

export { layOutTextBox } from "./layout/text-boxes.js";
export type {
  LayOutTextBoxInput,
  PlacedTextBox,
  TextBoxLayout,
  TextBoxRect,
} from "./layout/text-boxes.js";

export { layOutDocument } from "./layout/document.js";
export type { DocumentLayout, LaidOutDocument } from "./layout/document.js";

export { readAnchors } from "./docx/anchors.js";
export type { AnchorOrigin, AnchorPosition, FloatingAnchor, WrapMode } from "./docx/anchors.js";

export { placeFloat } from "./layout/floats.js";
export type { PartResolver, PlacedContent, PlacedFloat, PlaceFloatInput } from "./layout/floats.js";

export { readInlines } from "./docx/inlines.js";
export type { InlineDrawing } from "./docx/inlines.js";

export { placeInlines } from "./layout/inlines.js";
export type { PlacedInline, PlaceInlinesInput } from "./layout/inlines.js";

export { resolveParagraphFrame, resolveParagraphNumbering } from "./docx/styles.js";
export type { ParagraphAlignment, ParagraphFrame, ParagraphNumbering } from "./docx/styles.js";

export {
  readDrawingContent,
  DEFAULT_TEXT_INSETS,
  NO_CROP,
  PIC_NS,
  WPS_NS,
} from "./docx/drawing.js";
export type {
  CropInsets,
  DrawingContent,
  TextBoxAnchor,
  TextBoxBody,
  TextBoxInsets,
} from "./docx/drawing.js";

export { emuToPoints, twipsToPoints, EMU_PER_POINT, TWIPS_PER_POINT } from "./layout/units.js";
