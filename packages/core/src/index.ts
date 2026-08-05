export { DocxPagesError, isDocxPagesError } from "./errors.js";
export type { ContextValue, ErrorContext, DocxPagesErrorInit } from "./errors.js";

export { openDocx, partText, partXml, MAIN_DOCUMENT_PART } from "./docx/package.js";
export type { DocxPackage } from "./docx/package.js";

export { readSectionGeometry, W_NS } from "./docx/section.js";
export type { PageMargin, SectionGeometry } from "./docx/section.js";
export { readUnhonoured, withSubstitutedFaces } from "./docx/fidelity.js";
export type {
  Unhonoured,
  UnhonouredEffect,
  UnhonouredKind,
  UnhonouredPlace,
} from "./docx/fidelity.js";
export {
  drawablePicture,
  pictureExtension,
  METAFILE_EXTENSION,
  PICTURE_MEDIA_TYPES,
} from "./docx/pictures.js";
export {
  readDocumentSettings,
  roundsAnchorsToTwips,
  DEFAULT_SETTINGS,
  SETTINGS_PART,
} from "./docx/settings.js";
export type { DocumentSettings } from "./docx/settings.js";

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

export {
  readBlocks,
  readParagraphs,
  blocksIn,
  blockParagraphs,
  DEFAULT_TABLE_INSETS,
  MC_NS,
} from "./docx/blocks.js";
export type {
  Block,
  CellVerticalAlign,
  Paragraph,
  TableCell,
  TableInsets,
  TableRow,
} from "./docx/blocks.js";

export { paragraphText, paragraphDescendants } from "./docx/paragraphs.js";

export { facesUsed } from "./docx/faces.js";
export type { UsedFace } from "./docx/faces.js";

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
  defaultFooterPart,
  defaultHeaderPart,
  R_NS,
  PKG_REL_NS,
} from "./docx/relationships.js";
export type { Relationship } from "./docx/relationships.js";

export { breakLines, faceRequestFor, measureText } from "./layout/lines.js";
export type {
  BreakLinesInput,
  LineBreaking,
  TextMeasurement,
  LineSegment,
  LineTabs,
  MeasureFailure,
  TextLine,
} from "./layout/lines.js";

export {
  nextTabStop,
  tabStopsPt,
  DEFAULT_TAB_STOP_PT,
  type TabStopPt,
} from "./layout/tab-stops.js";

export { measureStack, shiftBox, shiftBoxes, shiftCells, WP_NS } from "./layout/stack.js";
export type {
  BandResolver,
  ClipRect,
  LayoutBlocker,
  MeasureStackInput,
  MetricsResolver,
  ParagraphBox,
  ParagraphMarker,
  ParagraphPaint,
  PlacedCell,
  PlacedLine,
  StackMeasurement,
} from "./layout/stack.js";

export {
  borderExtentPt,
  readBorder,
  readBorders,
  readShading,
  readTableBorders,
  resolveCellBorders,
  SIDES,
} from "./docx/borders.js";
export type {
  Border,
  Borders,
  BorderSide,
  BorderStyle,
  StatedBorders,
  TableBorders,
} from "./docx/borders.js";

export { paintOfCell, paintOfParagraph, PARAGRAPH_PAINT_PT } from "./layout/painting.js";
export type { Painted, PaintedFill, PaintedLine } from "./layout/painting.js";

export { fitLine, freeSpans } from "./layout/wrapping.js";
export type { FitLineInput, LineSlot, WrapBand } from "./layout/wrapping.js";

export { breakStack } from "./layout/pages.js";
export type { BreakStackInput, PageStack } from "./layout/pages.js";

export { layOutTextBox } from "./layout/text-boxes.js";
export type {
  LayOutTextBoxInput,
  PlacedTextBox,
  TextBoxLayout,
  TextBoxRect,
} from "./layout/text-boxes.js";

export { layOutDocument } from "./layout/document.js";
export type { DocumentLayout, LaidOutDocument, LaidOutPage } from "./layout/document.js";

export { readAnchors, WHOLE_FRAME } from "./docx/anchors.js";
export type {
  AnchorOrigin,
  AnchorPosition,
  FloatingAnchor,
  WrapArea,
  WrapDistances,
  WrapMode,
} from "./docx/anchors.js";

export { placeFloat, UNPAINTED } from "./layout/floats.js";
export type {
  FloatSize,
  PartResolver,
  PlacedContent,
  PlacedFloat,
  PlacedPaint,
  PlaceFloatInput,
} from "./layout/floats.js";

export { readInlines } from "./docx/inlines.js";
export type { InlineDrawing } from "./docx/inlines.js";

export { placeInlines } from "./layout/inlines.js";
export type { PlacedInline, PlaceInlinesInput } from "./layout/inlines.js";

export { resolveParagraphFrame, resolveParagraphNumbering } from "./docx/styles.js";
export type { ParagraphAlignment, ParagraphFrame, ParagraphNumbering } from "./docx/styles.js";

export {
  readDrawingContent,
  readDrawingFlip,
  DEFAULT_TEXT_INSETS,
  NO_CROP,
  NO_PAINT,
  PIC_NS,
  WPS_NS,
} from "./docx/drawing.js";
export type {
  CropInsets,
  DrawingContent,
  DrawingFlip,
  ShapeGeometry,
  ShapeOutline,
  ShapePaint,
  TextBoxAnchor,
  TextBoxBody,
  TextBoxInsets,
} from "./docx/drawing.js";

export { readColorReference, readTheme, themeColor, NO_THEME, THEME_PART } from "./docx/theme.js";
export type { ColorReference, Theme } from "./docx/theme.js";

export { emuToPoints, twipsToPoints, EMU_PER_POINT, TWIPS_PER_POINT } from "./layout/units.js";

export { substitutingMetrics, WORD_FALLBACK_FACES } from "./layout/substitution.js";
export type { Substitution, SubstitutingMetrics } from "./layout/substitution.js";

export { readMetafilePicture } from "./metafile/picture.js";
export type { MetafilePicture, MetafileRect, MetafileShape } from "./metafile/picture.js";
