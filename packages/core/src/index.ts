// What `@docx-pages/core` promises: enough to open a document, lay it out, draw
// the result, and be told what the document asked for and did not get.
//
// Every type is exported here, deliberately, while only the functions a caller
// invokes are. Naming a shape is not the same as calling into the middle of the
// layout, and a caller who cannot name `ParagraphBox` or `TextLine` cannot write a
// renderer at all.
//
// The plumbing underneath is in `internal.ts`, which the package's `exports` does
// not offer, so nothing that installs this can reach it.

export { DocxPagesError, isDocxPagesError } from "./errors.js";
export type { ContextValue, ErrorContext, DocxPagesErrorInit } from "./errors.js";

export { openDocx, MAIN_DOCUMENT_PART } from "./docx/package.js";
export type { DocxPackage } from "./docx/package.js";

// `partXml` answers with one of these, so anything walking a part needs to be able
// to name it and to read an attribute off it.

export type { XmlElement } from "./docx/xml.js";

export type { PageMargin, SectionGeometry } from "./docx/section.js";
export {
  readUnhonoured,
  withFallbackCharacters,
  withMissingGlyphs,
  withSubstitutedFaces,
} from "./docx/fidelity.js";
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
  BorrowedGlyph,
  FaceElsewhere,
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

export type {
  Block,
  CellVerticalAlign,
  Paragraph,
  TableCell,
  TableInsets,
  TableRow,
} from "./docx/blocks.js";

export type { UsedFace } from "./docx/faces.js";

export type { ParagraphNumber, ParagraphNumbers } from "./docx/list-numbers.js";

export type {
  LevelRestart,
  NumberFormat,
  NumberingLevel,
  NumberingTable,
  NumberSuffix,
} from "./docx/numbering.js";

export type { RunPiece, TextRun } from "./docx/runs.js";

export type { FontChoice, LineRule, MarkedRun, ParagraphMark, StyleTable } from "./docx/styles.js";

export type { Relationship } from "./docx/relationships.js";

export type {
  BreakLinesInput,
  LineBreaking,
  TextMeasurement,
  LineSegment,
  LineTabs,
  MeasureFailure,
  TextLine,
} from "./layout/lines.js";

export { type TabStopPt } from "./layout/tab-stops.js";

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

export type { BandSide, FitLineInput, LineSlot, WrapBand } from "./layout/wrapping.js";

export type { BreakStackInput, PageStack } from "./layout/pages.js";

export type {
  LayOutTextBoxInput,
  PlacedTextBox,
  TextBoxLayout,
  TextBoxRect,
} from "./layout/text-boxes.js";

export { layOutDocument } from "./layout/document.js";
export type { DocumentLayout, LaidOutDocument, LaidOutPage } from "./layout/document.js";

export { WHOLE_FRAME } from "./docx/anchors.js";
export type {
  AnchorOrigin,
  AnchorPosition,
  FloatingAnchor,
  WrapArea,
  WrapDistances,
  WrapMode,
  WrapSide,
} from "./docx/anchors.js";

export { UNPAINTED } from "./layout/floats.js";
export type {
  FloatSize,
  PartResolver,
  PlacedContent,
  PlacedFloat,
  PlacedPaint,
  PlaceFloatInput,
} from "./layout/floats.js";

export type { InlineDrawing } from "./docx/inlines.js";

export type { PlacedInline, PlaceInlinesInput } from "./layout/inlines.js";

export type { ParagraphAlignment, ParagraphFrame, ParagraphNumbering } from "./docx/styles.js";

export { DEFAULT_TEXT_INSETS, NO_CROP, NO_PAINT } from "./docx/drawing.js";
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

export {
  substitutingMetrics,
  WORD_CHARACTER_FALLBACK_FACES,
  WORD_EMOJI_FACE,
  WORD_FALLBACK_FACES,
  WORD_SANS_FALLBACK_FACE,
  WORD_SERIF_FALLBACK_FACE,
} from "./layout/substitution.js";
export type {
  FallbackCharacter,
  FallbackNames,
  MissingGlyph,
  Substitution,
  SubstitutingMetrics,
} from "./layout/substitution.js";

export { bestEffortMetrics } from "./layout/best-effort.js";
export type { BestEffortMetrics, FaceDefaults } from "./layout/best-effort.js";

export { readFaceShapes, FONT_TABLE_PART } from "./docx/font-table.js";
export type { FaceShape } from "./docx/font-table.js";

export { readMetafilePicture } from "./metafile/picture.js";
export type { MetafilePicture, MetafileRect, MetafileShape } from "./metafile/picture.js";
