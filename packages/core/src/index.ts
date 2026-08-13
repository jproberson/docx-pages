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
  OLD_METAFILE_EXTENSION,
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

export {
  readFontFaces,
  readFontFile,
  readFontMetrics,
  readGlyphIndex,
} from "./layout/font-file.js";
export type {
  FontFaceName,
  FontFileFormat,
  ReadFontFileResult,
  ReadFontMetricsResult,
  UnderlineMetrics,
} from "./layout/font-file.js";
export type { CodeToGlyph } from "./layout/glyphs.js";

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

// What a page draws and in what order. A renderer that worked this out for itself
// would be answering the stacking question a second time, and the two answers
// would drift.

export { drawablesOf } from "./layout/drawables.js";
export type { Drawable, PaintedParagraph } from "./layout/drawables.js";

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

export { DEFAULT_TEXT_INSETS, NO_CROP, NO_PAINT, ROUNDED_CORNER_FRACTION } from "./docx/drawing.js";
export type {
  CropInsets,
  DrawingContent,
  DrawingFlip,
  PathCommand,
  PathPoint,
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

export { boundsOfTurn, roomForTurn, turnedAbout, turnsOnItsSide } from "./layout/turns.js";
export type { TurnedRect, TurnedSize } from "./layout/turns.js";

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

// Every paragraph a page draws, groups walked into, which anything comparing a page
// against something else has to read rather than `page.body`.
export { paragraphBoxesOn } from "./layout/document.js";

// What a laid-out page says is wrong with itself, which is a different question from
// what the document asked for and did not get.
export { unshowableIn } from "./layout/unshowable.js";
export type { Unshowable, UnshowableKind } from "./layout/unshowable.js";

export {
  aliasedSymbolCharacter,
  aliasedSymbolText,
  isAliasedSymbolFace,
} from "./layout/symbol-aliases.js";

export { readFaceShapes, FONT_TABLE_PART } from "./docx/font-table.js";
export type { FaceShape } from "./docx/font-table.js";

export { readMetafilePicture } from "./metafile/picture.js";
export { pngFromMetafile, readMetafileBitmap } from "./metafile/wmf.js";
export type { DecodedBitmap } from "./metafile/wmf.js";
export type { MetafilePicture, MetafileRect, MetafileShape } from "./metafile/picture.js";

// A laid-out document written out as a pdf, and the one convenience that opens and
// lays a `.docx` out before doing it.
//
// Two names, and it should stay about two. Writing decides nothing about where
// anything sits: it walks `drawablesOf`, the same traversal the viewer walks, so a
// page comes out of the file where it came out on the screen. A caller states the
// faces the document draws in and where its pictures come from, and there is no
// knob here that moves anything on a page.
export { writePdf } from "./pdf/document.js";
export type { PdfFont, PdfMetadata, WritePdfOptions } from "./pdf/document.js";

export { pdfOfDocx } from "./pdf/docx-pdf.js";
export type { PdfOfDocxOptions } from "./pdf/docx-pdf.js";
