// Everything the library promises, and the plumbing underneath it as well.
//
// This module is deliberately absent from the package's `exports`, so nothing that
// installs `@docx-pages/core` can reach it: what a reader of pages is promised is
// `index.ts` and only that. It exists for `render`, which is this repository's
// Word oracle rather than a consumer of the library, and which reads a document's
// raw XML, censuses a corpus and drives the layout a stage at a time. `render` is
// never published and consumes this source rather than a built package.
//
// A name here is not a name anyone outside can depend on, so moving one costs
// nothing beyond this repository.

export * from "./index.js";

export { partText, partXml } from "./docx/package.js";

// `partXml` answers with one of these, so anything walking a part needs to be able
// to name it and to read an attribute off it.
export { attribute, childrenNamed, firstNamed } from "./docx/xml.js";

export { pageGeometrySignature, readSectionGeometry, W_NS } from "./docx/section.js";

export {
  honoursAWrapOnTheLeft,
  readDocumentSettings,
  roundsAnchorsToTwips,
  takesTheRightOnEqualSides,
  DEFAULT_SETTINGS,
  SETTINGS_PART,
} from "./docx/settings.js";

export {
  readBlocks,
  readParagraphs,
  blocksIn,
  blockParagraphs,
  DEFAULT_TABLE_INSETS,
  MC_NS,
} from "./docx/blocks.js";

export { paragraphText, paragraphDescendants } from "./docx/paragraphs.js";

export { facesUsed } from "./docx/faces.js";

export { numberParagraphs } from "./docx/list-numbers.js";

export { readNumberingTable, numberingLevel, NUMBERING_PART } from "./docx/numbering.js";

export { readRuns } from "./docx/runs.js";

export {
  readStyleTable,
  resolveParagraphFrame,
  resolveParagraphMark,
  resolveParagraphNumbering,
  resolveRuns,
  A_NS,
  WORD_DEFAULT_FONT_SIZE_PT,
} from "./docx/styles.js";

export {
  readRelationships,
  relationshipsPartFor,
  defaultFooterPart,
  defaultHeaderPart,
  R_NS,
  PKG_REL_NS,
} from "./docx/relationships.js";

export { breakLines, faceRequestFor, measureText } from "./layout/lines.js";

export { nextTabStop, tabStopsPt, DEFAULT_TAB_STOP_PT } from "./layout/tab-stops.js";

export { measureStack, shiftBox, shiftBoxes, shiftCells, WP_NS } from "./layout/stack.js";

export {
  borderExtentPt,
  readBorder,
  readBorders,
  readShading,
  readTableBorders,
  resolveCellBorders,
  SIDES,
} from "./docx/borders.js";

export { fitLine, freeSpans } from "./layout/wrapping.js";

export { breakStack } from "./layout/pages.js";

export { layOutTextBox } from "./layout/text-boxes.js";

export { readAnchors } from "./docx/anchors.js";

export { placeFloat } from "./layout/floats.js";

export { readInlines } from "./docx/inlines.js";

export { placeInlines } from "./layout/inlines.js";

export { readDrawingContent, readDrawingFlip, PIC_NS, WPS_NS } from "./docx/drawing.js";
