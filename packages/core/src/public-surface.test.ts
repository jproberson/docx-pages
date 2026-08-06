import { describe, expect, it } from "vitest";

import * as core from "./index.js";

// What `@docx-pages/core` promises anyone who installs it, pinned by name in one
// list, because nothing else here would notice a name going away: dropping an
// export typechecks perfectly inside this repository and breaks every consumer
// outside it. A name leaving this list is a breaking change and has to be spelt as
// one; a name joining it is not, and the list only wants updating.
//
// Only what exists at run time is pinned. A type-only export cannot be seen from
// here, and `tsc --build` is what holds those to account.
//
// It is a wide list, and a good deal of it is plumbing that `render` reaches for
// rather than anything a reader of pages needs. Narrowing it is a decision owed
// before 1.0, and until then this at least stops it widening by accident.
const CORE_SURFACE: readonly string[] = [
  "A_NS",
  "DEFAULT_SETTINGS",
  "DEFAULT_TABLE_INSETS",
  "DEFAULT_TAB_STOP_PT",
  "DEFAULT_TEXT_INSETS",
  "DocxPagesError",
  "EMU_PER_POINT",
  "MAIN_DOCUMENT_PART",
  "MC_NS",
  "METAFILE_EXTENSION",
  "NO_ADVANCES",
  "NO_CROP",
  "NO_PAINT",
  "NO_THEME",
  "NUMBERING_PART",
  "PARAGRAPH_PAINT_PT",
  "PICTURE_MEDIA_TYPES",
  "PIC_NS",
  "PKG_REL_NS",
  "R_NS",
  "SETTINGS_PART",
  "SIDES",
  "THEME_PART",
  "TWIPS_PER_POINT",
  "UNPAINTED",
  "WHOLE_FRAME",
  "WORD_CHARACTER_FALLBACK_FACES",
  "WORD_DEFAULT_FONT_SIZE_PT",
  "WORD_EMOJI_FACE",
  "WORD_FALLBACK_FACES",
  "WORD_SANS_FALLBACK_FACE",
  "WORD_SERIF_FALLBACK_FACE",
  "WPS_NS",
  "WP_NS",
  "W_NS",
  "advanceWidthPt",
  "ascentPt",
  "attribute",
  "blockParagraphs",
  "blocksIn",
  "borderExtentPt",
  "breakLines",
  "breakStack",
  "childrenNamed",
  "defaultFooterPart",
  "defaultHeaderPart",
  "drawablePicture",
  "emuToPoints",
  "faceRequestFor",
  "facesUsed",
  "firstNamed",
  "fitLine",
  "freeSpans",
  "honoursAWrapOnTheLeft",
  "isDocxPagesError",
  "layOutDocument",
  "layOutTextBox",
  "lineHeightPt",
  "lookupFontMetrics",
  "measureStack",
  "measureText",
  "nextTabStop",
  "numberParagraphs",
  "numberingLevel",
  "openDocx",
  "pageGeometrySignature",
  "paintOfCell",
  "paintOfParagraph",
  "paragraphDescendants",
  "paragraphText",
  "partText",
  "partXml",
  "pictureExtension",
  "placeFloat",
  "placeInlines",
  "readAnchors",
  "readBlocks",
  "readBorder",
  "readBorders",
  "readColorReference",
  "readDocumentSettings",
  "readDrawingContent",
  "readDrawingFlip",
  "readFontFile",
  "readFontMetrics",
  "readInlines",
  "readMetafilePicture",
  "readNumberingTable",
  "readParagraphs",
  "readRelationships",
  "readRuns",
  "readSectionGeometry",
  "readShading",
  "readStyleTable",
  "readTableBorders",
  "readTheme",
  "readUnhonoured",
  "relationshipsPartFor",
  "resolveCellBorders",
  "resolveParagraphFrame",
  "resolveParagraphMark",
  "resolveParagraphNumbering",
  "resolveRuns",
  "roundsAnchorsToTwips",
  "shiftBox",
  "shiftBoxes",
  "shiftCells",
  "substitutingMetrics",
  "tabStopsPt",
  "takesTheRightOnEqualSides",
  "themeColor",
  "twipsToPoints",
  "withFallbackCharacters",
  "withSubstitutedFaces",
];

const exportsOf = (module: object): readonly string[] =>
  Object.keys(module)
    .filter((name) => name !== "default")
    .sort();

describe("the public surface of @docx-pages/core", () => {
  it("exports what it promises, and nothing has quietly gone away", () => {
    expect(exportsOf(core)).toStrictEqual([...CORE_SURFACE].sort());
  });
});
