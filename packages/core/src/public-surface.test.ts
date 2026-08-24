import { describe, expect, it } from "vitest";

import * as core from "./index.js";

// What `@docx-pages/core` promises anyone who installs it, pinned by name, because
// nothing else here would notice a name going away: dropping an export typechecks
// perfectly inside this repository and breaks every consumer outside it. A name
// leaving this list is a breaking change and has to be spelt as one.
//
// Only what exists at run time is pinned, and only what the package's `exports`
// offers. The plumbing in `internal.ts` is not published and is not promised: it
// is `render`'s, and moving a name there costs nothing beyond this repository.
const CORE_SURFACE: readonly string[] = [
  "DEFAULT_TEXT_INSETS",
  "DocxPagesError",
  "EMU_PER_POINT",
  "FONT_TABLE_PART",
  "MAIN_DOCUMENT_PART",
  "METAFILE_EXTENSION",
  "METAFILE_PEN_OFFSET",
  "NO_ADVANCES",
  "NO_CROP",
  "NO_PAINT",
  "NO_THEME",
  "PARAGRAPH_PAINT_PT",
  "PICTURE_MEDIA_TYPES",
  "ROUNDED_CORNER_FRACTION",
  "THEME_PART",
  "TWIPS_PER_POINT",
  "UNPAINTED",
  "WHOLE_FRAME",
  "WORD_CHARACTER_FALLBACK_FACES",
  "WORD_EMOJI_FACE",
  "WORD_FALLBACK_FACES",
  "WORD_SANS_FALLBACK_FACE",
  "WORD_SERIF_FALLBACK_FACE",
  "advanceWidthPt",
  "aliasedSymbolCharacter",
  "aliasedSymbolText",
  "ascentPt",
  "bestEffortMetrics",
  "boundsOfTurn",
  "drawablePicture",
  "pngFromMetafile",
  "readMetafileBitmap",
  "OLD_METAFILE_EXTENSION",
  "drawablesOf",
  "isAliasedSymbolFace",
  "emuToPoints",
  "isDocxPagesError",
  "layOutDocument",
  "lineHeightPt",
  "lookupFontMetrics",
  "onTheDeviceGrid",
  "runWidthMadeUpBy",
  "openDocx",
  "paintOfCell",
  "paintOfParagraph",
  "paragraphBoxesOn",
  "pdfOfDocx",
  "pictureExtension",
  "readColorReference",
  "readFaceAlternatives",
  "readFaceShapes",
  "readFontFaces",
  "readFontFile",
  "readFontMetrics",
  "readGlyphIndex",
  "readSuppliedFace",
  "readMetafilePicture",
  "readTheme",
  "readUnhonoured",
  "roomForTurn",
  "substitutingMetrics",
  "themeColor",
  "turnedAbout",
  "turnsOnItsSide",
  "twipsToPoints",
  "unshowableIn",
  "withFallbackCharacters",
  "withMissingGlyphs",
  "withSubstitutedFaces",
  "writePdf",
];

describe("the public surface of @docx-pages/core", () => {
  it("exports what it promises, and nothing has quietly gone away", () => {
    const found = Object.keys(core)
      .filter((name) => name !== "default")
      .sort();

    expect(found).toStrictEqual([...CORE_SURFACE].sort());
  });
});
