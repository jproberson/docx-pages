import { existsSync, readFileSync } from "node:fs";

import {
  readFontFile,
  WORD_CHARACTER_FALLBACK_FACES,
  WORD_FALLBACK_FACES,
  WORD_SERIF_FALLBACK_FACE,
  type AdvanceTable,
  type SuppliedFace,
} from "@docx-pages/core";

// Where a face on this machine lives. A collection holds several faces in one
// file, so the one wanted out of it is named: Cambria Math is the second face of
// the file Cambria is the first of.
type FaceFile = {
  readonly paths: readonly string[];
  readonly faceName?: string;
};

// Where each face Word falls back on is to be found, which is the whole of what
// this project knows about the machine it runs on. Cambria stands in for a face
// nothing supplies at all; the rest draw a character the face a run states has no
// glyph for, and which of them answers turns on the kind of face that asked and on
// the character. See `WORD_CHARACTER_FALLBACK_FACES`.
const FILES: Readonly<Record<string, FaceFile>> = {
  Cambria: {
    paths: [
      "/Applications/Microsoft Word.app/Contents/Resources/DFonts/Cambria.ttc",
      "/Library/Fonts/Microsoft/Cambria.ttc",
      "/System/Library/Fonts/Supplemental/Cambria.ttc",
      "/usr/share/fonts/truetype/msttcorefonts/Cambria.ttf",
    ],
  },
  "Cambria Math": {
    paths: [
      "/Applications/Microsoft Word.app/Contents/Resources/DFonts/Cambria.ttc",
      "/Library/Fonts/Microsoft/Cambria.ttc",
      "/System/Library/Fonts/Supplemental/Cambria.ttc",
    ],
    faceName: "Cambria Math",
  },
  Arial: {
    paths: [
      "/Applications/Microsoft Word.app/Contents/Resources/DFonts/arial.ttf",
      "/Library/Fonts/Microsoft/Arial.ttf",
      "/System/Library/Fonts/Supplemental/Arial.ttf",
      "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
    ],
  },
  // The system's copy comes before Word's own, which is the other way round from
  // Symbol. See `timesNewRoman` for what each of the two answers.
  [WORD_SERIF_FALLBACK_FACE]: {
    paths: [
      "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
      "/Library/Fonts/Microsoft/Times New Roman.ttf",
      "/Applications/Microsoft Word.app/Contents/Resources/DFonts/times.ttf",
      "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman.ttf",
    ],
  },
  "Segoe UI Symbol": {
    paths: [
      "/Applications/Microsoft Word.app/Contents/Resources/DFonts/seguisym.ttf",
      "/Library/Fonts/Microsoft/Segoe UI Symbol.ttf",
    ],
  },
  "Apple Color Emoji": {
    paths: ["/System/Library/Fonts/Apple Color Emoji.ttc"],
  },
};

// Word's own Times New Roman, which is a different file from the system's and is
// read for its glyphs alone. See `timesNewRoman`.
const WORD_TIMES_NEW_ROMAN = "/Applications/Microsoft Word.app/Contents/Resources/DFonts/times.ttf";

// Where a face Word falls back on lives on this machine, or null where it has not
// got it.
export const fallbackFacePath = (name: string): string | null =>
  (FILES[name]?.paths ?? []).find((each) => existsSync(each)) ?? null;

const faceOf = (name: string, path: string): SuppliedFace => {
  const read = readFontFile(new Uint8Array(readFileSync(path)), FILES[name]?.faceName);
  return {
    name,
    bold: false,
    italic: false,
    metrics: read.metrics,
    advances: read.advances,
    sansSerif: read.sansSerif,
  };
};

/**
 * Times New Roman as Word uses it, which is two files.
 *
 * Both answer to the name on a Mac and they are not the same face. The system's
 * copy states a line gap of 87 units in its `hhea` and Word's own states none, and
 * Word makes a Times New Roman line 13.80pt tall at 12pt, which is the gap of 87.
 * But Word draws the hyphen `U+2010` in the face at 682 units, which is a glyph
 * Word's copy has and the system's has not got at all. Measured on 2026-08-06 off
 * Word's own pdf of the authored `unmapped-in-a-text-face` document.
 *
 * So the heights come out of the system's copy and a glyph out of Word's where the
 * system's has none. Reading only the system's file refuses a document over that
 * hyphen; reading only Word's puts every Times New Roman line half a point short.
 */
function timesNewRoman(path: string): SuppliedFace {
  const system = faceOf(WORD_SERIF_FALLBACK_FACE, path);
  if (path === WORD_TIMES_NEW_ROMAN || !existsSync(WORD_TIMES_NEW_ROMAN)) return system;

  const words = faceOf(WORD_SERIF_FALLBACK_FACE, WORD_TIMES_NEW_ROMAN);
  return { ...system, advances: eitherCopy(system.advances, words.advances) };
}

const eitherCopy = (first: AdvanceTable, second: AdvanceTable): AdvanceTable => {
  if (first.kind !== "advances" || second.kind !== "advances") return first;
  return {
    ...first,
    advanceFor: (codePoint) => first.advanceFor(codePoint) ?? second.advanceFor(codePoint),
  };
};

// The faces `substitutingMetrics` needs before it can stand anything down to what
// Word would have used. A machine without them is left as it was: a face nothing
// supplies still refuses the document rather than being drawn in a guess, and so
// does a character no face the machine has a glyph for.
export function fallbackFaces(): readonly SuppliedFace[] {
  return [...WORD_FALLBACK_FACES, ...WORD_CHARACTER_FALLBACK_FACES].flatMap((name) => {
    const path = fallbackFacePath(name);
    if (path === null) return [];
    return [name === WORD_SERIF_FALLBACK_FACE ? timesNewRoman(path) : faceOf(name, path)];
  });
}
