import { existsSync, readFileSync } from "node:fs";

import {
  readSuppliedFace,
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

// The four cuts of Times New Roman, each of which is the same pair of files over
// again: where the system keeps its copy of the cut, in the order of the entry
// above, and where Word keeps its own. See `timesNewRoman`.
type TimesNewRomanCut = {
  readonly bold: boolean;
  readonly italic: boolean;
  readonly paths: readonly string[];
  readonly wordsOwn: string;
};

const TIMES_NEW_ROMAN_CUTS: readonly TimesNewRomanCut[] = [
  {
    bold: false,
    italic: false,
    paths: [
      "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
      "/Library/Fonts/Microsoft/Times New Roman.ttf",
      "/Applications/Microsoft Word.app/Contents/Resources/DFonts/times.ttf",
      "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman.ttf",
    ],
    wordsOwn: "/Applications/Microsoft Word.app/Contents/Resources/DFonts/times.ttf",
  },
  {
    bold: true,
    italic: false,
    paths: [
      "/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf",
      "/Library/Fonts/Microsoft/Times New Roman Bold.ttf",
      "/Applications/Microsoft Word.app/Contents/Resources/DFonts/timesbd.ttf",
      "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Bold.ttf",
    ],
    wordsOwn: "/Applications/Microsoft Word.app/Contents/Resources/DFonts/timesbd.ttf",
  },
  {
    bold: false,
    italic: true,
    paths: [
      "/System/Library/Fonts/Supplemental/Times New Roman Italic.ttf",
      "/Library/Fonts/Microsoft/Times New Roman Italic.ttf",
      "/Applications/Microsoft Word.app/Contents/Resources/DFonts/timesi.ttf",
      "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Italic.ttf",
    ],
    wordsOwn: "/Applications/Microsoft Word.app/Contents/Resources/DFonts/timesi.ttf",
  },
  {
    bold: true,
    italic: true,
    paths: [
      "/System/Library/Fonts/Supplemental/Times New Roman Bold Italic.ttf",
      "/Library/Fonts/Microsoft/Times New Roman Bold Italic.ttf",
      "/Applications/Microsoft Word.app/Contents/Resources/DFonts/timesbi.ttf",
      "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Bold_Italic.ttf",
    ],
    wordsOwn: "/Applications/Microsoft Word.app/Contents/Resources/DFonts/timesbi.ttf",
  },
];

// Where a face Word falls back on lives on this machine, or null where it has not
// got it.
export const fallbackFacePath = (name: string): string | null =>
  (FILES[name]?.paths ?? []).find((each) => existsSync(each)) ?? null;

// Cambria Math comes out of this list and out of no other, so this is a face whose
// MATH table decides whether an equation can be set at all, and whose pairs the disk
// scan beside it has always carried: a face out of the pack shadows the same file
// found on the disk, so anything left off here is a run asking to kern and measured
// without it. `readSuppliedFace` is what carries them now.
const faceOf = (name: string, path: string): SuppliedFace => {
  const inFile = FILES[name]?.faceName;
  return readSuppliedFace(
    new Uint8Array(readFileSync(path)),
    { name, bold: false, italic: false },
    inFile === undefined ? {} : { inFile },
  );
};

/**
 * Times New Roman as Word uses it, which is two files a cut.
 *
 * Both answer to the name on a Mac and they are not the same face. The system's
 * copy states a line gap of 87 units in its `hhea` and Word's own states none, and
 * Word makes a Times New Roman line 13.80pt tall at 12pt, which is the gap of 87.
 * But Word draws the hyphen `U+2010` in the face at 682 units, which is a glyph
 * Word's copy has and the system's has not got at all. Measured on 2026-08-06 off
 * Word's own pdf of the authored `unmapped-in-a-text-face` document, and again on
 * 2026-08-21 file by file: every cut of each copy answers as its regular does.
 *
 * The two agree about everything else: read cut by cut on 2026-08-21, every one of
 * the 527 glyphs the two copies both hold across Latin-1 and Latin Extended advances
 * the same in each, so which copy answers matters only where one of them has no
 * glyph at all.
 *
 * So each cut's heights come out of the system's copy and a glyph out of Word's
 * where the system's has none. Reading only the system's files refuses a document
 * over that hyphen; reading any cut out of Word's alone puts its every line
 * 0.4248pt short at 10pt, which is what drifted a bold reference list down the
 * page a line at a time when the pack offered no bold cut at all.
 */
function timesNewRoman(): readonly SuppliedFace[] {
  return TIMES_NEW_ROMAN_CUTS.flatMap((cut) => {
    const path = cut.paths.find((each) => existsSync(each));
    if (path === undefined) return [];

    const system = {
      ...faceOf(WORD_SERIF_FALLBACK_FACE, path),
      bold: cut.bold,
      italic: cut.italic,
    };
    if (path === cut.wordsOwn || !existsSync(cut.wordsOwn)) return [system];

    const words = faceOf(WORD_SERIF_FALLBACK_FACE, cut.wordsOwn);
    return [{ ...system, advances: eitherCopy(system.advances, words.advances) }];
  });
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
// Read once a run, as the disk scan beside it is: the pack's files do not change while
// a sweep runs, and every document was paying for them again.
let pack: readonly SuppliedFace[] | null = null;

export function fallbackFaces(): readonly SuppliedFace[] {
  pack ??= readPackFaces();
  return pack;
}

function readPackFaces(): readonly SuppliedFace[] {
  return [...WORD_FALLBACK_FACES, ...WORD_CHARACTER_FALLBACK_FACES].flatMap((name) => {
    if (name === WORD_SERIF_FALLBACK_FACE) return timesNewRoman();
    const path = fallbackFacePath(name);
    if (path === null) return [];
    return [faceOf(name, path)];
  });
}
