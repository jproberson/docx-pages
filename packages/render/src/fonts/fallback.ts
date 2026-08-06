import { existsSync, readFileSync } from "node:fs";

import {
  readFontFile,
  WORD_CHARACTER_FALLBACK_FACE,
  WORD_FALLBACK_FACES,
  type SuppliedFace,
} from "@docx-pages/core";

// Where each face Word falls back on is to be found. Cambria ships with Word and
// with nothing else on a Mac, and only inside a collection, which is why
// `readFontFile` reads one. Times New Roman is the other, and it is not the same
// job: Cambria stands in for a whole face nothing supplies, and Times New Roman
// draws the one character a symbol face cannot.
const FILES: Readonly<Record<string, readonly string[]>> = {
  Cambria: [
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts/Cambria.ttc",
    "/Library/Fonts/Microsoft/Cambria.ttc",
    "/System/Library/Fonts/Supplemental/Cambria.ttc",
    "/usr/share/fonts/truetype/msttcorefonts/Cambria.ttf",
  ],
  // The system's copy comes before Word's own, which is the other way round from
  // Symbol. The two differ in one number, the line gap their `hhea` states: 87
  // units in the system's and none in Word's. Word makes a Times New Roman line
  // 13.80pt tall at 12pt, which is the gap of 87.
  [WORD_CHARACTER_FALLBACK_FACE]: [
    "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
    "/Library/Fonts/Microsoft/Times New Roman.ttf",
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts/times.ttf",
    "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman.ttf",
  ],
};

// Where a face Word falls back on lives on this machine, or null where it has not
// got it.
export const fallbackFacePath = (name: string): string | null =>
  (FILES[name] ?? []).find((each) => existsSync(each)) ?? null;

// The faces `substitutingMetrics` needs before it can stand anything down to what
// Word would have used. A machine without them is left as it was: a face nothing
// supplies still refuses the document rather than being drawn in a guess.
export function fallbackFaces(): readonly SuppliedFace[] {
  return [...WORD_FALLBACK_FACES, WORD_CHARACTER_FALLBACK_FACE].flatMap((name) => {
    const path = fallbackFacePath(name);
    if (path === null) return [];

    const read = readFontFile(new Uint8Array(readFileSync(path)));
    return [{ name, bold: false, italic: false, metrics: read.metrics, advances: read.advances }];
  });
}
