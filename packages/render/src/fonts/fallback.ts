import { existsSync, readFileSync } from "node:fs";

import { readFontFile, WORD_FALLBACK_FACES, type SuppliedFace } from "@docx-pages/core";

// Where each face Word falls back on is to be found. Cambria ships with Word and
// with nothing else on a Mac, and only inside a collection, which is why
// `readFontFile` reads one.
const FILES: Readonly<Record<string, readonly string[]>> = {
  Cambria: [
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts/Cambria.ttc",
    "/Library/Fonts/Microsoft/Cambria.ttc",
    "/System/Library/Fonts/Supplemental/Cambria.ttc",
    "/usr/share/fonts/truetype/msttcorefonts/Cambria.ttf",
  ],
};

// The faces `substitutingMetrics` needs before it can stand anything down to what
// Word would have used. A machine without them is left as it was: a face nothing
// supplies still refuses the document rather than being drawn in a guess.
export function fallbackFaces(): readonly SuppliedFace[] {
  return WORD_FALLBACK_FACES.flatMap((name) => {
    const path = (FILES[name] ?? []).find((each) => existsSync(each));
    if (path === undefined) return [];

    const read = readFontFile(new Uint8Array(readFileSync(path)));
    return [{ name, bold: false, italic: false, metrics: read.metrics, advances: read.advances }];
  });
}
