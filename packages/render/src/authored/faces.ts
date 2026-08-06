import { existsSync, readFileSync } from "node:fs";

import {
  readFontFile,
  substitutingMetrics,
  type MetricsResolver,
  type SuppliedFace,
} from "@docx-pages/core";

import { fallbackFaces, fallbackFacePath } from "../fonts/fallback.js";
import { FACE } from "./package.js";

// The authored documents are all laid out in one face, so the suite needs that one
// file and nothing else. It is looked for rather than configured: unlike the
// reference documents, nothing here depends on a manifest that stays out of the
// repository.
const CANDIDATES: readonly string[] = [
  "/Applications/Microsoft Word.app/Contents/Resources/DFonts/Calibri.ttf",
  "/Library/Fonts/Microsoft/Calibri.ttf",
  "/Library/Fonts/Calibri.ttf",
  "/System/Library/Fonts/Supplemental/Calibri.ttf",
  "/usr/share/fonts/truetype/msttcorefonts/Calibri.ttf",
];

export const facePath = (): string | null => CANDIDATES.find((path) => existsSync(path)) ?? null;

// The face every authored document is written in, or null on a machine that does
// not have it, which leaves the suite with nothing to measure against.
export function authoredFace(): SuppliedFace | null {
  return faceOf(FACE, facePath());
}

// Where the faces an authored document names by hand are to be found. One document
// asks what a face draws for a character it has no glyph for, and that question
// cannot be asked in the face everything else is written in: it needs a face that
// leaves a character out. Each of them ships with Word or with the system, and none
// is anyone's collateral, so they are looked for in the same way as the one face.
const STATED: Readonly<Record<string, readonly string[]>> = {
  Arial: [
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts/arial.ttf",
    "/Library/Fonts/Microsoft/Arial.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
  ],
  // Word's own Symbol, which is not the Symbol macOS keeps in
  // `/System/Library/Fonts`: the two files differ, and the Macintosh table of the
  // system one maps a byte Word's maps nothing at.
  Symbol: [
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts/symbol.ttf",
    "/Library/Fonts/Microsoft/Symbol.ttf",
    "/System/Library/Fonts/Supplemental/Symbol.ttf",
  ],
  Wingdings: [
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts/Wingdings.ttf",
    "/Library/Fonts/Microsoft/Wingdings.ttf",
    "/System/Library/Fonts/Supplemental/Wingdings.ttf",
  ],
};

// A document may state the face Word itself reaches for, which the fallback
// already knows where to find and supplies.
const statedPath = (name: string): string | null =>
  (STATED[name] ?? []).find((path) => existsSync(path)) ?? fallbackFacePath(name);

// A face an authored document states, or null where this machine has not got it,
// which is what leaves the document that names it out of the suites.
export function statedFace(name: string): SuppliedFace | null {
  return faceOf(name, statedPath(name));
}

export const hasStatedFaces = (names: readonly string[]): boolean =>
  names.every((name) => statedPath(name) !== null);

// Every face an authored document is laid out in: the one they are all written in,
// the ones a document about faces states by hand, and the face Word draws a
// character none of the document's own has a glyph for.
export function authoredSuiteFaces(): readonly SuppliedFace[] {
  const face = authoredFace();
  return [
    ...(face === null ? [] : [face]),
    ...Object.keys(STATED).flatMap((name) => {
      const stated = statedFace(name);
      return stated === null ? [] : [stated];
    }),
    ...fallbackFaces(),
  ];
}

// How an authored document resolves a face, which is not `lookupFontMetrics`: a
// character no face of the document's own maps is drawn out of another face
// altogether, and only a resolver holding every face can say which. No fallback
// name is offered, so a face this machine has not got still fails as it did rather
// than being stood in for quietly.
export const authoredMetrics = (): MetricsResolver =>
  substitutingMetrics(authoredSuiteFaces(), []).metricsFor;

function faceOf(name: string, path: string | null): SuppliedFace | null {
  if (path === null) return null;

  const read = readFontFile(new Uint8Array(readFileSync(path)));
  return {
    name,
    bold: false,
    italic: false,
    metrics: read.metrics,
    advances: read.advances,
  };
}
