import { existsSync, readFileSync } from "node:fs";

import { readFontFile, type SuppliedFace } from "@onepager/core";

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
  const path = facePath();
  if (path === null) return null;

  const read = readFontFile(new Uint8Array(readFileSync(path)));
  return {
    name: FACE,
    bold: false,
    italic: false,
    metrics: read.metrics,
    advances: read.advances,
  };
}
