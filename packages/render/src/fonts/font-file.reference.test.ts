import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

import { readFontFile, type GlyphAdvances } from "@onepager/core";

import { referenceFonts } from "../testing/cases.js";

// Licensed faces are not copied into this repo. The manifest points at wherever
// each one lives and records the metrics Word laid the reference documents out with.
const FONTS = referenceFonts().filter(
  (font) => font.filePath !== null && existsSync(font.filePath),
);

const bytesOf = (path: string): Uint8Array => new Uint8Array(readFileSync(path));

const advanceOf = (advanceFor: GlyphAdvances, character: string): number | null =>
  advanceFor(character.codePointAt(0) ?? 0);

describe.skipIf(FONTS.length === 0)("readFontFile on a real font file", () => {
  for (const font of FONTS) {
    const index = String(FONTS.indexOf(font));
    const path = font.filePath ?? "";

    it(`reads the metrics Word used for font ${index}`, () => {
      const result = readFontFile(bytesOf(path));

      if (font.fileFormat !== null) expect(result.format).toBe(font.fileFormat);
      expect(result.metrics).toStrictEqual(font.metrics);
    });

    // Only relations are asserted; the face's real widths stay out of the repo.
    it(`reads plausible advances for font ${index}`, () => {
      const advances = readFontFile(bytesOf(path)).advances;
      if (advances.kind !== "advances") throw new Error(advances.reason);

      const { advanceFor } = advances;
      const space = advanceOf(advanceFor, " ");
      const wide = advanceOf(advanceFor, "M");
      const narrow = advanceOf(advanceFor, "i");

      expect(space).toBeGreaterThan(0);
      expect(narrow).toBeGreaterThan(0);
      expect(wide).toBeLessThanOrEqual(font.metrics.unitsPerEm * 2);
      expect(wide).toBeGreaterThan(narrow ?? 0);
      expect(narrow).toBeGreaterThan(space ?? 0);
      expect(advanceOf(advanceFor, "\u{10ffff}")).toBeNull();
    });
  }
});
