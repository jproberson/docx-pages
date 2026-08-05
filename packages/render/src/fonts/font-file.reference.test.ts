import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

import { readFontFile, type FontMetrics, type GlyphAdvances } from "@docx-pages/core";

import { referenceFontFiles, referenceFonts } from "../testing/cases.js";

type Subject = {
  readonly filePath: string;
  readonly fileFormat: string | null;
  readonly metrics: FontMetrics;
};

// Licensed faces are not copied into this repo. The manifest points at wherever
// each one lives: the faces Word laid the reference documents out in, plus any
// other file the reader is expected to cope with.
const SUBJECTS: readonly Subject[] = [...referenceFonts(), ...referenceFontFiles()]
  .filter((font): font is Subject => font.filePath !== null && existsSync(font.filePath))
  .filter((font, at, all) => all.findIndex((other) => other.filePath === font.filePath) === at);

const bytesOf = (path: string): Uint8Array => new Uint8Array(readFileSync(path));

const advanceOf = (advanceFor: GlyphAdvances, character: string): number | null =>
  advanceFor(character.codePointAt(0) ?? 0);

describe.skipIf(SUBJECTS.length === 0)("readFontFile on a real font file", () => {
  for (const font of SUBJECTS) {
    const index = String(SUBJECTS.indexOf(font));

    it(`reads the metrics Word used for font ${index}`, () => {
      const result = readFontFile(bytesOf(font.filePath));

      if (font.fileFormat !== null) expect(result.format).toBe(font.fileFormat);
      expect(result.metrics).toStrictEqual(font.metrics);
    });

    // Only relations are asserted; the faces' real widths stay out of the repo. A
    // monospaced face, a symbol one and a face that maps no Latin at all all answer
    // here, so the only relation every face holds to is that a wide glyph is no
    // narrower than a narrow one.
    it(`reads plausible advances for font ${index}`, () => {
      const advances = readFontFile(bytesOf(font.filePath)).advances;
      if (advances.kind !== "advances") throw new Error(advances.reason);

      const { advanceFor } = advances;
      const wide = advanceOf(advanceFor, "M");
      const narrow = advanceOf(advanceFor, "i");

      expect(advanceOf(advanceFor, " ")).toBeGreaterThan(0);
      expect(advanceOf(advanceFor, "\u{10ffff}")).toBeNull();

      // A face for another script maps neither, and is no less readable for it.
      expect(wide === null).toBe(narrow === null);
      if (wide === null || narrow === null) return;
      expect(wide).toBeLessThanOrEqual(font.metrics.unitsPerEm * 2);
      expect(wide).toBeGreaterThanOrEqual(narrow);
    });
  }
});
