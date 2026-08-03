import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

import { readFontMetrics } from "@onepager/core";

import { referenceFonts } from "../testing/cases.js";

// Licensed faces are not copied into this repo. The manifest points at wherever
// each one lives and records the metrics Word laid the reference documents out with.
const FONTS = referenceFonts().filter(
  (font) => font.filePath !== null && existsSync(font.filePath),
);

describe.skipIf(FONTS.length === 0)("readFontMetrics on a real font file", () => {
  for (const font of FONTS) {
    it(`reads the metrics Word used for font ${String(FONTS.indexOf(font))}`, () => {
      const path = font.filePath ?? "";
      const result = readFontMetrics(new Uint8Array(readFileSync(path)));

      if (font.fileFormat !== null) expect(result.format).toBe(font.fileFormat);
      expect(result.metrics).toStrictEqual(font.metrics);
    });
  }
});
