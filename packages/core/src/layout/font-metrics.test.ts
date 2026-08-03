import { describe, expect, it } from "vitest";

import { ascentPt, lineHeightPt, lookupFontMetrics, type FontMetrics } from "./font-metrics.js";

const found = (name: string, supplied?: ReadonlyMap<string, FontMetrics>): FontMetrics => {
  const lookup = lookupFontMetrics(name, supplied);
  if (lookup.kind !== "found") throw new Error(`expected metrics for ${name}`);
  return lookup.metrics;
};

describe("lineHeightPt", () => {
  it("matches the line height Word produced for Arial", () => {
    // Derived from Word's own output: 7 empty 12pt paragraphs spanned 96.65pt.
    expect(lineHeightPt(found("Arial"), 12)).toBeCloseTo(13.7988, 4);
  });

  it("scales linearly with font size", () => {
    const arial = found("Arial");
    expect(lineHeightPt(arial, 24)).toBeCloseTo(lineHeightPt(arial, 12) * 2, 10);
  });

  it("is zero at zero size", () => {
    expect(lineHeightPt(found("Arial"), 0)).toBe(0);
  });
});

describe("ascentPt", () => {
  it("uses the ascender alone, not the full line height", () => {
    expect(ascentPt(found("Arial"), 12)).toBeCloseTo(10.8633, 4);
  });
});

describe("lookupFontMetrics", () => {
  it("resolves a built-in font", () => {
    expect(lookupFontMetrics("Times New Roman")).toStrictEqual({
      kind: "found",
      source: "builtin",
      metrics: { unitsPerEm: 2048, ascender: 1825, descender: -443, lineGap: 87 },
    });
  });

  it("ignores case and surrounding whitespace in the font name", () => {
    expect(found("  arial  ")).toStrictEqual(found("Arial"));
  });

  it("reports a missing font rather than substituting a default", () => {
    expect(lookupFontMetrics("Meridian Sans")).toStrictEqual({
      kind: "missing",
      fontName: "Meridian Sans",
    });
  });

  it("takes supplied metrics for a font it does not know", () => {
    const metrics: FontMetrics = { unitsPerEm: 2048, ascender: 1944, descender: -546, lineGap: 0 };

    expect(lookupFontMetrics("Meridian Sans", new Map([["Meridian Sans", metrics]]))).toStrictEqual(
      { kind: "found", source: "supplied", metrics },
    );
  });

  it("lets supplied metrics override a built-in, so a substituted face keeps the original metrics", () => {
    const narrower: FontMetrics = { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 };
    const lookup = lookupFontMetrics("Arial", new Map([["Arial", narrower]]));

    expect(lookup).toStrictEqual({ kind: "found", source: "supplied", metrics: narrower });
  });
});
