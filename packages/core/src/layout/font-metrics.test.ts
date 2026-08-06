import { describe, expect, it } from "vitest";

import {
  advanceWidthPt,
  ascentPt,
  lineHeightPt,
  lookupFontMetrics,
  NO_ADVANCES,
  type FaceRequest,
  type FontMetrics,
  type SuppliedFace,
} from "./font-metrics.js";

const asked = (name: string, bold = false, italic = false): FaceRequest => ({
  name,
  bold,
  italic,
});

const found = (name: string, supplied?: readonly SuppliedFace[]): FontMetrics => {
  const lookup = lookupFontMetrics(asked(name), supplied);
  if (lookup.kind !== "found") throw new Error(`expected metrics for ${name}`);
  return lookup.metrics;
};

const face = (
  name: string,
  metrics: FontMetrics,
  style: Partial<FaceRequest> = {},
): SuppliedFace => ({
  name,
  bold: style.bold ?? false,
  italic: style.italic ?? false,
  metrics,
  advances: NO_ADVANCES,
});

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
  it("counts the line gap above the ascender, since that is where Word leads", () => {
    expect(ascentPt(found("Arial"), 12)).toBeCloseTo(11.2559, 4);
  });

  it("leaves only the descender below the baseline", () => {
    const arial = found("Arial");
    expect(lineHeightPt(arial, 12) - ascentPt(arial, 12)).toBeCloseTo(2.543, 4);
  });
});

describe("advanceWidthPt", () => {
  it("scales a font-unit advance by the size it is set at", () => {
    expect(
      advanceWidthPt(500, { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 }, 12),
    ).toBeCloseTo(6, 10);
  });
});

describe("lookupFontMetrics", () => {
  const NARROW: FontMetrics = { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 };
  const SUPPLIED: FontMetrics = { unitsPerEm: 2048, ascender: 1944, descender: -546, lineGap: 0 };
  const advances: SuppliedFace["advances"] = {
    kind: "advances",
    advanceFor: () => 500,
  };

  it("resolves a built-in font", () => {
    expect(lookupFontMetrics(asked("Times New Roman"))).toStrictEqual({
      kind: "found",
      source: "builtin",
      metrics: { unitsPerEm: 2048, ascender: 1825, descender: -443, lineGap: 87 },
      advances: NO_ADVANCES,
    });
  });

  it("reports a built-in font as unmeasurable for text, since no file supplied its widths", () => {
    const lookup = lookupFontMetrics(asked("Arial"));
    expect(lookup.kind === "found" && lookup.advances).toStrictEqual(NO_ADVANCES);
  });

  it("ignores case and surrounding whitespace in the font name", () => {
    expect(found("  arial  ")).toStrictEqual(found("Arial"));
  });

  it("reports a missing font rather than substituting a default", () => {
    expect(lookupFontMetrics(asked("Meridian Sans"))).toStrictEqual({
      kind: "missing",
      fontName: "Meridian Sans",
    });
  });

  it("takes supplied metrics for a font it does not know", () => {
    const supplied = [face("Meridian Sans", SUPPLIED)];

    expect(lookupFontMetrics(asked("Meridian Sans"), supplied)).toStrictEqual({
      kind: "found",
      source: "supplied",
      metrics: SUPPLIED,
      advances: NO_ADVANCES,
    });
  });

  it("lets supplied metrics override a built-in, so a substituted face keeps the original metrics", () => {
    expect(lookupFontMetrics(asked("Arial"), [face("Arial", NARROW)])).toStrictEqual({
      kind: "found",
      source: "supplied",
      metrics: NARROW,
      advances: NO_ADVANCES,
    });
  });

  it("carries a supplied face's advances through to the caller", () => {
    const supplied = [{ ...face("Meridian Sans", NARROW), advances }];
    const lookup = lookupFontMetrics(asked("Meridian Sans"), supplied);

    expect(lookup.kind === "found" && lookup.advances).toBe(advances);
  });

  it("picks the face whose style the run asks for", () => {
    const regular = { ...face("Meridian Sans", NARROW), advances };
    const bold = {
      ...face("Meridian Sans", NARROW, { bold: true }),
      advances: {
        kind: "advances",
        advanceFor: () => 600,
      } satisfies SuppliedFace["advances"],
    };
    const lookup = lookupFontMetrics(asked("Meridian Sans", true), [regular, bold]);

    expect(lookup.kind === "found" && lookup.advances).toBe(bold.advances);
  });

  // Placing paragraphs only needs the family's vertical metrics, so a missing bold
  // face still measures height; it just must not measure text at the wrong widths.
  it("keeps the family's metrics but refuses advances when the asked-for style is absent", () => {
    const supplied = [{ ...face("Meridian Sans", NARROW), advances }];
    const lookup = lookupFontMetrics(asked("Meridian Sans", true), supplied);

    expect(lookup).toStrictEqual({
      kind: "found",
      source: "supplied",
      metrics: NARROW,
      advances: { kind: "unavailable", reason: "style-unsupplied" },
    });
  });
});
