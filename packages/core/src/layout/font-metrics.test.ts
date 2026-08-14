import { describe, expect, it } from "vitest";

import {
  advanceWidthPt,
  ascentPt,
  lineHeightPt,
  lookupFontMetrics,
  NO_ADVANCES,
  NO_KERNING,
  runKerns,
  runsKernAcross,
  type FaceRequest,
  type FontMetrics,
  type KerningTable,
  type RunKerning,
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

  // The line gap read 87 until 2026-08-14, which is 0.51pt a line at 12pt. Both this
  // machine's own copy and the 151 pdfs Word embedded `TimesNewRomanPSMT` into state
  // none at all.
  it("resolves a built-in font", () => {
    expect(lookupFontMetrics(asked("Times New Roman"))).toStrictEqual({
      kind: "found",
      source: "builtin",
      metrics: { unitsPerEm: 2048, ascender: 1825, descender: -443, lineGap: 0 },
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

// A line is the sum of its characters' advances and of what each pair of them
// moves, so the pairs travel the same road the widths do: off the file, onto the
// supplied face, and out of the lookup a run resolves to.
describe("the pairs a lookup carries", () => {
  const NARROW: FontMetrics = { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 };
  const kerning: KerningTable = {
    kind: "kerning",
    source: "kern",
    kerningBetween: () => -80,
    subtablesLeftUnread: 0,
  };

  it("carries a supplied face's pairs through to the caller", () => {
    const supplied = [{ ...face("Meridian Sans", NARROW), kerning }];
    const lookup = lookupFontMetrics(asked("Meridian Sans"), supplied);

    expect(lookup.kind === "found" && lookup.kerning).toBe(kerning);
  });

  // The pairs of a near miss are as much the wrong style's as its widths are.
  it("refuses the pairs when the asked-for style is absent, as it refuses the widths", () => {
    const supplied = [{ ...face("Meridian Sans", NARROW), kerning }];
    const lookup = lookupFontMetrics(asked("Meridian Sans", true), supplied);

    expect(lookup.kind === "found" && lookup.kerning).toStrictEqual({
      kind: "unavailable",
      reason: "style-unsupplied",
    });
  });

  // A face nothing asked for its pairs is not a face that states none, so a lookup
  // over one says nothing about them at all.
  it("says nothing about the pairs of a face that was never asked for them", () => {
    const lookup = lookupFontMetrics(asked("Meridian Sans"), [face("Meridian Sans", NARROW)]);

    expect(lookup.kind === "found" && lookup.kerning).toBeUndefined();
  });

  it("says nothing about the pairs of a built-in font, which comes from no file", () => {
    const lookup = lookupFontMetrics(asked("Arial"));

    expect(lookup.kind === "found" && lookup.kerning).toBeUndefined();
    expect(NO_KERNING).toStrictEqual({ kind: "unavailable", reason: "unsupplied" });
  });
});

const CALIBRI = { name: "Calibri", bold: false, italic: false };

const set = (sizePt: number, kernFromHalfPoints: number | null): RunKerning => ({
  ...CALIBRI,
  sizePt,
  kernFromHalfPoints,
});

// Measured on 2026-08-13 off Word's own pdf, over right aligned lines of kerning
// pairs in Calibri written three times each, against a control line of no pairs
// that started at 448.34 in every case.
describe("runKerns", () => {
  // A line stating nothing starts at 427.95, exactly where one stating zero
  // starts, against 432.66 where it states one.
  it("says a run that states nothing does not kern, since kerning is opt-in", () => {
    expect(runKerns(set(12, null))).toBe(false);
  });

  it("says a run stating zero does not kern, which is where a run stating nothing lands", () => {
    expect(runKerns(set(12, 0))).toBe(false);
  });

  it("says a run stating a threshold its size reaches kerns", () => {
    expect(runKerns(set(12, 1))).toBe(true);
  });

  // Against `w:kern w:val="32"`: 15.5pt starts at 384.77, which is where that size
  // unkerned starts, and 16pt at 384.88, which is where that size kerned starts.
  it("kerns at the very size the threshold names", () => {
    expect(runKerns(set(16, 32))).toBe(true);
  });

  it("does not kern half a point under it", () => {
    expect(runKerns(set(15.5, 32))).toBe(false);
    expect(runKerns(set(11, 32))).toBe(false);
  });

  it("kerns above it", () => {
    expect(runKerns(set(16.5, 32))).toBe(true);
  });

  // The threshold is stated in half-points and the run is set in points, which is
  // what everything else in the layout is measured in.
  it("reads the threshold as half-points against a size in points", () => {
    expect(runKerns(set(8, 16))).toBe(true);
    expect(runKerns(set(7.5, 16))).toBe(false);
  });
});

// `WAVY` in one run and as `WA` beside `VY` were drawn identically, from 546.79
// and 29.20 wide; where the second run is bold its `V` starts at 562.67 and the
// first ends at 562.68, without the half point an `AV` pulls.
describe("runsKernAcross", () => {
  const KERNING = set(12, 1);

  it("lets a pair cross a boundary between runs set alike", () => {
    expect(runsKernAcross(KERNING, { ...KERNING })).toBe(true);
  });

  it("stops a pair where the weight changes, as Word drew the bold V", () => {
    expect(runsKernAcross(KERNING, { ...KERNING, bold: true })).toBe(false);
  });

  it("stops a pair where the slant changes", () => {
    expect(runsKernAcross(KERNING, { ...KERNING, italic: true })).toBe(false);
  });

  it("stops a pair where the size changes", () => {
    expect(runsKernAcross(KERNING, { ...KERNING, sizePt: 14 })).toBe(false);
  });

  it("stops a pair where the face changes", () => {
    expect(runsKernAcross(KERNING, { ...KERNING, name: "Cambria" })).toBe(false);
  });

  it("calls a face by the same name the same face however it is written", () => {
    expect(runsKernAcross(KERNING, { ...KERNING, name: "  calibri " })).toBe(true);
  });

  // A run that does not kern has no pairs to carry over a boundary, whichever side
  // of it the run is on.
  it("stops a pair where either run does not kern at all", () => {
    expect(runsKernAcross(set(12, null), KERNING)).toBe(false);
    expect(runsKernAcross(KERNING, set(12, null))).toBe(false);
    expect(runsKernAcross(set(12, 32), set(12, 32))).toBe(false);
  });
});
