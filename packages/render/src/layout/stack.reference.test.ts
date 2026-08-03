import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

import { layOutDocument, lookupFontMetrics, openDocx, type FontMetrics } from "@onepager/core";

const SAMPLES = resolve(process.env["ONEPAGER_SAMPLE_DIR"] ?? "samples");

// Read from the real face in a session that has it; these are its hhea values.
const MERIDIAN_SANS: FontMetrics = { unitsPerEm: 1000, ascender: 971, descender: -242, lineGap: 0 };
const metricsFor = (name: string) =>
  lookupFontMetrics(name, new Map([["Meridian Sans Medium", MERIDIAN_SANS]]));

const sample = (name: string) => resolve(SAMPLES, `${name}.docx`);

const layout = (name: string) => {
  const result = layOutDocument(openDocx(new Uint8Array(readFileSync(sample(name)))), metricsFor);
  if (result.kind !== "laid-out") throw new Error(`blocked: ${result.blocker.kind}`);
  return result;
};

const topOf = (name: string, index: number): number => {
  const box = layout(name).body[index];
  if (box === undefined) throw new Error(`no paragraph ${String(index)}`);
  return box.topPt;
};

const headerTopOf = (name: string, index: number): number => {
  const box = layout(name).header[index];
  if (box === undefined) throw new Error(`no header paragraph ${String(index)}`);
  return box.topPt;
};

const REFERENCE = "Reference";
const COMPARISON = "Reference Comparison";

describe.skipIf(!existsSync(sample(REFERENCE)))("paragraph stack against Word", () => {
  // Word's own output: these tops were derived from the anchor offsets in the
  // reference pdf, so they are what the layout has to reproduce.
  it("puts paragraph 13 where Word put it", () => {
    expect(topOf(REFERENCE, 13)).toBeCloseTo(249.998, 0);
  });

  it("puts paragraph 20 where Word put it", () => {
    expect(topOf(REFERENCE, 20)).toBeCloseTo(346.65, 0);
  });

  it("accumulates no drift between the two, which is the part that must be exact", () => {
    const span = topOf(REFERENCE, 20) - topOf(REFERENCE, 13);
    // 0.06pt over seven paragraphs, which is under the precision of the reference
    // numbers themselves. Any real drift would compound to far more than this.
    expect(Math.abs(span - (346.65 - 249.998))).toBeLessThan(0.1);
  });

  it("starts the body below the header rather than at the top margin", () => {
    const { bodyTopPt, page } = layout(REFERENCE);
    expect(bodyTopPt).toBeGreaterThan(page.margin.topTwips / 20);
    expect(bodyTopPt).toBeCloseTo(86.75, 0);
  });
});

// The Comparison header is a one-row table, so its two images sit side by side.
// Word drew them at these tops in the reference pdf; the banner has no vertical
// crop, but srcRect hides 7.272% off the top and bottom of the logo, so the
// logo's visible top is that much below the bitmap the pdf draws.
const BANNER = { topPt: 21.6, heightPt: 71.19 };
const LOGO_VISIBLE_TOP_PT = 21.36 + 0.07272 * 71.67;

describe.skipIf(!existsSync(sample(COMPARISON)))("header row against Word", () => {
  it("seats the banner at the top of the row it makes as tall as itself", () => {
    expect(headerTopOf(COMPARISON, 3)).toBeCloseTo(BANNER.topPt, 1);
  });

  it("centres the shorter logo in that row rather than stacking it above", () => {
    expect(headerTopOf(COMPARISON, 0)).toBeCloseTo(LOGO_VISIBLE_TOP_PT, 1);
  });

  it("resumes the header stack below the row, not below both images", () => {
    expect(headerTopOf(COMPARISON, 4)).toBeCloseTo(BANNER.topPt + BANNER.heightPt, 1);
  });

  it("starts the body where Word starts it", () => {
    expect(layout(COMPARISON).bodyTopPt).toBeCloseTo(118.26, 0);
  });
});
