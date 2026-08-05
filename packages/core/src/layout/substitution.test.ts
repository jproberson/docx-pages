import { describe, expect, it } from "vitest";

import { NO_ADVANCES, type SuppliedFace } from "./font-metrics.js";
import { substitutingMetrics } from "./substitution.js";

const METRICS = { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 };

const measurable = (name: string, style: Partial<SuppliedFace> = {}): SuppliedFace => ({
  name,
  bold: false,
  italic: false,
  metrics: METRICS,
  advances: { kind: "advances", advanceFor: () => 500 },
  ...style,
});

// A face whose vertical metrics are known and whose widths are not, which is what
// the builtin table hands back: enough to place an empty paragraph, not enough to
// measure a word.
const unmeasurable = (name: string): SuppliedFace => ({
  name,
  bold: false,
  italic: false,
  metrics: METRICS,
  advances: NO_ADVANCES,
});

const ask = (name: string, style: Partial<{ bold: boolean; italic: boolean }> = {}) => ({
  name,
  bold: false,
  italic: false,
  ...style,
});

describe("substitutingMetrics", () => {
  it("leaves a face the document has alone, and says nothing was stood in for", () => {
    const faces = substitutingMetrics([measurable("Meridian Sans")], ["Calibri"]);
    const found = faces.metricsFor(ask("Meridian Sans"));

    expect(found.kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([]);
  });

  it("stands the first usable fallback in for a face nothing supplies", () => {
    const faces = substitutingMetrics([measurable("Calibri")], ["Nothing Has This", "Calibri"]);
    const found = faces.metricsFor(ask("Meridian Sans"));

    expect(found.kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Meridian Sans"), used: ask("Calibri") },
    ]);
  });

  // Vertical metrics alone place an empty paragraph and cannot break a line, so a
  // face known only that far is stood in for as if it were missing.
  it("stands in for a face whose widths are unknown, not only for a missing one", () => {
    const faces = substitutingMetrics(
      [unmeasurable("Meridian Sans"), measurable("Calibri")],
      ["Calibri"],
    );

    expect(faces.metricsFor(ask("Meridian Sans")).kind).toBe("found");
    expect(faces.substitutions().map((each) => each.used.name)).toStrictEqual(["Calibri"]);
  });

  it("takes the style asked for over the same style of another face", () => {
    const faces = substitutingMetrics(
      [measurable("Calibri"), measurable("Calibri", { bold: true })],
      ["Calibri"],
    );

    expect(faces.metricsFor(ask("Meridian Sans", { bold: true }))).toMatchObject({ kind: "found" });
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Meridian Sans", { bold: true }), used: ask("Calibri", { bold: true }) },
    ]);
  });

  // A run drawn in its own family at the wrong weight is nearer than the right
  // weight of a stranger, so the family is tried without the style first.
  it("gives up the style before it gives up the face", () => {
    const faces = substitutingMetrics(
      [measurable("Meridian Sans"), measurable("Calibri", { italic: true })],
      ["Calibri"],
    );

    expect(faces.metricsFor(ask("Meridian Sans", { italic: true }))).toMatchObject({
      kind: "found",
    });
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Meridian Sans", { italic: true }), used: ask("Meridian Sans") },
    ]);
  });

  it("records a face once however many times the document asks for it", () => {
    const faces = substitutingMetrics([measurable("Calibri")], ["Calibri"]);
    faces.metricsFor(ask("Meridian Sans"));
    faces.metricsFor(ask("Meridian Sans"));
    faces.metricsFor(ask("Meridian Sans", { italic: true }));

    expect(faces.substitutions()).toHaveLength(2);
  });

  // A page laid out in nothing is worse than one that refuses to be laid out, so a
  // face no fallback answers for fails the way it always did.
  it("lets a face no fallback can answer for fail, rather than laying out in nothing", () => {
    const faces = substitutingMetrics([], ["Nothing Has This Either"]);

    expect(faces.metricsFor(ask("Meridian Sans"))).toStrictEqual({
      kind: "missing",
      fontName: "Meridian Sans",
    });
    expect(faces.substitutions()).toStrictEqual([]);
  });

  it("does not stand a face in for itself", () => {
    const faces = substitutingMetrics([unmeasurable("Calibri")], ["calibri"]);

    expect(faces.metricsFor(ask("Calibri")).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([]);
  });
});
