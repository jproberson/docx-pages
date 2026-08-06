import { describe, expect, it } from "vitest";

import { buildFace } from "../testing/build-font.js";
import { NO_ADVANCES, type SuppliedFace } from "./font-metrics.js";
import { substitutingMetrics, WORD_CHARACTER_FALLBACK_FACE } from "./substitution.js";

const METRICS = { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 };

const measurable = (name: string, style: Partial<SuppliedFace> = {}): SuppliedFace => ({
  name,
  bold: false,
  italic: false,
  metrics: METRICS,
  advances: { kind: "advances", symbolEncoded: false, advanceFor: () => 500 },
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

// Word draws a character a symbol face cannot alias into its own page out of Times
// New Roman, which is the last of the three rules an unmapped character meets and
// the only one that leaves the face the run states. Measured on 2026-08-06 off
// Word's pdf of the authored `unmapped-characters` document.
describe("a character the face it stands in cannot draw", () => {
  const SERIF = { unitsPerEm: 2048, ascender: 1825, descender: -443, lineGap: 87 };

  const BULLET = "\u2022";
  const NO_BREAK_SPACE = "\u00a0";

  // A symbol face carrying no bullet and no low byte to alias one to, which
  // answers for the no-break space out of its own .notdef.
  const symbols = buildFace({
    name: "Meridian Symbols",
    metrics: METRICS,
    subtables: ["symbol"],
    characters: "AB",
    advance: 480,
    notdefAdvance: 1229,
  });

  const lastResort = buildFace({
    name: WORD_CHARACTER_FALLBACK_FACE,
    metrics: SERIF,
    characters: BULLET,
    advance: 717,
  });

  const foundIn = (
    faces: ReturnType<typeof substitutingMetrics>,
    name: string,
    style: Partial<{ bold: boolean; italic: boolean }> = {},
  ) => {
    const found = faces.metricsFor(ask(name, style));
    if (found.kind !== "found") throw new Error(found.fontName);
    return found;
  };

  const codePointOf = (character: string): number => character.codePointAt(0) ?? 0;

  it("is drawn out of Word's fallback face, which comes back whole", () => {
    const faces = substitutingMetrics([symbols, lastResort], []);
    const { elsewhere } = foundIn(faces, "Meridian Symbols");

    // Whole, because Word measures the line over the borrowed face as well: the
    // width is no use without the metrics it is a share of.
    expect(elsewhere?.advanceFor(codePointOf(BULLET))).toBe(717);
    expect(elsewhere?.metrics).toStrictEqual(SERIF);
  });

  it("says which character it was and what drew it", () => {
    const faces = substitutingMetrics([symbols, lastResort], []);
    foundIn(faces, "Meridian Symbols").elsewhere?.advanceFor(codePointOf(BULLET));

    expect(faces.fallbackCharacters()).toStrictEqual([
      {
        face: ask("Meridian Symbols"),
        used: ask(WORD_CHARACTER_FALLBACK_FACE),
        codePoint: codePointOf(BULLET),
      },
    ]);
  });

  it("records a character once however many times the layout measures it", () => {
    const faces = substitutingMetrics([symbols, lastResort], []);
    const { elsewhere } = foundIn(faces, "Meridian Symbols");
    elsewhere?.advanceFor(codePointOf(BULLET));
    elsewhere?.advanceFor(codePointOf(BULLET));

    expect(faces.fallbackCharacters()).toHaveLength(1);
  });

  // The face's own page answers before anything else is asked, so nothing the
  // face can draw at all reaches another one.
  it("leaves a character the face's own page answers for where it is", () => {
    const faces = substitutingMetrics([symbols, lastResort], []);
    const found = foundIn(faces, "Meridian Symbols");
    if (found.advances.kind !== "advances") throw new Error(found.advances.reason);

    expect(found.advances.advanceFor(codePointOf(NO_BREAK_SPACE))).toBe(1229);
    expect(faces.fallbackCharacters()).toStrictEqual([]);
  });

  // What a text face does with a character it has no glyph for is a different
  // question and an unmeasured one, so it fails as it did rather than taking a
  // width from a face Word may never have reached for.
  it("is not asked of a text face, which fails as it did", () => {
    const text = buildFace({ name: "Meridian Sans", metrics: METRICS, characters: "AB" });
    const faces = substitutingMetrics([text, lastResort], []);

    expect(foundIn(faces, "Meridian Sans").elsewhere).toBeUndefined();
  });

  it("reaches for nothing on a machine that has not got the fallback face", () => {
    const faces = substitutingMetrics([symbols], []);

    expect(foundIn(faces, "Meridian Symbols").elsewhere).toBeUndefined();
  });

  // A bold run's fallback weight is unmeasured, and a width out of the regular
  // weight of the right face is nearer than no page at all.
  it("takes the plain style of the fallback face where the machine has only that", () => {
    const faces = substitutingMetrics([{ ...symbols, bold: true }, lastResort], []);
    const { elsewhere } = foundIn(faces, "Meridian Symbols", { bold: true });

    expect(elsewhere?.advanceFor(codePointOf(BULLET))).toBe(717);
    expect(faces.fallbackCharacters()[0]?.used).toStrictEqual(ask(WORD_CHARACTER_FALLBACK_FACE));
  });
});
