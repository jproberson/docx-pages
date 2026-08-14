import { describe, expect, it } from "vitest";

import { buildFace } from "../testing/build-font.js";
import { NO_ADVANCES, type SuppliedFace } from "./font-metrics.js";
import {
  substitutingMetrics,
  WORD_EMOJI_FACE,
  WORD_SANS_FALLBACK_FACE,
  WORD_SERIF_FALLBACK_FACE,
} from "./substitution.js";

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

// **The family a name begins with, before any stranger.** Read off `9d343f6d69d7` on
// 2026-08-14: its headings ask for `Aptos Display`, which no machine here holds, and
// standing Cambria in for them made every heading's line 1.25pt short of Word's.
describe("the family a face's name begins with", () => {
  it("stands the family in before it reaches a fallback", () => {
    const faces = substitutingMetrics([measurable("Aptos"), measurable("Cambria")], ["Cambria"]);

    expect(faces.metricsFor(ask("Aptos Display")).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Aptos Display"), used: ask("Aptos") },
    ]);
  });

  it("keeps the style asked for where the family answers in it", () => {
    const faces = substitutingMetrics(
      [measurable("Arial"), measurable("Arial", { bold: true }), measurable("Cambria")],
      ["Cambria"],
    );

    expect(faces.metricsFor(ask("Arial Nova", { bold: true })).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Arial Nova", { bold: true }), used: ask("Arial", { bold: true }) },
    ]);
  });

  // The same order the request itself is tried in: a face of the right family at the
  // wrong weight is nearer than the right weight of a stranger.
  it("gives up the style before it gives up the family", () => {
    const faces = substitutingMetrics(
      [measurable("Arial"), measurable("Cambria", { bold: true })],
      ["Cambria"],
    );

    expect(faces.metricsFor(ask("Arial Nova", { bold: true })).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Arial Nova", { bold: true }), used: ask("Arial") },
    ]);
  });

  // **The case that must not change.** `Times New Roman` shortens to `Times New`,
  // which names no face at all, so the request falls through to the fallbacks exactly
  // as it did before there was a family step.
  it("falls through where the shortened name reaches nothing", () => {
    const faces = substitutingMetrics([measurable("Cambria")], ["Cambria"]);

    expect(faces.metricsFor(ask("Times New Roman")).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Times New Roman"), used: ask("Cambria") },
    ]);
  });

  it("shortens a name of one word to nothing, and reaches the fallback as before", () => {
    const faces = substitutingMetrics([measurable("Cambria")], ["Cambria"]);

    expect(faces.metricsFor(ask("Nonesuch")).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Nonesuch"), used: ask("Cambria") },
    ]);
  });

  // A face the document names is never given up for its own family, however many
  // words its name holds.
  it("leaves a face the document has alone, family or no family", () => {
    const faces = substitutingMetrics(
      [measurable("Aptos"), measurable("Aptos Display")],
      ["Cambria"],
    );

    expect(faces.metricsFor(ask("Aptos Display")).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([]);
  });

  // Vertical metrics alone cannot break a line, so a family that answers with them
  // is no more use than one the machine does not hold at all.
  it("passes over a family that cannot be measured", () => {
    const faces = substitutingMetrics([unmeasurable("Aptos"), measurable("Cambria")], ["Cambria"]);

    expect(faces.metricsFor(ask("Aptos Display")).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Aptos Display"), used: ask("Cambria") },
    ]);
  });

  // Only the last word goes. A name three words long reaches the family two of them
  // name and stops there, rather than walking back to the first.
  it("drops one word and no more", () => {
    const faces = substitutingMetrics(
      [measurable("Aptos"), measurable("Aptos Display"), measurable("Cambria")],
      ["Cambria"],
    );

    expect(faces.metricsFor(ask("Aptos Display Condensed")).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Aptos Display Condensed"), used: ask("Aptos Display") },
    ]);
  });

  it("takes the extra space in a name for the whitespace it is", () => {
    const faces = substitutingMetrics([measurable("Aptos"), measurable("Cambria")], ["Cambria"]);

    expect(faces.metricsFor(ask("  Aptos   Display  ")).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("  Aptos   Display  "), used: ask("Aptos") },
    ]);
  });
});

// The last of the three rules an unmapped character meets, and the only one that
// leaves the face the run states: see `WORD_CHARACTER_FALLBACK_FACES`.
describe("a character the face it stands in cannot draw", () => {
  const SERIF_METRICS = { unitsPerEm: 2048, ascender: 1825, descender: -443, lineGap: 87 };
  const SANS_METRICS = { unitsPerEm: 2048, ascender: 1854, descender: -434, lineGap: 67 };
  const EMOJI_METRICS = { unitsPerEm: 800, ascender: 800, descender: -250, lineGap: 0 };

  const BULLET = "\u2022";
  const SMALL_SQUARE = "\u25aa";
  const HYPHEN = "\u2010";
  const WORD_JOINER = "\u2060";
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

  // The five faces Word reaches for, each carrying what it was measured answering
  // for and no more, so which one answered is readable off the width alone.
  const serif = buildFace({
    name: WORD_SERIF_FALLBACK_FACE,
    metrics: SERIF_METRICS,
    characters: BULLET + SMALL_SQUARE,
    advance: 717,
    sansSerif: false,
  });

  const sans = buildFace({
    name: WORD_SANS_FALLBACK_FACE,
    metrics: SANS_METRICS,
    characters: BULLET + SMALL_SQUARE,
    advance: 726,
    sansSerif: true,
  });

  const emoji = buildFace({
    name: WORD_EMOJI_FACE,
    metrics: EMOJI_METRICS,
    characters: SMALL_SQUARE,
    advance: 800,
  });

  const maths = buildFace({
    name: "Cambria Math",
    metrics: SERIF_METRICS,
    characters: HYPHEN,
    advance: 680,
  });

  const symbolText = buildFace({
    name: "Segoe UI Symbol",
    metrics: SERIF_METRICS,
    characters: HYPHEN + WORD_JOINER,
    advance: 819,
  });

  const everywhere = [emoji, sans, serif, maths, symbolText];

  const textFace = (name: string, sansSerif: boolean) =>
    buildFace({ name, metrics: METRICS, characters: "AB", sansSerif });

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

  const drawn = (
    faces: ReturnType<typeof substitutingMetrics>,
    name: string,
    character: string,
    style: Partial<{ bold: boolean; italic: boolean }> = {},
  ) => foundIn(faces, name, style).elsewhere?.(codePointOf(character)) ?? null;

  it("is drawn out of the face Word reaches for, which comes back whole", () => {
    const faces = substitutingMetrics([symbols, ...everywhere], []);

    // Whole, because Word measures the line over the borrowed face as well: the
    // width is no use without the metrics it is a share of.
    expect(drawn(faces, "Meridian Symbols", BULLET)).toStrictEqual({
      metrics: SERIF_METRICS,
      advance: 717,
    });
  });

  it("says which character it was and what drew it", () => {
    const faces = substitutingMetrics([symbols, ...everywhere], []);
    drawn(faces, "Meridian Symbols", BULLET);

    expect(faces.fallbackCharacters()).toStrictEqual([
      {
        face: ask("Meridian Symbols"),
        used: ask(WORD_SERIF_FALLBACK_FACE),
        codePoint: codePointOf(BULLET),
      },
    ]);
  });

  it("records a character once however many times the layout measures it", () => {
    const faces = substitutingMetrics([symbols, ...everywhere], []);
    drawn(faces, "Meridian Symbols", BULLET);
    drawn(faces, "Meridian Symbols", BULLET);

    expect(faces.fallbackCharacters()).toHaveLength(1);
  });

  // The face's own page answers before anything else is asked, so nothing the
  // face can draw at all reaches another one.
  it("leaves a character the face's own page answers for where it is", () => {
    const faces = substitutingMetrics([symbols, ...everywhere], []);
    const found = foundIn(faces, "Meridian Symbols");
    if (found.advances.kind !== "advances") throw new Error(found.advances.reason);

    expect(found.advances.advanceFor(codePointOf(NO_BREAK_SPACE))).toBe(1229);
    expect(faces.fallbackCharacters()).toStrictEqual([]);
  });

  // The measured half of the rule: one character, two faces with no glyph for it,
  // and the kind of face that asked decides which face answers.
  it("draws a sans face's character out of Arial and a serif face's out of Times New Roman", () => {
    const faces = substitutingMetrics(
      [textFace("Meridian Sans", true), textFace("Meridian Serif", false), ...everywhere],
      [],
    );

    expect(drawn(faces, "Meridian Sans", BULLET)?.advance).toBe(726);
    expect(drawn(faces, "Meridian Serif", BULLET)?.advance).toBe(717);
  });

  // A face that says nothing about itself is not a sans one, which is what leaves
  // a symbol face on the serif half of the rule.
  it("draws a face that classifies itself as nothing out of Times New Roman", () => {
    const plain = buildFace({ name: "Meridian Plain", metrics: METRICS, characters: "AB" });
    const faces = substitutingMetrics([plain, ...everywhere], []);

    expect(drawn(faces, "Meridian Plain", BULLET)?.advance).toBe(717);
  });

  // The one of the four geometric bullets Unicode gives an emoji, which came out
  // of the emoji face though both Arial and Times New Roman carry the shape.
  it("takes a character the emoji face has before the face for its kind", () => {
    const faces = substitutingMetrics([textFace("Meridian Sans", true), ...everywhere], []);

    expect(drawn(faces, "Meridian Sans", SMALL_SQUARE)).toStrictEqual({
      metrics: EMOJI_METRICS,
      advance: 800,
    });
  });

  // Neither Arial nor Times New Roman carries a hyphen, and the face for the other
  // kind is never tried: a sans face's went to Cambria Math rather than to Times
  // New Roman.
  it("goes on to Cambria Math for a character neither face of its kind carries", () => {
    const faces = substitutingMetrics([textFace("Meridian Sans", true), ...everywhere], []);

    expect(drawn(faces, "Meridian Sans", HYPHEN)?.advance).toBe(680);
  });

  it("goes on to Segoe UI Symbol for a character Cambria Math has not got either", () => {
    const faces = substitutingMetrics([textFace("Meridian Sans", true), ...everywhere], []);

    expect(drawn(faces, "Meridian Sans", WORD_JOINER)?.advance).toBe(819);
  });

  it("reaches for nothing on a machine that has not got any of the faces", () => {
    const faces = substitutingMetrics([symbols], []);

    expect(foundIn(faces, "Meridian Symbols").elsewhere).toBeUndefined();
  });

  it("passes over a face the machine has not got and asks the next", () => {
    const faces = substitutingMetrics([textFace("Meridian Sans", true), serif, maths], []);

    expect(drawn(faces, "Meridian Sans", HYPHEN)?.advance).toBe(680);
  });

  // A bold run's fallback weight is unmeasured, and a width out of the regular
  // weight of the right face is nearer than no page at all.
  it("takes the plain style of the fallback face where the machine has only that", () => {
    const faces = substitutingMetrics([{ ...symbols, bold: true }, ...everywhere], []);

    expect(drawn(faces, "Meridian Symbols", BULLET, { bold: true })?.advance).toBe(717);
    expect(faces.fallbackCharacters()[0]?.used).toStrictEqual(ask(WORD_SERIF_FALLBACK_FACE));
  });
});
