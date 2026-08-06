import {
  lookupFontMetrics,
  type FaceRequest,
  type FontMetrics,
  type GlyphAdvances,
  type MetricsLookup,
  type SuppliedFace,
} from "./font-metrics.js";
import type { MetricsResolver } from "./lines.js";

// A face the document asked for that another one answered for. Word does this
// silently; this says so, because a document drawn in a face it was not written in
// is not the document Word would draw, and every line in it may break elsewhere.
export type Substitution = {
  readonly requested: FaceRequest;
  // What stood in for it, which may differ from the request in style as well as in
  // name: a bold run drawn in the regular weight is still a substitution.
  readonly used: FaceRequest;
};

/**
 * What Word itself reaches for when it cannot resolve the name a document asks
 * for. Six names nothing supplies were put to Word on 2026-08-05 - `Nonesuch
 * Sans`, `Nonesuch Serif`, `Zapfino Nonesuch`, `Verdana;Arial`, `Nonesuch
 * Sans;Arial` and `OpenSymbol;Arial Unicode MS` - and its own pdf came back in
 * Cambria for every one of them, whatever the name suggested about the face.
 *
 * Word knows a table of its own beside this, which resolves a name it recognises
 * somewhere else entirely: `DejaVu Sans` comes back Verdana, `Liberation
 * Serif;Times New Roman` comes back Times New Roman, and anything opening `Open
 * Sans;` comes back Segoe UI even though Open Sans itself is installed. That
 * table is Word's own and is not reproduced here, so a document naming one of
 * those is laid out in the last resort rather than in what Word chose.
 */
export const WORD_FALLBACK_FACES: readonly string[] = ["Cambria"];

// The emoji face, and the face for each kind of text face. Named one at a time
// because a caller reaches for them one at a time: `fallback.ts` knows where Times
// New Roman lives and that it is the one face read out of two files.
export const WORD_EMOJI_FACE = "Apple Color Emoji";
export const WORD_SANS_FALLBACK_FACE = "Arial";
export const WORD_SERIF_FALLBACK_FACE = "Times New Roman";

// What both kinds go on to for a character neither Arial nor Times New Roman
// carries, in the order measured.
const AFTERWARDS: readonly string[] = ["Cambria Math", "Segoe UI Symbol"];

/**
 * Every face Word draws a character out of when the face it is written in has no
 * glyph for it and no place in its own page to alias it to, which is what a caller
 * has to supply before any of this can happen at all. Which one answers is not one
 * name: it turns on the kind of face that asked and then on the character.
 *
 * Measured on 2026-08-06 off Word's own pdf of the authored `unmapped-characters`
 * and `unmapped-in-a-text-face` documents, twelve cases over six stated faces, each
 * on a page of its own so that a face the pdf names there is one Word reached for
 * itself. Every advance below is the borrowed face's own, read off where Word put
 * the letter after the character rather than off the pdf's rounded `/Widths`.
 *
 * - **The emoji face answers first for a character it has one for.** `U+25AA` in
 *   Cambria came out of Apple Color Emoji a whole em wide, though both Arial and
 *   Times New Roman carry the same shape at 726 units in 2048. It is the one of
 *   the four geometric bullets Unicode gives an emoji, and the emoji face maps
 *   little else, so the face's own map is the list of characters this covers.
 * - **Then the kind of face that asked decides.** One character in four faces that
 *   have no glyph for it: the two sans faces were answered by Arial and the two
 *   serif faces by Times New Roman. The bullet Wingdings could not alias went to
 *   Times New Roman too, so a symbol face falls on the serif half of this rule
 *   rather than on a last resort of Word's.
 * - **Then the character decides, where neither of those carries it.** A hyphen in
 *   Arial and a maths italic letter in Cambria both went to Cambria Math, at 680
 *   and 1141 units, against the 819 and 1141 of Segoe UI Symbol; the word joiner,
 *   which is not meant to be drawn at all, went to Segoe UI Symbol at no width.
 *
 * So which face answers is a question about the machine: four of these ship with
 * Word and the fifth is macOS's own emoji face, and a machine holding a different
 * set would answer differently. Nothing in a document says which face Word will
 * name, which is why `lookupFontMetrics` never offers a fallback of its own and
 * only a resolver holding every face the machine has can.
 */
export const WORD_CHARACTER_FALLBACK_FACES: readonly string[] = [
  WORD_EMOJI_FACE,
  WORD_SANS_FALLBACK_FACE,
  WORD_SERIF_FALLBACK_FACE,
  ...AFTERWARDS,
];

// What a face of the given kind reaches for, in the order worth reaching. Only one
// of the two the kind decides is in either list: a sans face met with the hyphen
// it has no glyph for went to Cambria Math at 680 units rather than to Times New
// Roman's 682, so the face for the other kind is never tried.
const reachedBy = (sansSerif: boolean): readonly string[] => [
  WORD_EMOJI_FACE,
  sansSerif ? WORD_SANS_FALLBACK_FACE : WORD_SERIF_FALLBACK_FACE,
  ...AFTERWARDS,
];

// A character measured out of a face the document never asked for. Word does this
// silently too, and it is not the same thing as standing a whole face in: the
// character takes the room Word gave it, so nothing moves, and only the glyph
// drawn in that room is anyone's guess.
export type FallbackCharacter = {
  // The face that was measured with rather than the one the document asked for: a
  // run whose face was stood in for asks this of the stand-in.
  readonly face: FaceRequest;
  readonly used: FaceRequest;
  readonly codePoint: number;
};

// A character nothing on hand could draw, which the best-effort resolver answers
// with the missing-glyph box rather than refusing the document over.
export type MissingGlyph = {
  readonly face: FaceRequest;
  readonly codePoint: number;
};

export type SubstitutingMetrics = {
  readonly metricsFor: MetricsResolver;
  // Every face that was stood in for, once each, in the order they were first
  // asked for. Read it after laying out: nothing is known until the layout has
  // asked for the faces it needs.
  substitutions(): readonly Substitution[];
  // Every character measured out of another face, once each face and character,
  // and read after laying out as the substitutions are.
  fallbackCharacters(): readonly FallbackCharacter[];
  // Every character drawn as a missing-glyph box, which only the best-effort
  // resolver ever answers with: this one never does, so it never says.
  missingGlyphs?(): readonly MissingGlyph[];
};

// Measuring text needs the widths of the face's own glyphs. A face whose metrics
// are known but whose widths are not can place an empty paragraph and nothing
// else, so it is no more use here than one that is missing.
const measures = (lookup: MetricsLookup): boolean =>
  lookup.kind === "found" && lookup.advances.kind === "advances";

const keyOf = (request: FaceRequest): string =>
  `${request.name.trim().toLowerCase()}|${request.bold ? "b" : ""}|${request.italic ? "i" : ""}`;

const same = (one: FaceRequest, other: FaceRequest): boolean => keyOf(one) === keyOf(other);

// Which names stand behind a request. A fixed list serves a caller who knows the
// machine; the best-effort resolver chooses per name instead, since the right
// stand-in for a serif name is not the right one for a sans name.
export type FallbackNames = readonly string[] | ((request: FaceRequest) => readonly string[]);

const namesFor = (fallbackNames: FallbackNames, request: FaceRequest): readonly string[] =>
  typeof fallbackNames === "function" ? fallbackNames(request) : fallbackNames;

// What to try, in the order worth trying it. The face the document asked for comes
// first; then the same face without the style, since a run drawn in its own family
// at the wrong weight is nearer than the right weight of a stranger; then the
// fallbacks, each in the style asked for before the style is given up on.
function candidates(
  request: FaceRequest,
  fallbackNames: readonly string[],
): readonly FaceRequest[] {
  const plain = { ...request, bold: false, italic: false };
  return [
    request,
    plain,
    ...fallbackNames.map((name) => ({ ...request, name })),
    ...fallbackNames.map((name) => ({ ...plain, name })),
  ];
}

/**
 * Resolves faces as `lookupFontMetrics` does, and where a document asks for one
 * that cannot be measured, stands the nearest usable face in its place and records
 * that it did.
 *
 * The alternative is refusing the document outright, which is right when being
 * exact is the whole point and wrong when a reader would rather see the page. What
 * makes this safe is that it is never quiet: whatever was stood in for comes back
 * out of `substitutions()` for the caller to say so, and a page drawn on the back
 * of one is no longer the page Word would draw.
 *
 * A face nothing can answer for is left to fail as it would have, so the caller
 * still gets the blocker naming it rather than a page laid out in nothing.
 */
export function substitutingMetrics(
  supplied: readonly SuppliedFace[],
  fallbackNames: FallbackNames,
): SubstitutingMetrics {
  const stood = new Map<string, Substitution>();
  const characters = new Map<string, FallbackCharacter>();

  const metricsFor: MetricsResolver = (request) => {
    let asked: MetricsLookup | null = null;

    for (const used of candidates(request, namesFor(fallbackNames, request))) {
      const found = lookupFontMetrics(used, supplied);
      asked ??= found;
      if (!measures(found)) continue;

      if (!same(used, request)) {
        const key = keyOf(request);
        if (!stood.has(key)) stood.set(key, { requested: request, used });
      }
      return throughAnotherFace(found, used, supplied, characters);
    }

    return asked ?? lookupFontMetrics(request, supplied);
  };

  return {
    metricsFor,
    substitutions: () => [...stood.values()],
    fallbackCharacters: () => [...characters.values()],
  };
}

// Hangs the faces the stated one borrows from off the lookup, so that the line
// measurer can ask them for a character the stated face has no glyph for. Both a
// symbol face and a text face reach: everything a symbol face can draw at all it
// has already answered for out of its own page, and a text face has nothing of the
// kind, so a character either of them reports unmapped is one Word could not keep
// in the face either.
function throughAnotherFace(
  found: MetricsLookup,
  face: FaceRequest,
  supplied: readonly SuppliedFace[],
  characters: Map<string, FallbackCharacter>,
): MetricsLookup {
  if (found.kind !== "found" || found.advances.kind !== "advances") return found;

  const reached = reachableFaces(face, supplied);
  if (reached.length === 0) return found;

  return {
    ...found,
    // Asked only of a character the face itself has no glyph for, which is what
    // makes this the place to say that one was met.
    elsewhere: (codePoint) => {
      for (const each of reached) {
        const advance = each.advanceFor(codePoint);
        if (advance === null) continue;

        const key = `${keyOf(face)}|${String(codePoint)}`;
        if (!characters.has(key)) characters.set(key, { face, used: each.used, codePoint });
        return { metrics: each.metrics, advance };
      }
      return null;
    },
  };
}

type ReachedFace = {
  readonly used: FaceRequest;
  readonly metrics: FontMetrics;
  readonly advanceFor: GlyphAdvances;
};

// A face measures text only where the machine has the very style asked for, since
// a near miss comes back with the wrong style's widths and no advances at all. So
// the exact entry is the one that says what kind of face this is.
const suppliedAs = (face: FaceRequest, supplied: readonly SuppliedFace[]): SuppliedFace | null =>
  supplied.find((each) => same(each, face)) ?? null;

// The faces this one borrows from, in order, and only those the machine has: a
// missing one is passed over rather than refused, which leaves a machine with none
// of them exactly where it was.
//
// The style is worth trying for and not worth refusing the document over: what a
// bold run's fallback weighs is unmeasured, and the regular weight of the right
// face is nearer than no page at all.
function reachableFaces(
  face: FaceRequest,
  supplied: readonly SuppliedFace[],
): readonly ReachedFace[] {
  const kind = reachedBy(suppliedAs(face, supplied)?.sansSerif ?? false);

  return kind.flatMap((name) => {
    const named = { ...face, name };
    for (const used of [named, { ...named, bold: false, italic: false }]) {
      const elsewhere = lookupFontMetrics(used, supplied);
      if (elsewhere.kind !== "found" || elsewhere.advances.kind !== "advances") continue;
      if (same(used, face)) continue;
      return [{ used, metrics: elsewhere.metrics, advanceFor: elsewhere.advances.advanceFor }];
    }
    return [];
  });
}
