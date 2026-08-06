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

/**
 * The face Word draws a character out of when the face it is written in has no
 * glyph for it and no place in its own page to alias it to.
 *
 * Measured on 2026-08-06 off Word's own pdf of the authored `unmapped-characters`
 * document: Word drew `U+2022` in Wingdings from Times New Roman, which the
 * document never mentions, at Times New Roman's own bullet of 717 units in 2048.
 * `U+2022` is above `0xFF` and so has no low byte to alias to, which is what tells
 * this from the `.notdef` a symbol face answers its own page with.
 */
export const WORD_CHARACTER_FALLBACK_FACE = "Times New Roman";

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

export type SubstitutingMetrics = {
  readonly metricsFor: MetricsResolver;
  // Every face that was stood in for, once each, in the order they were first
  // asked for. Read it after laying out: nothing is known until the layout has
  // asked for the faces it needs.
  substitutions(): readonly Substitution[];
  // Every character measured out of another face, once each face and character,
  // and read after laying out as the substitutions are.
  fallbackCharacters(): readonly FallbackCharacter[];
};

// Measuring text needs the widths of the face's own glyphs. A face whose metrics
// are known but whose widths are not can place an empty paragraph and nothing
// else, so it is no more use here than one that is missing.
const measures = (lookup: MetricsLookup): boolean =>
  lookup.kind === "found" && lookup.advances.kind === "advances";

const keyOf = (request: FaceRequest): string =>
  `${request.name.trim().toLowerCase()}|${request.bold ? "b" : ""}|${request.italic ? "i" : ""}`;

const same = (one: FaceRequest, other: FaceRequest): boolean => keyOf(one) === keyOf(other);

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
  fallbackNames: readonly string[],
): SubstitutingMetrics {
  const stood = new Map<string, Substitution>();
  const characters = new Map<string, FallbackCharacter>();

  const metricsFor: MetricsResolver = (request) => {
    let asked: MetricsLookup | null = null;

    for (const used of candidates(request, fallbackNames)) {
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

// Only a symbol face asks this. Everything such a face can draw at all it has
// already answered for, so a character it reports unmapped is one Word could not
// keep in the face either. What a text face does with a character it has no glyph
// for is a different question and an unmeasured one, so that one is left to refuse
// the document rather than drawn in a guess.
function throughAnotherFace(
  found: MetricsLookup,
  face: FaceRequest,
  supplied: readonly SuppliedFace[],
  characters: Map<string, FallbackCharacter>,
): MetricsLookup {
  if (found.kind !== "found" || found.advances.kind !== "advances") return found;
  if (!found.advances.symbolEncoded) return found;

  const reached = reachableFace(face, supplied);
  if (reached === null) return found;

  const { used } = reached;
  return {
    ...found,
    elsewhere: {
      metrics: reached.metrics,
      // Asked only of a character the face itself has no glyph for, which is what
      // makes this the place to say that one was met.
      advanceFor: (codePoint) => {
        const drawn = reached.advanceFor(codePoint);
        if (drawn === null) return null;
        const key = `${keyOf(face)}|${String(codePoint)}`;
        if (!characters.has(key)) characters.set(key, { face, used, codePoint });
        return drawn;
      },
    },
  };
}

type ReachedFace = {
  readonly used: FaceRequest;
  readonly metrics: FontMetrics;
  readonly advanceFor: GlyphAdvances;
};

// The style is worth trying for and not worth refusing the document over: what a
// bold run's fallback weighs is unmeasured, and the regular weight of the right
// face is nearer than no page at all.
function reachableFace(face: FaceRequest, supplied: readonly SuppliedFace[]): ReachedFace | null {
  const named = { ...face, name: WORD_CHARACTER_FALLBACK_FACE };
  for (const used of [named, { ...named, bold: false, italic: false }]) {
    const elsewhere = lookupFontMetrics(used, supplied);
    if (elsewhere.kind !== "found" || elsewhere.advances.kind !== "advances") continue;
    return { used, metrics: elsewhere.metrics, advanceFor: elsewhere.advances.advanceFor };
  }
  return null;
}
