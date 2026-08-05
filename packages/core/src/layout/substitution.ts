import {
  lookupFontMetrics,
  type FaceRequest,
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

export type SubstitutingMetrics = {
  readonly metricsFor: MetricsResolver;
  // Every face that was stood in for, once each, in the order they were first
  // asked for. Read it after laying out: nothing is known until the layout has
  // asked for the faces it needs.
  substitutions(): readonly Substitution[];
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
      return found;
    }

    return asked ?? lookupFontMetrics(request, supplied);
  };

  return { metricsFor, substitutions: () => [...stood.values()] };
}
