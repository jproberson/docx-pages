import type { FaceShape } from "../docx/font-table.js";
import { lookupFontMetrics, type SuppliedFace } from "./font-metrics.js";
import type { MetricsResolver } from "./lines.js";
import {
  substitutingMetrics,
  type MissingGlyph,
  type SubstitutingMetrics,
} from "./substitution.js";
import { aliasedSymbolCharacter, isAliasedSymbolFace } from "./symbol-aliases.js";

/**
 * The faces that stand behind every name once being exact has already been given
 * up on: a pack of real faces, which twin the metrics of the ones documents
 * usually name, and one of them to answer for each shape a name can have.
 *
 * `twins` maps a lowercased document name to the pack face whose advances match
 * it glyph for glyph, so a page laid out on one breaks its lines where the named
 * face would have. The shape names answer for everything else, chosen by the
 * classification the document itself carries for the face (`readFaceShapes`).
 *
 * `lastResort` answers for a name nothing classifies. Word's own last resort is
 * Cambria, a serif, whatever the name suggested (see `WORD_FALLBACK_FACES`), so
 * a pack that wants to land where Word lands names Cambria's twin here rather
 * than a sans.
 */
export type FaceDefaults = {
  readonly faces: readonly SuppliedFace[];
  readonly twins: Readonly<Record<string, string>>;
  readonly sansSerif: string;
  readonly serif: string;
  readonly monospace: string;
  readonly lastResort: string;
};

export type BestEffortMetrics = SubstitutingMetrics & {
  missingGlyphs(): readonly MissingGlyph[];
};

const normalise = (name: string): string => name.trim().toLowerCase();

const sameRequest = (
  one: { readonly name: string; readonly bold: boolean; readonly italic: boolean },
  other: { readonly name: string; readonly bold: boolean; readonly italic: boolean },
): boolean =>
  normalise(one.name) === normalise(other.name) &&
  one.bold === other.bold &&
  one.italic === other.italic;

/**
 * Resolves faces so that every request is answered and every document lays out:
 * the caller's faces first, then the twin whose widths match the name asked for,
 * then the default of the name's own shape, then the sans default; and a
 * character that even the borrowing chain cannot draw is answered with a
 * missing-glyph box at the box's own advance rather than refusing the page.
 *
 * This is the resolution for a caller who would rather see the page than be
 * refused it, and it stays as loud as the exact one: every stand-in comes back
 * out of `substitutions()`, every borrowed character out of
 * `fallbackCharacters()`, every box out of `missingGlyphs()`, and the layout
 * folds all three into `unhonoured`. A page laid out over any of them is not the
 * page Word would draw, and says so.
 *
 * What it cannot promise is Word's own choice of stand-in: what Word reaches for
 * on a machine without a face is a question about that machine, so the choice
 * here is only disclosed, never claimed as Word's.
 */
export function bestEffortMetrics(
  supplied: readonly SuppliedFace[],
  defaults: FaceDefaults,
  shapes: ReadonlyMap<string, FaceShape> = new Map(),
): BestEffortMetrics {
  const all = [...supplied, ...defaults.faces];
  const boxes = new Map<string, MissingGlyph>();

  const shapeName = (shape: FaceShape | undefined): string | undefined => {
    if (shape === "serif") return defaults.serif;
    if (shape === "monospace") return defaults.monospace;
    return shape === "sans-serif" ? defaults.sansSerif : undefined;
  };

  const behind = (request: { readonly name: string }): readonly string[] => {
    const key = normalise(request.name);
    const names = [defaults.twins[key], shapeName(shapes.get(key)), defaults.lastResort];
    return [...new Set(names.filter((name): name is string => name !== undefined))];
  };

  const under = substitutingMetrics(all, behind);

  // The box of last resort, for the one case where the face that answered was
  // built by hand and does not say what its own box advances.
  const lastResort = lookupFontMetrics(
    { name: defaults.lastResort, bold: false, italic: false },
    defaults.faces,
  );
  const lastResortBox =
    lastResort.kind === "found" && lastResort.advances.kind === "advances"
      ? { metrics: lastResort.metrics, advance: lastResort.advances.notDefAdvance }
      : null;

  const metricsFor: MetricsResolver = Object.assign(
    (request: Parameters<MetricsResolver>[0]) => {
      const found = under.metricsFor(request);
      if (found.kind !== "found" || found.advances.kind !== "advances") return found;

      // A symbol face that was stood in for holds positions in its own page, not
      // text: the stand-in's letters would be the wrong glyphs wearing the right
      // widths. So each position is measured as its Unicode meaning, and one the
      // tables do not carry goes to the box rather than to the stand-in's letter.
      const symbolPage =
        isAliasedSymbolFace(request.name) &&
        under.substitutions().some((each) => sameRequest(each.requested, request));

      const advances = found.advances;
      const advanceFor = !symbolPage
        ? advances.advanceFor
        : (codePoint: number) => {
            const alias = aliasedSymbolCharacter(request.name, codePoint);
            return alias === null ? null : advances.advanceFor(alias.codePointAt(0) ?? 0);
          };

      const own = advances.notDefAdvance;
      const reached = found.elsewhere;
      return {
        ...found,
        advances: { ...advances, advanceFor },
        elsewhere: (codePoint: number) => {
          const alias = symbolPage ? aliasedSymbolCharacter(request.name, codePoint) : null;
          const meant = alias?.codePointAt(0) ?? codePoint;
          const borrowed = symbolPage && alias === null ? null : (reached?.(meant) ?? null);
          if (borrowed !== null) return borrowed;

          const box =
            own !== undefined
              ? { metrics: found.metrics, advance: own }
              : lastResortBox?.advance !== undefined
                ? { metrics: lastResortBox.metrics, advance: lastResortBox.advance }
                : null;
          if (box === null) return null;

          const key = `${normalise(request.name)}|${request.bold ? "b" : ""}${request.italic ? "i" : ""}|${String(codePoint)}`;
          if (!boxes.has(key)) boxes.set(key, { face: request, codePoint });
          return box;
        },
      };
    },
    { answersForUnresolved: true as const },
  );

  return {
    metricsFor,
    substitutions: () => under.substitutions(),
    fallbackCharacters: () => under.fallbackCharacters(),
    missingGlyphs: () => [...boxes.values()],
  };
}
