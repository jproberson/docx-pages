import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { featuresIn, type DocumentCensus } from "./census.js";

// Ranking the gaps by what a corpus of real documents actually asks for.
//
// This joins the two halves that neither can answer alone: the sweep says which
// documents met a gap, and the census says how many documents had the feature to
// meet it with. A gap in two hundred documents means one thing where nine hundred
// have the construct and another where two hundred and ten do.
//
// **The numbers never enter the repository.** They are written under
// `samples/corpus/`, which is gitignored whole, along with everything else derived
// from documents that are not this project's to keep. All that is worth carrying
// away is the order this ranking put the gaps in, which says nothing about any
// document.

const SWEEP_PATH = process.env["DOCX_PAGES_CORPUS_REPORT"] ?? "samples/corpus/sweep.jsonl";
const CENSUS_PATH = process.env["DOCX_PAGES_CORPUS_CENSUS"] ?? "samples/corpus/census.jsonl";
const RANKING_PATH = process.env["DOCX_PAGES_CORPUS_GAPS"] ?? "samples/corpus/gaps.md";

// Which feature of a document a gap is about, where the census counts one. A gap
// nothing in the census answers for is ranked against the whole corpus, since
// every document could have met it.
const FEATURE_OF: Readonly<Record<string, string>> = {
  "merged-cells": "table-cells",
  "table-style-conditional-formatting": "tables",
  "text-columns": "multiple-columns",
  "column-break": "column-breaks",
  footnote: "notes",
  "undrawable-picture": "picture",
  "unknown-drawing": "drawings",
  "legacy-text-box": "legacy-drawings",
  "legacy-drawing": "legacy-drawings",
  "approximated-border": "paragraph-borders",
  // A path drawn point by point is a shape's, and an equation this cannot read is
  // one of the equations. Without these two the whole corpus is the denominator,
  // and each reads as a gap in a tenth of one per cent of it rather than one in the
  // few hundred documents that could have stated it at all.
  "custom-geometry": "shape",
  equation: "equations",
};

export type GapRank = {
  readonly kind: string;
  // Documents that met the gap, and documents that could have.
  readonly met: number;
  readonly could: number;
};

type Swept = { readonly id: string; readonly asks: readonly string[] };

export function rankGaps(
  swept: readonly Swept[],
  censuses: readonly DocumentCensus[],
): readonly GapRank[] {
  const features = new Map(featuresIn(censuses).map((each) => [each.feature, each.documents]));
  const met = new Map<string, number>();
  for (const each of swept) {
    for (const kind of new Set(each.asks)) met.set(kind, (met.get(kind) ?? 0) + 1);
  }

  return [...met.entries()]
    .map(([kind, documents]) => ({
      kind,
      met: documents,
      could: features.get(FEATURE_OF[kind] ?? "") ?? swept.length,
    }))
    .sort((one, other) => other.met - one.met || one.kind.localeCompare(other.kind));
}

const share = (count: number, of: number): string =>
  of === 0 ? "n/a" : `${((count / of) * 100).toFixed(1)}%`;

export function rankingReport(ranked: readonly GapRank[], documents: number): string {
  const lines = [
    "# What the corpus asks for and does not get",
    "",
    `Generated from a sweep of ${String(documents)} documents. **Not for the repository**:`,
    "everything here is derived from documents this project does not own. Only the order",
    "the gaps came out in is worth carrying away.",
    "",
    "| gap | documents | of those that could | share |",
    "| --- | --- | --- | --- |",
  ];

  for (const each of ranked) {
    lines.push(
      `| ${each.kind} | ${String(each.met)} | ${String(each.could)} | ${share(each.met, each.could)} |`,
    );
  }
  return lines.join("\n") + "\n";
}

const linesOf = (path: string): readonly unknown[] =>
  readFileSync(resolve(path), "utf8")
    .split("\n")
    .filter((line) => line.trim() !== "")
    .map((line): unknown => JSON.parse(line));

function sweptFrom(value: unknown): Swept {
  if (typeof value !== "object" || value === null) throw new Error("a sweep row is not an object");
  const row: Record<string, unknown> = { ...value };
  const asks = row["asks"];
  return {
    id: typeof row["id"] === "string" ? row["id"] : "",
    asks: Array.isArray(asks)
      ? asks.filter((each): each is string => typeof each === "string")
      : [],
  };
}

function censusFrom(value: unknown): DocumentCensus {
  if (typeof value !== "object" || value === null) throw new Error("a census row is not an object");
  const row: Record<string, unknown> = { ...value };
  const counts: Record<string, number> = {};
  for (const [key, each] of Object.entries({ ...(row["counts"] ?? {}) })) {
    if (typeof each === "number") counts[key] = each;
  }
  return {
    id: typeof row["id"] === "string" ? row["id"] : "",
    bytes: typeof row["bytes"] === "number" ? row["bytes"] : 0,
    counts,
  };
}

function main(): void {
  const swept = linesOf(SWEEP_PATH).map(sweptFrom);
  const censuses = linesOf(CENSUS_PATH).map(censusFrom);
  const ranked = rankGaps(swept, censuses);

  mkdirSync(dirname(resolve(RANKING_PATH)), { recursive: true });
  writeFileSync(resolve(RANKING_PATH), rankingReport(ranked, swept.length));
  process.stdout.write(`${rankingReport(ranked, swept.length)}\nWritten to ${RANKING_PATH}\n`);
  process.stdout.write(
    `\nThe order, which is all that is worth carrying away:\n${ranked.map((each, at) => `  ${String(at + 1)}. ${each.kind}`).join("\n")}\n`,
  );
}

// Compared against this module's own path: a guard naming the built `.js` never
// fires under tsx, which is how these are run, and the sweep then does nothing at
// all and says nothing about it.
if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main();
}
