import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import {
  censusOf,
  coveringDocuments,
  distinctDocuments,
  featuresIn,
  type DocumentCensus,
} from "./census.js";
import { documentsIn, CORPUS_DIRECTORY } from "./sweep.js";

// Taking a census of a corpus, and choosing the smallest sample of it worth
// looking at.
//
// This is the cheap half of a sweep: no fonts and no layout, so it answers in
// minutes what the layout sweep answers in an hour, and it answers a question the
// layout sweep cannot ask at all.

const CENSUS_PATH = process.env["DOCX_PAGES_CORPUS_CENSUS"] ?? "samples/corpus/census.jsonl";

// How many documents have to hold a feature before the sample is taken to cover
// it. One document holding a feature is one document's worth of evidence about it,
// and where a corpus has only one that is all there is.
const COVER_EACH = 3;

const share = (count: number, of: number): string =>
  of === 0 ? "0%" : `${((count / of) * 100).toFixed(1)}%`;

export function reportOf(censuses: readonly DocumentCensus[]): string {
  // A corpus holds the same bytes under more than one name, and everything counted
  // below is over documents rather than files, so the share is over documents too.
  const documents = distinctDocuments(censuses).length;
  const lines = [`${String(documents)} documents`, "", "what they hold:"];

  for (const each of featuresIn(censuses)) {
    const met = `${String(each.documents).padStart(5)} (${share(each.documents, documents).padStart(6)})`;
    lines.push(`  ${met}  ${each.feature.padEnd(28)} ${String(each.occurrences)} times`);
  }

  const covering = coveringDocuments(censuses, COVER_EACH);
  lines.push(
    "",
    `${String(covering.length)} documents cover every feature at least ${String(COVER_EACH)} times over:`,
    ...covering.map((each) => `  ${each.id}  ${String(each.bytes)} bytes`),
  );
  return lines.join("\n");
}

function main(): void {
  if (CORPUS_DIRECTORY === null) {
    process.stdout.write(
      "No corpus configured: set DOCX_PAGES_CORPUS to a directory of .docx files.\n",
    );
    return;
  }

  const paths = documentsIn(CORPUS_DIRECTORY);
  process.stdout.write(`Taking a census of ${String(paths.length)} documents\n`);

  const censuses: DocumentCensus[] = [];
  for (const [at, path] of paths.entries()) {
    censuses.push(censusOf(new Uint8Array(readFileSync(path))));
    if ((at + 1) % 100 === 0) process.stdout.write(`  ${String(at + 1)}/${String(paths.length)}\n`);
  }

  mkdirSync(dirname(resolve(CENSUS_PATH)), { recursive: true });
  writeFileSync(
    resolve(CENSUS_PATH),
    censuses.map((each) => JSON.stringify(each)).join("\n") + "\n",
  );
  process.stdout.write(`\n${reportOf(censuses)}\n\nWritten to ${CENSUS_PATH}\n`);
}

// Compared against this module's own path: a guard naming the built `.js` never
// fires under tsx, which is how these are run, and the sweep then does nothing at
// all and says nothing about it.
if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main();
}
