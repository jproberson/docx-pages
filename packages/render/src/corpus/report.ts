import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";

import { corpusFaces } from "./faces.js";
import {
  changesBetween,
  summaryOf,
  sweepCorpus,
  CORPUS_DIRECTORY,
  type SweepSummary,
  type SweptDocument,
} from "./sweep.js";

// Running a corpus sweep and saying what it found.
//
// The report is written where the reference documents live, which is gitignored
// whole, so a run leaves nothing behind that could be committed by accident. It
// holds no filenames: see the note at the top of `sweep.ts`.

const REPORT_PATH = process.env["DOCX_PAGES_CORPUS_REPORT"] ?? "samples/corpus/sweep.jsonl";

export function writeSweep(swept: readonly SweptDocument[], path: string): void {
  mkdirSync(dirname(resolve(path)), { recursive: true });
  writeFileSync(resolve(path), swept.map((each) => JSON.stringify(each)).join("\n") + "\n");
}

export function readSweep(path: string): readonly SweptDocument[] {
  return readFileSync(resolve(path), "utf8")
    .split("\n")
    .filter((line) => line.trim() !== "")
    .map((line, at) => sweptFrom(JSON.parse(line), at));
}

// An earlier run's report, read back to be compared against. This tool wrote it,
// so what is checked is that the file is a sweep at all: a report out of another
// version answers with the line that does not fit rather than with a comparison
// nobody can trust.
function sweptFrom(value: unknown, at: number): SweptDocument {
  const where = `line ${String(at + 1)}`;
  const row = objectAt(value, where);
  const outcome = objectAt(row["outcome"], `${where} outcome`);
  const kind = outcome["kind"];
  if (kind !== "laid-out" && kind !== "blocked" && kind !== "refused" && kind !== "threw") {
    throw new Error(`${where} has no outcome this sweep knows`);
  }
  const pages = row["pages"];

  return {
    id: stringAt(row, "id", where),
    bytes: numberAt(row, "bytes", where),
    asks: asksAt(row, where),
    facesStoodIn: numberAt(row, "facesStoodIn", where),
    outcome: kind === "laid-out" ? { kind } : { kind, detail: stringAt(outcome, "detail", where) },
    pages: typeof pages === "number" ? pages : null,
    millis: numberAt(row, "millis", where),
  };
}

function objectAt(value: unknown, where: string): Record<string, unknown> {
  if (typeof value !== "object" || value === null) throw new Error(`${where} is not an object`);
  return { ...value };
}

function stringAt(row: Record<string, unknown>, name: string, where: string): string {
  const value = row[name];
  if (typeof value !== "string") throw new Error(`${where} has no ${name}`);
  return value;
}

function numberAt(row: Record<string, unknown>, name: string, where: string): number {
  const value = row[name];
  if (typeof value !== "number") throw new Error(`${where} has no ${name}`);
  return value;
}

function asksAt(row: Record<string, unknown>, where: string): readonly string[] {
  const value = row["asks"];
  if (!Array.isArray(value)) throw new Error(`${where} has no asks`);
  return value.map((each) => {
    if (typeof each !== "string") throw new Error(`${where} asks for something unnamed`);
    return each;
  });
}

const share = (count: number, of: number): string =>
  of === 0 ? "0%" : `${((count / of) * 100).toFixed(1)}%`;

export function reportOf(summary: SweepSummary): string {
  const lines = [
    `${String(summary.documents)} documents`,
    `  laid out  ${String(summary.laidOut)} (${share(summary.laidOut, summary.documents)})`,
    `  blocked   ${String(summary.blocked)}`,
    `  refused   ${String(summary.refused)}`,
    `  threw     ${String(summary.threw)}`,
  ];

  if (summary.failures.length > 0) {
    lines.push("", "what stopped a document, most common first:");
    for (const each of summary.failures) {
      lines.push(`  ${String(each.documents).padStart(5)}  ${each.kind}`);
    }
  }

  lines.push(
    "",
    `faces stood in for: ${String(summary.facesStoodIn)} documents needed at least one`,
    "",
    "what a document asked for and did not get, by documents met in:",
  );
  for (const each of summary.kinds) {
    const met = `${String(each.documents).padStart(5)} (${share(each.documents, summary.documents).padStart(6)})`;
    lines.push(`  ${met}  ${each.kind.padEnd(38)} ${String(each.occurrences)} times`);
  }

  return lines.join("\n");
}

function main(): void {
  if (CORPUS_DIRECTORY === null) {
    process.stdout.write(
      "No corpus configured: set DOCX_PAGES_CORPUS to a directory of .docx files.\n",
    );
    return;
  }

  const against = process.argv[2] ?? null;
  // A sweep measured with too few faces counts font trouble where there is none,
  // so the set it found is part of the result rather than a detail of the run.
  process.stdout.write(`Sweeping ${CORPUS_DIRECTORY} with ${String(corpusFaces().length)} faces\n`);

  let reported = 0;
  const swept = sweepCorpus(CORPUS_DIRECTORY, (done, total) => {
    // A sweep of a thousand documents is long enough to want to know it is alive,
    // and noisy enough that every one of them is too much to say.
    if (done !== total && done - reported < 50) return;
    reported = done;
    process.stdout.write(`  ${String(done)}/${String(total)}\n`);
  });

  writeSweep(swept, REPORT_PATH);
  process.stdout.write(`\n${reportOf(summaryOf(swept))}\n\nWritten to ${REPORT_PATH}\n`);

  if (against === null) return;
  const changes = changesBetween(readSweep(against), swept);
  process.stdout.write(`\n${String(changes.length)} documents changed since ${against}\n`);
  for (const each of changes.slice(0, 40)) {
    process.stdout.write(`  ${each.id}\n    was ${each.was}\n    now ${each.now}\n`);
  }
}

if (process.argv[1]?.endsWith("report.js") === true) main();
