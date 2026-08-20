import { createHash } from "node:crypto";
import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { join, resolve } from "node:path";

import {
  layOutDocument,
  openDocx,
  readUnhonoured,
  substitutingMetrics,
  isDocxPagesError,
  WORD_FALLBACK_FACES,
  type SuppliedFace,
  readFaceAlternatives,
} from "@docx-pages/core";

import { corpusFaces } from "./faces.js";

// A sweep of a corpus of real documents: what each one asks for, and whether this
// project can lay it out at all.
//
// The corpus is other people's work and never enters the repository. It is found
// through `DOCX_PAGES_CORPUS`, exactly as the seven reference one-pagers are found
// through their own manifest, and with no corpus configured the sweep says it has
// nothing to run rather than failing.
//
// **Nothing the sweep writes names a document.** One is known by the first twelve
// characters of the hash of its bytes, which is enough to follow the same document
// between two runs and tells a reader nothing about what it was. That is what makes
// a report safe to keep, to diff, and to quote.

export const CORPUS_DIRECTORY = process.env["DOCX_PAGES_CORPUS"] ?? null;

// How a document ended up. A blocker is this project refusing the document in the
// way it is designed to, which is a different thing from throwing: the first is a
// gap it knows about, the second is a gap it does not.
export type SweepOutcome =
  | { readonly kind: "laid-out" }
  | { readonly kind: "blocked"; readonly detail: string }
  | { readonly kind: "refused"; readonly detail: string }
  | { readonly kind: "threw"; readonly detail: string };

export type SweptDocument = {
  readonly id: string;
  readonly bytes: number;
  // What the document asks for and does not get, read from the package alone. This
  // needs no fonts and no layout, so it answers even for a document that will not
  // lay out at all.
  readonly asks: readonly string[];
  // How many faces the document asked for that another one answered for. This is a
  // fact about the machine's fonts rather than about the document, so it is counted
  // apart from `asks` and never ranked beside it.
  readonly facesStoodIn: number;
  readonly outcome: SweepOutcome;
  readonly pages: number | null;
  readonly millis: number;
};

export const identityOf = (bytes: Uint8Array): string =>
  createHash("sha256").update(bytes).digest("hex").slice(0, 12);

// A rule that can only reach the documents stating one thing is judged by scoring
// those documents rather than the corpus: the rest come back with the row they
// already had, so an hour of drawing answers a question a few minutes answers.
// `DOCX_PAGES_CORPUS_ONLY` names the ids to read, as a file of them or a list, and
// with nothing named every document is read.
const namedOnly = (process.env["DOCX_PAGES_CORPUS_ONLY"] ?? "").trim();

export const idsAsked = (): ReadonlySet<string> | null => {
  if (namedOnly === "") return null;
  const text = existsSync(namedOnly) ? readFileSync(namedOnly, "utf8") : namedOnly;
  return new Set(text.split(/[\s,]+/).filter((each) => each !== ""));
};

// Word writes a lock file beside an open document, and both macOS and Windows
// leave their own litter about. Neither is a document.
const isDocument = (name: string): boolean =>
  name.toLowerCase().endsWith(".docx") && !name.startsWith("~$") && !name.startsWith(".");

export function documentsIn(directory: string): readonly string[] {
  const found: string[] = [];

  const visit = (at: string): void => {
    for (const entry of readdirSync(at, { withFileTypes: true })) {
      const path = join(at, entry.name);
      if (entry.isDirectory()) {
        visit(path);
        continue;
      }
      if (entry.isFile() && isDocument(entry.name)) found.push(path);
    }
  };

  visit(resolve(directory));
  return found.sort();
}

const errnoOf = (value: unknown): string | null => {
  if (!(value instanceof Error) || !("code" in value)) return null;
  return typeof value.code === "string" ? value.code : null;
};

// What a failure is allowed to say. A file the sweep could not read throws a
// message holding the whole path, and a library throws one holding the text it
// choked on, so an errno is reported as itself, anything quoted is emptied, and a
// path takes the rest of its line with it, since a document's name holds spaces as
// often as not. Nothing a report can be quoted from may name a document.
export const described = (value: unknown): string => {
  if (isDocxPagesError(value)) return value.code;
  const errno = errnoOf(value);
  if (errno !== null) return errno;
  return (value instanceof Error ? value.message : String(value))
    .replace(/'[^']*'/g, "'...'")
    .replace(/"[^"]*"/g, '"..."')
    .replace(/\S*[/\\].*/g, "...")
    .slice(0, 200);
};

/**
 * One document, read and laid out, with everything it can do to fail caught.
 *
 * The two halves are separate on purpose. What a document **asks for** is read
 * from the package, so it answers even where the layout gives up; whether it
 * **lays out** needs faces, and is asked with the substituting resolver so that a
 * face this machine has not got costs a note rather than the whole document.
 */
export function sweepDocument(bytes: Uint8Array, faces: readonly SuppliedFace[]): SweptDocument {
  const startedAt = performance.now();
  const asks = asked(bytes);
  const placed = laidOut(bytes, faces);

  return {
    id: identityOf(bytes),
    bytes: bytes.length,
    asks,
    ...placed,
    millis: Math.round(performance.now() - startedAt),
  };
}

// A package this project cannot even open says nothing about what it asked for,
// and the layout below is what will report why.
function asked(bytes: Uint8Array): readonly string[] {
  try {
    return readUnhonoured(openDocx(bytes)).map((each) => each.kind);
  } catch {
    return [];
  }
}

type Placement = {
  readonly outcome: SweepOutcome;
  readonly pages: number | null;
  readonly facesStoodIn: number;
};

function laidOut(bytes: Uint8Array, faces: readonly SuppliedFace[]): Placement {
  // **The document's own alternatives are part of how a face is stood in**, so the
  // resolver is made after the package is opened. A package that will not open at
  // all leaves it holding none, which is what it held before this was read.
  let measuring = substitutingMetrics(faces, WORD_FALLBACK_FACES);
  try {
    const pkg = openDocx(bytes);
    measuring = substitutingMetrics(faces, WORD_FALLBACK_FACES, readFaceAlternatives(pkg));
    const laid = layOutDocument(pkg, measuring);
    const facesStoodIn = measuring.substitutions().length;
    if (laid.kind !== "laid-out") {
      return { outcome: { kind: "blocked", detail: laid.blocker.kind }, pages: null, facesStoodIn };
    }
    return { outcome: { kind: "laid-out" }, pages: laid.pages.length, facesStoodIn };
  } catch (thrown) {
    // A located error is this project saying no in its own words; anything else is
    // a document doing something nothing here anticipated, which is the sweep's
    // most valuable finding.
    const kind = isDocxPagesError(thrown) ? "refused" : "threw";
    return {
      outcome: { kind, detail: described(thrown) },
      pages: null,
      facesStoodIn: measuring.substitutions().length,
    };
  }
}

export type SweepProgress = (done: number, total: number) => void;

export function sweepCorpus(
  directory: string,
  onProgress?: SweepProgress,
): readonly SweptDocument[] {
  const paths = documentsIn(directory);
  const swept: SweptDocument[] = [];
  // Read once for the whole sweep: every face is a file off the disk, and reading
  // them again for each of a thousand documents is the sweep's whole cost.
  const faces = corpusFaces();

  // The same document under two names is one document. A corpus gathered from the
  // wild is full of them: the 966 files first swept here are 718 documents, and one
  // of them is in there 24 times. Counted by the file, that document's features
  // weigh 24 times what another's do, and every share ever read off this sweep was
  // wrong by however many copies it happened to be looking at.
  const already = new Set<string>();

  for (const [at, path] of paths.entries()) {
    // A document too big to read is still a fact about the corpus, so even reading
    // one is allowed to fail without ending the sweep.
    let bytes: Uint8Array;
    try {
      bytes = new Uint8Array(readFileSync(path));
    } catch (thrown) {
      swept.push({
        id: identityOf(new TextEncoder().encode(path)),
        bytes: statSizeOf(path),
        asks: [],
        facesStoodIn: 0,
        outcome: { kind: "threw", detail: described(thrown) },
        pages: null,
        millis: 0,
      });
      continue;
    }

    const id = identityOf(bytes);
    if (already.has(id)) continue;
    already.add(id);

    swept.push(sweepDocument(bytes, faces));
    onProgress?.(at + 1, paths.length);
  }

  return swept;
}

const statSizeOf = (path: string): number => {
  try {
    return statSync(path).size;
  } catch {
    return 0;
  }
};

// How often each kind was met, and in how many documents. Both are worth having:
// a kind met three hundred times in one document is not the same problem as one
// met once in three hundred.
export type KindTally = {
  readonly kind: string;
  readonly documents: number;
  readonly occurrences: number;
};

export type SweepSummary = {
  readonly documents: number;
  readonly laidOut: number;
  readonly blocked: number;
  readonly refused: number;
  readonly threw: number;
  // Every kind, most documents first, which is the ranking the gap list is meant
  // to be built from.
  readonly kinds: readonly KindTally[];
  // What stopped a document, most common first.
  readonly failures: readonly KindTally[];
  // How many documents were drawn in a face they did not ask for, which says how
  // much this machine's fonts are colouring the rest.
  readonly facesStoodIn: number;
};

export function summaryOf(swept: readonly SweptDocument[]): SweepSummary {
  const kinds = tallied(swept.map((each) => each.asks));
  const failures = tallied(
    swept.flatMap((each) =>
      each.outcome.kind === "laid-out" ? [] : [[`${each.outcome.kind}: ${each.outcome.detail}`]],
    ),
  );
  const counting = (kind: SweepOutcome["kind"]): number =>
    swept.filter((each) => each.outcome.kind === kind).length;

  return {
    documents: swept.length,
    laidOut: counting("laid-out"),
    blocked: counting("blocked"),
    refused: counting("refused"),
    threw: counting("threw"),
    kinds,
    failures,
    facesStoodIn: swept.filter((each) => each.facesStoodIn > 0).length,
  };
}

function tallied(perDocument: readonly (readonly string[])[]): readonly KindTally[] {
  const documents = new Map<string, number>();
  const occurrences = new Map<string, number>();

  for (const each of perDocument) {
    for (const kind of each) occurrences.set(kind, (occurrences.get(kind) ?? 0) + 1);
    for (const kind of new Set(each)) documents.set(kind, (documents.get(kind) ?? 0) + 1);
  }

  return [...documents.entries()]
    .map(([kind, met]) => ({ kind, documents: met, occurrences: occurrences.get(kind) ?? 0 }))
    .sort((one, other) => other.documents - one.documents || one.kind.localeCompare(other.kind));
}

// What changed between two sweeps, which is what says whether a layout change was
// worth making. A document is followed by its own identity, so one added to or
// taken out of the corpus does not read as a change to the rest.
export type SweepChange = {
  readonly id: string;
  readonly was: string;
  readonly now: string;
};

export function changesBetween(
  before: readonly SweptDocument[],
  after: readonly SweptDocument[],
): readonly SweepChange[] {
  const earlier = new Map(before.map((each) => [each.id, each]));
  const changes: SweepChange[] = [];

  for (const each of after) {
    const was = earlier.get(each.id);
    if (was === undefined) continue;
    const before = stateOf(was);
    const now = stateOf(each);
    if (before !== now) changes.push({ id: each.id, was: before, now });
  }

  return changes;
}

// What a run says about one document, in one line, so that two runs differ exactly
// where the layout did. Timing is left out: it changes every run and means nothing.
const stateOf = (each: SweptDocument): string =>
  `${each.outcome.kind === "laid-out" ? `${String(each.pages)} pages` : `${each.outcome.kind}: ${each.outcome.detail}`} | ${[...each.asks].sort().join(",")}`;
