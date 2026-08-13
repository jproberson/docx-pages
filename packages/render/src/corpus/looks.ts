import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import {
  looksOf,
  type DrawOurs,
  shareOf,
  shareOfLooks,
  workspaceIn,
  worstPageOf,
  type Looks,
  type Workspace,
} from "../raster/compare.js";
import { canDraw } from "../raster/draw.js";
import { ourWrittenPages } from "../raster/written.js";
import { renderedPath } from "./render.js";
import { CORPUS_DIRECTORY, documentsIn, identityOf, idsAsked } from "./sweep.js";

// How different every document in a corpus looks from Word's own drawing of it.
//
// **The score beside this one cannot see a page.** `agreement.ts` counts the
// baselines and lefts of lines we laid out, over the lines Word drew the same
// text for. So it is blind to paint, pictures, borders, shapes, crops and what
// covers what, and blindest of all to content we draw nowhere at all, which never
// enters its denominator and therefore costs nothing. On 2026-08-10 a page
// scoring 35 of its 35 lines was looked at for the first time and was wrong five
// ways; the bug behind two of the three pages looked at that day was one line of
// header resolution, and fixing it took the line-perfect clean documents from 122
// to 353. **No number this project had ever computed could see it.**
//
// So this asks the other question, and it asks it of the page rather than of the
// layout: draw ours, draw Word's, and count how much of the page does not match.
// **Keep both.** The raster says which page is wrong and where on it; the lines
// and `inspect.ts` say by how much and in which direction, which is what names
// the rule. Neither replaces the other.
//
// **Nothing this writes names a document.** A row says the first twelve
// characters of the hash of the document's bytes and nothing more, and no number
// out of it may be committed.

const REPORT_PATH = process.env["DOCX_PAGES_CORPUS_LOOKS"] ?? "samples/corpus/looks.jsonl";

const AGREEMENT_PATH =
  process.env["DOCX_PAGES_CORPUS_AGREEMENT"] ?? "samples/corpus/agreement.jsonl";

const DIRECTORY = process.env["DOCX_PAGES_RASTER"] ?? "samples/corpus/raster";

// How many documents are drawn at once. Each wants a browser of its own and a
// browser is most of a second before it has drawn anything, so the run is bound
// by how many can be started rather than by anything this code does.
const AT_ONCE = 4;

// What a document has to be under to be saying nothing at all.
//
// **Re-measured on 2026-08-12, against a `--written` sweep.** It was two percent,
// chosen when our side was photographed in a browser and Word's came out of
// poppler, and most of what it allowed for was the second rasteriser rather than
// anything on the page. Through one rasteriser 407 of the 581 clean documents draw
// their worst page **exactly** as Word drew it, cell for cell, and the tail thins
// out immediately after: 14 more under a tenth of a percent, 7 more under a fifth.
// Two percent was letting 72 documents through that are saying something real.
const FLOOR = 0.002;

const share = (count: number, of: number): string =>
  of === 0 ? "n/a" : `${((count / of) * 100).toFixed(1)}%`;

export type LineScore = { readonly placed: number; readonly lines: number };

// The line score of the same documents, where a sweep has left one, so that the
// two readings of a document stand beside each other rather than in two files.
export function linesPlacedIn(text: string): ReadonlyMap<string, LineScore> {
  const placed = new Map<string, LineScore>();

  for (const line of text.split("\n")) {
    if (line.trim() === "") continue;
    const row: unknown = JSON.parse(line);
    if (typeof row !== "object" || row === null) continue;
    const { id, placed: put, lines }: Record<string, unknown> = { ...row };
    if (typeof id !== "string" || typeof put !== "number" || typeof lines !== "number") continue;
    placed.set(id, { placed: put, lines });
  }

  return placed;
}

/**
 * The documents Word drew shrunk, which the sweep beside this one names and this one
 * cannot see at all.
 *
 * **The raster has no way of its own to know.** Word prints a page it cannot fit
 * scaled down and keeps the paper, so both sides come out the same size and every
 * cell inside differs: `d823aa8de433` is our layout under `0.75x + 1`, and it led the
 * ranking of 2026-08-13 at 65.4% while placing 0 of its 88 lines. The paper was
 * checked first and says nothing, page for page.
 *
 * **A stale sweep is still authoritative about this**, which is what makes joining
 * the two files safe here where it would not be for a placement: the scale is read
 * off the size Word set the text at against the size the document asks for, and
 * neither of those is anything this project computes.
 */
export function drawnToAScaleIn(text: string): ReadonlySet<string> {
  const scaled = new Set<string>();

  for (const line of text.split("\n")) {
    if (line.trim() === "") continue;
    const row: unknown = JSON.parse(line);
    if (typeof row !== "object" || row === null) continue;
    const { id, outcome }: Record<string, unknown> = { ...row };
    if (typeof id !== "string" || outcome !== "drawn to a scale") continue;
    scaled.add(id);
  }

  return scaled;
}

const agreementText = (): string =>
  existsSync(resolve(AGREEMENT_PATH)) ? readFileSync(resolve(AGREEMENT_PATH), "utf8") : "";

const linesPlaced = (): ReadonlyMap<string, LineScore> => linesPlacedIn(agreementText());

export function reportOf(
  rows: readonly Looks[],
  lines: ReadonlyMap<string, LineScore>,
  scaled: ReadonlySet<string> = new Set(),
): string {
  const compared = rows.filter((each) => each.outcome === "compared" && !scaled.has(each.id));
  const clean = compared.filter((each) => each.facesStoodIn === 0);

  const totals = (list: readonly Looks[]): string => {
    const interesting = list.reduce((sum, each) => sum + each.interesting, 0);
    const differing = list.reduce((sum, each) => sum + each.differing, 0);
    const right = list.filter((each) => shareOfLooks(each) <= FLOOR).length;
    // **Exactly, and not nearly.** Through one rasteriser a page holding the same
    // things in the same places comes out cell for cell the same, so this counts
    // what is provably right rather than what is close enough to argue about. It
    // was worth nothing while a browser drew our side and every page differed.
    const exact = list.filter((each) => each.differing === 0).length;
    const pages = list.filter((each) => each.pagesOurs !== each.pagesWord).length;
    return (
      `    ${String(list.length)} documents, ${String(interesting)} cells drawn in\n` +
      `    not matching Word's drawing: ${String(differing)} (${share(differing, interesting)})\n` +
      `    documents drawn exactly as Word drew them: ${String(exact)} (${share(exact, list.length)})\n` +
      `    documents inside the floor: ${String(right)} (${share(right, list.length)})\n` +
      `    documents making the wrong number of pages: ${String(pages)}\n`
    );
  };

  const out = [
    `${String(rows.length)} documents`,
    `  compared  ${String(compared.length)}`,
    `  blocked   ${String(rows.filter((each) => each.outcome === "blocked").length)}`,
    `  threw     ${String(rows.filter((each) => each.outcome === "threw").length)}`,
    `  not drawn ${String(rows.filter((each) => each.outcome === "not drawn").length)}`,
    `  drawn to a scale ${String(rows.filter((each) => scaled.has(each.id)).length)}`,
    "",
    "every document compared:",
    totals(compared),
    "documents needing no face stood in, which are the ones worth ranking by:",
    totals(clean),
    // **Ranked by the worst page and not by the document.** A document's own share
    // answers how much of it is wrong, which is not the question a queue asks: one
    // badly wrong page moves a document of twenty-two by about two percent and is
    // the whole of a one-pager, so a ranking read off the total is sorted by how
    // short a document is. `page` says which page to put up beside Word's.
    "the worst of those, by the worst page each of them draws:",
    `  ${"document".padEnd(14)} ${"page".padStart(6)} ${"of it".padStart(6)} ${"whole".padStart(6)} ` +
      `${"cells".padStart(7)} ${"pages".padStart(7)} ${"lines placed".padStart(12)}  asks`,
  ];

  const worst = [...clean].sort((one, other) => worstPageOf(other) - worstPageOf(one)).slice(0, 40);
  for (const each of worst) {
    const line = lines.get(each.id);
    const placed = line === undefined ? "" : `${String(line.placed)}/${String(line.lines)}`;
    const at = each.pages.findIndex((page) => shareOf(page) === worstPageOf(each));
    out.push(
      `  ${each.id.padEnd(14)} ${String(at + 1).padStart(6)} ` +
        `${`${(worstPageOf(each) * 100).toFixed(1)}%`.padStart(6)} ` +
        `${share(each.differing, each.interesting).padStart(6)} ` +
        `${String(each.interesting).padStart(7)} ` +
        `${`${String(each.pagesOurs)}/${String(each.pagesWord)}`.padStart(7)} ` +
        `${placed.padStart(12)}  ${[...new Set(each.asks)].join(" ")}`,
    );
  }

  return out.join("\n");
}

type Wanted = { readonly id: string; readonly path: string };

// **The 966 files are 718 documents.** A tool that walks the directory and
// forgets to dedupe reports one of them seven times over and ranks it that many
// times too high.
function documentsWanted(directory: string): readonly Wanted[] {
  const asked = idsAsked();
  const already = new Set<string>();
  const wanted: Wanted[] = [];
  for (const path of documentsIn(directory)) {
    const id = identityOf(new Uint8Array(readFileSync(path)));
    if (already.has(id)) continue;
    already.add(id);
    if (asked !== null && !asked.has(id)) continue;
    wanted.push({ id, path });
  }
  return wanted;
}

async function sweep(
  wanted: readonly Wanted[],
  workspace: Workspace,
  onProgress: (done: number) => void,
  drawOurs: DrawOurs | null,
): Promise<readonly Looks[]> {
  const rows: Looks[] = [];
  let next = 0;

  // A browser will not share the directory it keeps its profile in, so each of
  // the four drawing at once is given one of its own.
  const worker = async (profile: string): Promise<void> => {
    const mine = { ...workspace, profile };
    for (let at = next++; at < wanted.length; at = next++) {
      const each = wanted[at];
      if (each === undefined) continue;
      const bytes = new Uint8Array(readFileSync(each.path));
      rows.push(await looksOf(bytes, each.id, renderedPath(each.id), mine, undefined, drawOurs));
      onProgress(rows.length);
    }
  };

  await Promise.all(
    Array.from({ length: AT_ONCE }, (_, at) => worker(`${workspace.profile}-${String(at)}`)),
  );
  return rows;
}

async function main(): Promise<void> {
  if (CORPUS_DIRECTORY === null) {
    process.stdout.write(
      "No corpus configured: set DOCX_PAGES_CORPUS to a directory of .docx files.\n",
    );
    return;
  }
  if (!canDraw()) {
    process.stdout.write("No rasteriser: this wants Google Chrome and pdftoppm.\n");
    return;
  }

  // `--written` draws our side by writing a pdf of it and handing that to the very
  // rasteriser Word's goes through, which leaves the floor with nothing in it but
  // the two drawings differing. Measured on 2026-08-11 over the eight documents
  // already known to be right: the floor fell from 0.6% to 0.4%, and eighteen of
  // their twenty pages went to no differing cells at all.
  const drawnByWriter = process.argv.includes("--written");

  const wanted = documentsWanted(CORPUS_DIRECTORY);
  process.stdout.write(`${String(wanted.length)} documents to draw beside Word's drawing\n`);

  const started = Date.now();
  const rows = await sweep(
    wanted,
    workspaceIn(DIRECTORY, false),
    (done) => {
      if (done % 25 !== 0) return;
      const each = (Date.now() - started) / done;
      const left = Math.round((each * (wanted.length - done)) / 1000);
      process.stdout.write(`  ${String(done)}/${String(wanted.length)}, ${String(left)}s left\n`);
    },
    drawnByWriter ? ourWrittenPages : null,
  );

  const ordered = [...rows].sort((one, other) => one.id.localeCompare(other.id));
  mkdirSync(dirname(resolve(REPORT_PATH)), { recursive: true });
  writeFileSync(
    resolve(REPORT_PATH),
    ordered.map((each) => JSON.stringify(each)).join("\n") + "\n",
  );

  process.stdout.write(
    `\n${reportOf(ordered, linesPlaced(), drawnToAScaleIn(agreementText()))}\n\n` +
      `Written to ${REPORT_PATH}\n`,
  );
}

// Compared against this module's own path: a guard naming the built `.js` never
// fires under tsx, which is how these are run, and the run then does nothing at
// all and says nothing about it.
if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}
