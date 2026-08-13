import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import {
  layOutDocument,
  openDocx,
  substitutingMetrics,
  WORD_FALLBACK_FACES,
} from "@docx-pages/core";

import { agreementWith } from "../pdf/agreement.js";
import { readDrawnText } from "../pdf/text.js";
import { corpusFaces } from "./faces.js";
import { renderedPath } from "./render.js";
import { CORPUS_DIRECTORY, documentsIn, identityOf, idsAsked } from "./sweep.js";

// How far every document in a corpus agrees with Word's own drawing of it.
//
// This is the question the corpus could never ask. Nothing was pinned against it,
// so a sweep could say a document changed, crashed or was refused, and never that
// it was wrong. With Word's pdf beside each document it can: the same comparison
// the eight reference documents are held to, asked of hundreds nobody has measured.
//
// **The eight reference documents are all one page.** So no rule about which page
// something lands on has ever been checked against a real document: not widow
// control, not a row torn by a break, not a paragraph held to the one after it.
// Most of the documents here run to several pages, and they are the first evidence
// of that kind there has been.
//
// **Read a score with the faces in mind.** A document laid out in a face this
// machine has not got has every line in the wrong place before any rule is
// consulted, and its score says something about substitution rather than about
// layout. `facesStoodIn` is carried beside every row so the two can be told apart,
// and the ranking is worth reading off the rows where it is nought.
//
// **Nothing this writes names a document**, as with everything else here: a row
// says the first twelve characters of the hash of the document's bytes and nothing
// more, and no number out of it may be committed.

const REPORT_PATH = process.env["DOCX_PAGES_CORPUS_AGREEMENT"] ?? "samples/corpus/agreement.jsonl";

// A line is agreed when it sits within a point of where Word drew it, which is the
// tolerance the authored documents are held to. Word writes positions on a grid a
// little finer than that.
const TOLERANCE_PT = 1;

// How far the size Word drew a document's text at may stand from the size the
// document asks for before the drawing is taken to be of another page than the one
// laid out here. Three of the 718 came back drawn at about seven tenths, every line
// of them in proportion, and the ranking read that as a layout wrong about
// everything: the worst-placed document that needed no face stood in was one of
// them, and nothing was wrong with it here at all.
//
// Nothing sits between the two answers. A drawing of the same page differs by up to
// 0.6%, which is the tenth of a point Word writes a size to on a 17.5pt run; the
// three shrunk ones are out by 25 and 28.
const SAME_SCALE = 0.05;

export type Agreed = {
  readonly id: string;
  readonly pages: number | null;
  readonly facesStoodIn: number;
  readonly asks: readonly string[];
  // Lines we laid out, lines of those Word drew the same text for, and lines of
  // those we put where Word put them.
  readonly lines: number;
  readonly matched: number;
  readonly placed: number;
  readonly runsMatched: number;
  readonly runsPlaced: number;
  // A drawing of the same page at another size is no oracle for this one: `drawn
  // to a scale` is the row saying so, and its counts are left out of every total.
  readonly outcome: "compared" | "blocked" | "threw" | "not drawn" | "drawn to a scale";
  readonly detail: string;
};

const empty = (id: string, outcome: Agreed["outcome"], detail: string): Agreed => ({
  id,
  pages: null,
  facesStoodIn: 0,
  asks: [],
  lines: 0,
  matched: 0,
  placed: 0,
  runsMatched: 0,
  runsPlaced: 0,
  outcome,
  detail,
});

export async function agreementOf(bytes: Uint8Array, id: string): Promise<Agreed> {
  const drawnPath = renderedPath(id);
  if (!existsSync(drawnPath)) return empty(id, "not drawn", "no pdf of Word's");

  const measuring = substitutingMetrics(corpusFaces(), WORD_FALLBACK_FACES);
  try {
    const pkg = openDocx(bytes);
    const laid = layOutDocument(pkg, measuring);
    if (laid.kind !== "laid-out") return empty(id, "blocked", laid.blocker.kind);

    const drawn = await readDrawnText(new Uint8Array(readFileSync(drawnPath)));
    const agreed = agreementWith(laid, drawn, TOLERANCE_PT);

    const scale = agreed.drawnScale;
    const scaled = scale !== null && Math.abs(scale - 1) > SAME_SCALE;

    return {
      id,
      pages: laid.pages.length,
      facesStoodIn: measuring.substitutions().length,
      asks: laid.unhonoured.map((each) => each.kind),
      lines: agreed.lines,
      matched: agreed.matched,
      placed: agreed.placed,
      runsMatched: agreed.runsMatched,
      runsPlaced: agreed.runsPlaced,
      outcome: scaled ? "drawn to a scale" : "compared",
      detail: scaled ? `drawn at ${(scale * 100).toFixed(1)}% of the stated size` : "",
    };
  } catch (thrown) {
    return empty(id, "threw", thrown instanceof Error ? thrown.message : String(thrown));
  }
}

const share = (count: number, of: number): string =>
  of === 0 ? "n/a" : `${((count / of) * 100).toFixed(1)}%`;

export function reportOf(rows: readonly Agreed[]): string {
  const compared = rows.filter((each) => each.outcome === "compared");
  const clean = compared.filter((each) => each.facesStoodIn === 0);

  const totals = (list: readonly Agreed[]): string => {
    const lines = list.reduce((sum, each) => sum + each.lines, 0);
    const matched = list.reduce((sum, each) => sum + each.matched, 0);
    const placed = list.reduce((sum, each) => sum + each.placed, 0);
    return (
      `    ${String(list.length)} documents, ${String(lines)} lines\n` +
      `    broken as Word broke them: ${String(matched)} (${share(matched, lines)})\n` +
      `    and put where Word put them: ${String(placed)} (${share(placed, matched)} of those)\n`
    );
  };

  const lines = [
    `${String(rows.length)} documents`,
    `  compared  ${String(compared.length)}`,
    `  blocked   ${String(rows.filter((each) => each.outcome === "blocked").length)}`,
    `  threw     ${String(rows.filter((each) => each.outcome === "threw").length)}`,
    `  not drawn ${String(rows.filter((each) => each.outcome === "not drawn").length)}`,
    `  drawn to a scale ${String(rows.filter((each) => each.outcome === "drawn to a scale").length)}`,
    "",
    "every document compared:",
    totals(compared),
    "documents needing no face stood in, which are the ones worth ranking by:",
    totals(clean),
  ];

  // Which gaps stand over the documents that agree with Word least. A gap met only
  // in documents that already land every line is not what is costing anything.
  const byGap = new Map<string, { documents: number; lines: number; placed: number }>();
  for (const each of clean) {
    for (const kind of new Set(each.asks)) {
      const row = byGap.get(kind) ?? { documents: 0, lines: 0, placed: 0 };
      byGap.set(kind, {
        documents: row.documents + 1,
        lines: row.lines + each.matched,
        placed: row.placed + each.placed,
      });
    }
  }

  lines.push("what the documents stating each gap agree with Word about:");
  lines.push(`  ${"gap".padEnd(38)} ${"documents".padStart(9)} ${"lines placed".padStart(13)}`);
  for (const [kind, row] of [...byGap].sort(
    (one, other) => one[1].placed / (one[1].lines || 1) - other[1].placed / (other[1].lines || 1),
  )) {
    lines.push(
      `  ${kind.padEnd(38)} ${String(row.documents).padStart(9)} ${share(row.placed, row.lines).padStart(13)}`,
    );
  }

  return lines.join("\n");
}

async function main(): Promise<void> {
  if (CORPUS_DIRECTORY === null) {
    process.stdout.write(
      "No corpus configured: set DOCX_PAGES_CORPUS to a directory of .docx files.\n",
    );
    return;
  }

  const asked = idsAsked();
  const already = new Set<string>();
  const rows: Agreed[] = [];
  const paths = documentsIn(CORPUS_DIRECTORY);

  for (const [at, path] of paths.entries()) {
    const bytes = new Uint8Array(readFileSync(path));
    const id = identityOf(bytes);
    if (already.has(id)) continue;
    already.add(id);
    if (asked !== null && !asked.has(id)) continue;

    rows.push(await agreementOf(bytes, id));
    if (rows.length % 25 === 0) {
      process.stdout.write(`  ${String(rows.length)} compared, ${String(at + 1)} files read\n`);
    }
  }

  mkdirSync(dirname(resolve(REPORT_PATH)), { recursive: true });
  writeFileSync(resolve(REPORT_PATH), rows.map((each) => JSON.stringify(each)).join("\n") + "\n");
  process.stdout.write(`\n${reportOf(rows)}\n\nWritten to ${REPORT_PATH}\n`);
}

// Compared against this module's own path: a guard naming the built `.js` never
// fires under tsx, which is how these are run, and the run then does nothing at all
// and says nothing about it.
if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}
