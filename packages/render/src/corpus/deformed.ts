import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import {
  layOutDocument,
  openDocx,
  substitutingMetrics,
  WORD_FALLBACK_FACES,
} from "@docx-pages/core";

import { agreementWith, type PageAgreement, type PageShape } from "../pdf/agreement.js";
import { readDrawnText } from "../pdf/text.js";
import { corpusFaces } from "./faces.js";
import { renderedPath } from "./render.js";
import { CORPUS_DIRECTORY, documentsIn, identityOf } from "./sweep.js";

// Which documents a preview cannot be shown for, which is a third question and not a
// better answer to either of the other two.
//
// **`agreement.ts` ranks by how many lines missed, and `looks.ts` by how many cells
// differ. Neither can tell a page that is shifted from a page that is deformed**, and
// for a preview those are opposite verdicts. A page broken one paragraph early is the
// page Word drew, moved: it reads well, the reader loses nothing but where the break
// fell, and `looks.ts` scores it near 100% wrong because every line of it sits in the
// next line's cells. A page drawing text over text scores about the same and cannot
// be shown to anybody. The two are one number today, and this is the number that
// tells them apart: `agreementWith` matches each of our lines to the item Word drew
// the same text as and now keeps the displacement, so a page is read by whether one
// offset, one drift, or nothing at all accounts for its lines.
//
// **Keep all three.** The cells see paint, pictures and what covers what, which no
// reading of lines can; the lines say by how much and in which direction, which is
// what names a rule; this says whether the page is worth showing at all. `docs/gaps.md`
// says twice why the first two are both kept, and this adds a third for the same
// reason: the day the raster was built, a page scoring 35 of its 35 lines was wrong
// five ways, and on 2026-08-12 the eight documents whose pages are Word's cell for
// cell were being called `missing` by this, over every bullet in them.
//
// **Nothing this writes names a document.** A row says the first twelve characters of
// the hash of the document's bytes and nothing more, and no number out of it may be
// committed.

const REPORT_PATH = process.env["DOCX_PAGES_CORPUS_DEFORMED"] ?? "samples/corpus/deformed.jsonl";

// A line is agreed when it sits within a point of where Word drew it, as everywhere
// else here. It is also the residual a shift or a drift has to explain a page inside.
const TOLERANCE_PT = 1;

// A drawing of the page at another size is no oracle for this one. The same rule and
// the same reasoning as `agreement.ts`: three of the 718 came back drawn at about
// seven tenths, and every line of them is out in proportion.
const SAME_SCALE = 0.05;

type PageRow = {
  readonly page: number;
  readonly shape: PageShape;
  readonly lines: number;
  readonly matched: number;
  readonly oursAlone: number;
  readonly theirsAlone: number;
  readonly inkOursPt: number;
  readonly inkTheirsPt: number;
  readonly worstPt: number;
  readonly downPt: number | null;
  readonly acrossPt: number | null;
  readonly driftPerPt: number | null;
};

export type Deformed = {
  readonly id: string;
  readonly outcome: "compared" | "blocked" | "threw" | "not drawn" | "drawn to a scale";
  readonly facesStoodIn: number;
  readonly asks: readonly string[];
  readonly pagesOurs: number;
  readonly pagesWord: number;
  // How many of our pages came out each way. What the ranking is read off is the sum
  // of the last two: a page that is deformed or missing content is a page a preview
  // cannot be shown for, and the three before it are pages that can.
  readonly agrees: number;
  readonly shifted: number;
  readonly drifting: number;
  readonly deformed: number;
  readonly missing: number;
  // The line score of the same run, so that the two readings stand in one row rather
  // than in two files. `looks.ts` has to read `agreement.jsonl` for this; here the
  // same comparison answers both.
  readonly lines: number;
  readonly placed: number;
  readonly pages: readonly PageRow[];
  readonly detail: string;
};

const rowOf = (page: PageAgreement): PageRow => ({
  page: page.index + 1,
  shape: page.shape,
  lines: page.lines,
  matched: page.matched,
  oursAlone: page.oursAlone,
  theirsAlone: page.theirsAlone,
  inkOursPt: Math.round(page.inkOursPt),
  inkTheirsPt: Math.round(page.inkTheirsPt),
  worstPt: Number(page.worstPt.toFixed(2)),
  downPt: page.offsetPt === null ? null : Number(page.offsetPt.downPt.toFixed(2)),
  acrossPt: page.offsetPt === null ? null : Number(page.offsetPt.leftPt.toFixed(2)),
  driftPerPt: page.driftPerPt === null ? null : Number(page.driftPerPt.toFixed(4)),
});

const empty = (id: string, outcome: Deformed["outcome"], detail: string): Deformed => ({
  id,
  outcome,
  facesStoodIn: 0,
  asks: [],
  pagesOurs: 0,
  pagesWord: 0,
  agrees: 0,
  shifted: 0,
  drifting: 0,
  deformed: 0,
  missing: 0,
  lines: 0,
  placed: 0,
  pages: [],
  detail,
});

const counting = (pages: readonly PageAgreement[], shape: PageShape): number =>
  pages.filter((page) => page.shape === shape).length;

export async function deformedIn(bytes: Uint8Array, id: string): Promise<Deformed> {
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
      outcome: scaled ? "drawn to a scale" : "compared",
      facesStoodIn: measuring.substitutions().length,
      asks: laid.unhonoured.map((each) => each.kind),
      pagesOurs: laid.pages.length,
      pagesWord: agreed.pagesDrawn,
      agrees: counting(agreed.pages, "agrees"),
      shifted: counting(agreed.pages, "shifted"),
      drifting: counting(agreed.pages, "drifting"),
      deformed: counting(agreed.pages, "deformed"),
      missing: counting(agreed.pages, "missing"),
      lines: agreed.lines,
      placed: agreed.placed,
      pages: agreed.pages.map(rowOf),
      detail: scaled ? `drawn at ${(scale * 100).toFixed(1)}% of the stated size` : "",
    };
  } catch (thrown) {
    return empty(id, "threw", thrown instanceof Error ? thrown.message : String(thrown));
  }
}

// Pages a preview cannot be shown for. A shifted or drifting page is the page Word
// drew, put down wrong; these two are a different drawing.
const unshowable = (row: Deformed): number => row.deformed + row.missing;

const share = (count: number, of: number): string =>
  of === 0 ? "n/a" : `${((count / of) * 100).toFixed(1)}%`;

// The worst page of a document, which is what names the fault: deformed before
// missing before drifting before shifted, as the shapes themselves are ordered, and
// the furthest displaced of those.
const SEVERITY: readonly PageShape[] = ["deformed", "missing", "drifting", "shifted", "agrees"];

function worstPageOf(row: Deformed): PageRow | null {
  const ranked = [...row.pages].sort(
    (one, other) =>
      SEVERITY.indexOf(one.shape) - SEVERITY.indexOf(other.shape) || other.worstPt - one.worstPt,
  );
  return ranked[0] ?? null;
}

export function rankedBy(rows: readonly Deformed[]): readonly Deformed[] {
  return [...rows].sort(
    (one, other) =>
      unshowable(other) - unshowable(one) ||
      unshowable(other) / (other.pagesOurs || 1) - unshowable(one) / (one.pagesOurs || 1) ||
      (worstPageOf(other)?.worstPt ?? 0) - (worstPageOf(one)?.worstPt ?? 0),
  );
}

// How many rows of the ranking are printed. The file holds all of them; this is what
// a reader is asked to look at, and a bucket is worth reading before it is believed.
const PRINTED = 25;

export function reportOf(rows: readonly Deformed[]): string {
  const compared = rows.filter((each) => each.outcome === "compared");
  const clean = compared.filter((each) => each.facesStoodIn === 0);

  const pagesIn = (list: readonly Deformed[], shape: PageShape): number =>
    list.reduce((sum, each) => sum + each[shape], 0);

  const totals = (list: readonly Deformed[]): string => {
    const pages = list.reduce((sum, each) => sum + each.pagesOurs, 0);
    const shapes = SEVERITY.map(
      (shape) =>
        `    ${shape.padEnd(9)} ${String(pagesIn(list, shape)).padStart(6)} pages  ${share(pagesIn(list, shape), pages).padStart(6)}\n`,
    ).join("");
    const showable = list.filter((each) => unshowable(each) === 0).length;
    return (
      `    ${String(list.length)} documents, ${String(pages)} pages\n` +
      shapes +
      `    every page showable in ${String(showable)} documents (${share(showable, list.length)})\n`
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

  // A page count that disagrees is the most visible error a paginated preview can
  // make, and it is not a shape: our page 3 against Word's page 3 says nothing about
  // a page Word made and we did not.
  const miscounted = clean.filter((each) => each.pagesOurs !== each.pagesWord);
  lines.push(
    `  and ${String(miscounted.length)} of the clean make the wrong number of pages\n`,
    "the clean documents a preview cannot be shown for, worst first:",
    `  ${"document".padEnd(14)} ${"pages".padStart(5)} ${"bad".padStart(4)} ${"worst page".padStart(24)} ${"lines".padStart(11)}  asks`,
  );

  const ranked = rankedBy(clean).filter((each) => unshowable(each) > 0);
  for (const row of ranked.slice(0, PRINTED)) {
    const worst = worstPageOf(row);
    const said =
      worst === null
        ? ""
        : `p${String(worst.page)} ${worst.shape} ${worst.worstPt.toFixed(1)}pt` +
          (worst.oursAlone + worst.theirsAlone > 0
            ? ` -${String(worst.oursAlone)}/+${String(worst.theirsAlone)}`
            : "");
    lines.push(
      `  ${row.id.padEnd(14)} ${String(row.pagesOurs).padStart(5)} ${String(unshowable(row)).padStart(4)} ` +
        `${said.padStart(24)} ${`${String(row.placed)}/${String(row.lines)}`.padStart(11)}  ${row.asks.join(" ")}`,
    );
  }
  if (ranked.length > PRINTED) {
    lines.push(`  and ${String(ranked.length - PRINTED)} more, all of them in the file`);
  }

  // Which gaps stand over the documents whose pages cannot be shown. A gap met only
  // in documents whose every page is showable is not what is costing anything, and a
  // gap named by most of the population diagnoses nothing whatever it is beside.
  const byGap = new Map<string, { documents: number; pages: number; bad: number }>();
  for (const each of clean) {
    for (const kind of new Set(each.asks)) {
      const row = byGap.get(kind) ?? { documents: 0, pages: 0, bad: 0 };
      byGap.set(kind, {
        documents: row.documents + 1,
        pages: row.pages + each.pagesOurs,
        bad: row.bad + unshowable(each),
      });
    }
  }

  lines.push("", "what the documents stating each gap cannot show:");
  lines.push(
    `  ${"gap".padEnd(38)} ${"documents".padStart(9)} ${"pages".padStart(6)} ${"cannot show".padStart(12)}`,
  );
  for (const [kind, row] of [...byGap].sort(
    (one, other) => other[1].bad / (other[1].pages || 1) - one[1].bad / (one[1].pages || 1),
  )) {
    lines.push(
      `  ${kind.padEnd(38)} ${String(row.documents).padStart(9)} ${String(row.pages).padStart(6)} ${share(row.bad, row.pages).padStart(12)}`,
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

  const already = new Set<string>();
  const rows: Deformed[] = [];
  const paths = documentsIn(CORPUS_DIRECTORY);

  for (const [at, path] of paths.entries()) {
    const bytes = new Uint8Array(readFileSync(path));
    const id = identityOf(bytes);
    if (already.has(id)) continue;
    already.add(id);

    rows.push(await deformedIn(bytes, id));
    if (rows.length % 25 === 0) {
      process.stdout.write(`  ${String(rows.length)} read, ${String(at + 1)} files opened\n`);
    }
  }

  mkdirSync(dirname(resolve(REPORT_PATH)), { recursive: true });
  writeFileSync(resolve(REPORT_PATH), rows.map((each) => JSON.stringify(each)).join("\n") + "\n");
  process.stdout.write(`\n${reportOf(rows)}\n\nWritten to ${REPORT_PATH}\n`);
}

// Compared against this module's own path: a guard naming the built `.js` never fires
// under tsx, which is how these are run, and the run then does nothing at all and says
// nothing about it.
if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}
