import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { referenceCases } from "../testing/cases.js";
import { looksOf, shareOf, shareOfLooks, workspaceIn, worstPageOf, type Looks } from "./compare.js";
import { canDraw } from "./draw.js";
import { ourWrittenPages } from "./written.js";

/**
 * What a ranking read off a raster has to subtract before it means anything.
 *
 * **The name is now wrong, and a reader in a month will take the older meaning.**
 * This began as the price of two rasterisers: our side was photographed in a browser,
 * Word's came out of poppler, and the same glyph in the same place differed pixel for
 * pixel whatever anyone did. `--written` put both sides through the one rasteriser and
 * took that away. What is left is not the cost of drawing twice at all. **It is what
 * this project still gets wrong on the documents it calls right**, and subtracting it
 * as though it were noise forgives a real fault on every document in the corpus.
 *
 * **What it is made of, measured on 2026-08-14**, ours written and both sides through
 * poppler. 361 cells of 52735 over the eight, **every one of them on a first page**:
 *
 * - 175 are a chart nobody has built, on the one document that reports a gap.
 * - 186 are lines landing a whole unit of the device grid from where Word put them,
 *   on three documents. About half the lines of those three land on a unit of their
 *   own, so the arithmetic under them is out by something like a tenth of a point,
 *   which the grid then makes a visible step of.
 *
 * Neither is rasteriser noise, and the reading is given both ways below so that the
 * subtraction can be argued with rather than inherited.
 *
 * **Against the 0.4% CLAUDE.md records from 2026-08-11**: that is this same reading,
 * every cell of all eight, and today's is 0.68%. Eighteen of the twenty pages were
 * exactly equal then and sixteen are now. The older number was never broken down, so
 * what rose between the two is not established here; what is established is what
 * today's is made of, which is the two things above. **The per-page reading is the one
 * to carry forward**, and the one over the documents that report no gap is the one a
 * ranking should subtract.
 */

const DIRECTORY = process.env["DOCX_PAGES_RASTER"] ?? "samples/corpus/raster";

/**
 * The three a report carries about this machine rather than about the document.
 *
 * A face stood in for and a character borrowed from another face are what the layout
 * found when it went looking, not what the file asked for, and **neither stops a page
 * being drawn cell for cell as Word drew it**: on 2026-08-14 one of the eight reported
 * a borrowed character on every page and every page of it was exactly equal. So they
 * are no reason to set a document aside, and `facesStoodIn` is where a reading that
 * cares about them looks.
 */
const ABOUT_THE_MACHINE: ReadonlySet<string> = new Set([
  "substituted-face",
  "character-from-another-face",
  "missing-glyph",
]);

/**
 * Whether a document says of itself that it holds something nobody has built.
 *
 * **Such a document cannot be drawn cell for cell, so its cells are a feature counted
 * as though it were noise.** One of the eight is one: it holds a chart, and 86% of the
 * ink Word put on that page and we did not lies inside the chart's own band, where we
 * draw nothing whatever. Its drawing is a bare `c:chart` with no fallback picture
 * beside it, so nothing in the package can stand in, and **no corpus document holds a
 * chart part at all**: building one would move that page and nothing else anyone has.
 *
 * So it is set aside here rather than built, and the figure is given both ways.
 * `readUnhonoured` is what answers, so this keeps no list of names and says nothing a
 * document does not say about itself.
 */
export const reportsAGap = (looks: Looks): boolean =>
  looks.asks.some((kind) => !ABOUT_THE_MACHINE.has(kind));

/**
 * The floor as a page reads it.
 *
 * **A page is the unit, because a document is not.** A document's own share dilutes
 * a long one: on 2026-08-14 the five differing pages carried every one of the 361
 * cells and the other sixteen pages were exactly equal, which a share over all cells
 * reports as an even 0.7% of everything. `corpus/looks.ts` already ranks by the worst
 * page for the same reason, so the floor it subtracts has to be measured in the same
 * quantity or the two are not comparable.
 */
export type FloorReading = {
  readonly documents: number;
  readonly pages: number;
  // Drawn cell for cell as Word drew it. **Exactly, and not nearly.**
  readonly pagesExactlyEqual: number;
  readonly cells: number;
  readonly differing: number;
  // What a reading over every cell of every document gives, which is the older
  // number and the one to hold beside CLAUDE.md's.
  readonly ofEveryCell: number;
  // The worst single page any of these documents draws, which is what a queue is
  // read off and what a ranking should subtract.
  readonly worstPage: number;
};

export function floorOf(rows: readonly Looks[]): FloorReading {
  const cells = rows.reduce((sum, each) => sum + each.interesting, 0);
  const differing = rows.reduce((sum, each) => sum + each.differing, 0);
  const pages = rows.flatMap((each) => each.pages);

  return {
    documents: rows.length,
    pages: pages.length,
    pagesExactlyEqual: pages.filter((page) => page.differing === 0).length,
    cells,
    differing,
    ofEveryCell: cells === 0 ? 0 : differing / cells,
    worstPage: rows.length === 0 ? 0 : Math.max(...rows.map(worstPageOf)),
  };
}

const asShare = (value: number): string => `${(value * 100).toFixed(2)}%`;

export function readingOf(name: string, reading: FloorReading): string {
  return (
    `  ${name}\n` +
    `    ${String(reading.documents)} documents, ${String(reading.pages)} pages, ` +
    `${String(reading.cells)} cells drawn in\n` +
    `    pages drawn exactly as Word drew them: ${String(reading.pagesExactlyEqual)} of ` +
    `${String(reading.pages)}\n` +
    `    the worst single page: ${asShare(reading.worstPage)}\n` +
    `    over every cell of every document: ${asShare(reading.ofEveryCell)} ` +
    `(${String(reading.differing)} cells)\n`
  );
}

async function main(): Promise<void> {
  if (!canDraw()) {
    process.stdout.write("No rasteriser: this wants Google Chrome and pdftoppm.\n");
    return;
  }

  const cases = referenceCases().filter((each) => each.renderedPath !== null);
  if (cases.length === 0) {
    process.stdout.write("No reference documents: this wants the manifest.\n");
    return;
  }

  const workspace = workspaceIn(DIRECTORY, process.argv.includes("--keep"));

  // `--written` draws our side by writing a pdf of it and handing that to the very
  // rasteriser Word's goes through, which is what leaves this with nothing in it but
  // the two drawings differing. Without it the older meaning of the name still
  // applies and the number is mostly the second rasteriser.
  const written = process.argv.includes("--written");
  const drawOurs = written ? ourWrittenPages : null;

  process.stdout.write(
    `${String(cases.length)} documents whose pages are already right, ` +
      `ours drawn by ${written ? "the pdf writer, through poppler" : "the browser"}\n\n`,
  );

  const rows: Looks[] = [];
  for (const each of cases) {
    const bytes = new Uint8Array(readFileSync(each.documentPath));
    const looks = await looksOf(
      bytes,
      each.id,
      each.renderedPath ?? "",
      workspace,
      undefined,
      drawOurs,
    );
    rows.push(looks);
    const worst = worstPageOf(looks);
    const worstAt = worst === 0 ? -1 : looks.pages.findIndex((page) => shareOf(page) === worst);
    process.stdout.write(
      `  ${each.id.padEnd(6)} ${asShare(shareOfLooks(looks)).padStart(7)} ` +
        `of ${String(looks.interesting).padStart(6)} cells, ` +
        `${String(looks.pages.filter((page) => page.differing === 0).length)}/${String(looks.pages.length)} pages exact, ` +
        `worst page ${asShare(worst)}${worstAt === -1 ? "" : ` (page ${String(worstAt + 1)})`}` +
        (looks.asks.length === 0 ? "" : `, asks for ${looks.asks.join(" and ")}`) +
        (looks.detail === "" ? "" : ` ${looks.detail}`) +
        "\n",
    );
  }

  const answering = rows.filter((each) => !reportsAGap(each));
  const setAside = rows.filter(reportsAGap);

  process.stdout.write(`\n${readingOf("over all of them:", floorOf(rows))}`);
  process.stdout.write(
    `\n${readingOf(
      `over the ${String(answering.length)} that report no gap` +
        `${setAside.length === 0 ? "" : `, ${setAside.map((each) => each.id).join(", ")} set aside`}:`,
      floorOf(answering),
    )}`,
  );
  process.stdout.write(
    "\n  **The floor is what we still get wrong on the documents we call right.**\n" +
      "  It is not what two rasterisers cost: both sides go through the one.\n",
  );
}

if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}
