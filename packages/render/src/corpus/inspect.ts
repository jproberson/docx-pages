import { readFileSync } from "node:fs";

import {
  layOutDocument,
  openDocx,
  substitutingMetrics,
  WORD_FALLBACK_FACES,
  type LaidOutPage,
} from "@docx-pages/core";
import { readTextPlacements, type TextPlacement } from "../pdf/text.js";
import { corpusFaces } from "./faces.js";
import { renderedPath } from "./render.js";
import { documentsIn, identityOf } from "./sweep.js";

// One corpus document beside Word's own drawing of it, line by line.
//
// This is the loop that finds a bug rather than counting one. A score says a
// document disagrees; only the lines say how. Both of the faults found on
// 2026-08-07 were read straight off this: a table style left out of the cascade
// showed up as every row 9pt too tall and the error accumulating down the page,
// and a page break in the wrong place shows up as a whole page off by one constant
// while every horizontal position still agrees.
//
// Read the two columns together. A drift that grows is a height; a constant is a
// break; a difference in the left is an indent or a width, and there has not been
// one of those yet.
//
// **Nothing of the document's text is printed**, only how many characters a line
// holds: the corpus is other people's work and none of it may be quoted or
// committed. The document is named by the twelve characters of its hash that the
// sweep beside it uses.
//
// Run it as:
//   DOCX_PAGES_CORPUS=... tsx --tsconfig packages/render/tsconfig.json \
//     packages/render/src/corpus/inspect.ts <id> [page]
//
// Without a page it reads the first one that disagrees about anything, which is
// almost always the one wanted: a long document agrees about its opening pages and
// the fault is wherever the agreement stops.
// One of our lines beside the item Word drew the same text as, on the same page.
// `off` is how far apart the two are, or null where Word drew that text nowhere on
// the page: the second is a line we broke somewhere Word did not, and reads
// differently from a line that is merely in the wrong place.
type Reading = {
  readonly leftPt: number;
  readonly baselinePt: number;
  readonly characters: number;
  readonly drawn: TextPlacement | null;
  readonly off: number | null;
};

function linesOn(page: LaidOutPage, drawn: readonly TextPlacement[]): readonly Reading[] {
  const readings: Reading[] = [];
  for (const box of page.body) {
    for (const line of box.lines) {
      const text = line.line.segments
        .map((segment) => (segment.kind === "text" ? segment.text : ""))
        .join("")
        .trim();
      if (text === "") continue;

      const near =
        drawn.find(
          (item) => item.pageIndex === page.index && item.text.trim().startsWith(text.slice(0, 12)),
        ) ?? null;
      readings.push({
        leftPt: line.leftPt,
        baselinePt: line.baselinePt,
        characters: text.length,
        drawn: near,
        off:
          near === null
            ? null
            : Math.max(
                Math.abs(line.leftPt - near.leftPt),
                Math.abs(line.baselinePt - near.baselinePt),
              ),
      });
    }
  }
  return readings;
}

async function main(): Promise<void> {
  const wanted = process.argv[2] ?? "";

  for (const path of documentsIn(process.env["DOCX_PAGES_CORPUS"] ?? "")) {
    const bytes = new Uint8Array(readFileSync(path));
    const id = identityOf(bytes);
    if (id !== wanted) continue;

    const measuring = substitutingMetrics(corpusFaces(), WORD_FALLBACK_FACES);
    const laid = layOutDocument(openDocx(bytes), measuring);
    if (laid.kind !== "laid-out") {
      process.stdout.write(`blocked: ${laid.blocker.kind}\n`);
      return;
    }
    const drawn = await readTextPlacements(new Uint8Array(readFileSync(renderedPath(id))));

    process.stdout.write(
      `page geometry: ${JSON.stringify(laid.page)}\n` +
        `body from ${laid.bodyTopPt.toFixed(2)} to ${laid.bodyBottomPt.toFixed(2)}\n` +
        `header lines ${String(laid.header.length)}, footer lines ${String(laid.footer.length)}\n` +
        `our pages ${String(laid.pages.length)}, Word's ${String(Math.max(...drawn.map((d) => d.pageIndex)) + 1)}\n\n`,
    );

    const sizes = new Map<string, number>();
    for (const item of drawn)
      sizes.set(
        `${item.fontName} ${item.fontSizePt.toFixed(1)}`,
        (sizes.get(`${item.fontName} ${item.fontSizePt.toFixed(1)}`) ?? 0) + 1,
      );
    process.stdout.write("Word drew in:\n");
    for (const [what, count] of [...sizes].sort((a, b) => b[1] - a[1]).slice(0, 6)) {
      process.stdout.write(`  ${String(count).padStart(4)}  ${what}\n`);
    }
    const ours = new Map<string, number>();
    for (const page of laid.pages)
      for (const box of page.body)
        for (const line of box.lines)
          for (const segment of line.line.segments)
            if (segment.kind === "text") {
              const key = `${segment.mark.font.kind === "named" ? segment.mark.font.name : "?"} ${segment.mark.fontSizePt.toFixed(1)}`;
              ours.set(key, (ours.get(key) ?? 0) + 1);
            }
    process.stdout.write("we lay out in:\n");
    for (const [what, count] of [...ours].sort((a, b) => b[1] - a[1]).slice(0, 6)) {
      process.stdout.write(`  ${String(count).padStart(4)}  ${what}\n`);
    }
    // Which page to read is the whole of the skill here, and a twelve page document
    // agrees about its first page and disagrees about its eighth. So the tally comes
    // first, page by page, and the lines follow for whichever page is asked for.
    const readings = laid.pages.map((page) => linesOn(page, drawn));

    process.stdout.write("\nlines placed, page by page\n");
    for (const [at, lines] of readings.entries()) {
      const placed = lines.filter((line) => line.off !== null && Math.abs(line.off) <= 1).length;
      const lost = lines.filter((line) => line.off === null).length;
      process.stdout.write(
        `  page ${String(at + 1).padStart(2)}: ${String(placed).padStart(3)} of ${String(lines.length).padStart(3)} placed` +
          (lost === 0 ? "" : `, ${String(lost)} Word drew nowhere on it`) +
          `  cells ${String(laid.pages[at]?.cells.length ?? 0)}\n`,
      );
    }

    // The page asked for, or the first one that disagrees about anything, since
    // that is the one worth reading and finding it by eye is the tedious half.
    const asked = Number(process.argv[3] ?? Number.NaN);
    const firstOff = readings.findIndex((lines) =>
      lines.some((line) => line.off === null || Math.abs(line.off) > 1),
    );
    const at = Number.isFinite(asked) ? asked - 1 : firstOff;
    const page = laid.pages[at];
    if (page === undefined) return;

    process.stdout.write(`\npage ${String(at + 1)}, line by line\n`);
    for (const cell of page.cells) {
      process.stdout.write(
        `  cell at ${cell.leftPt.toFixed(1)},${cell.topPt.toFixed(1)} ${cell.widthPt.toFixed(1)}x${cell.heightPt.toFixed(1)}\n`,
      );
    }
    process.stdout.write("  ours (left, baseline, chars)  |  Word's, by the text\n");
    for (const line of readings[at] ?? []) {
      process.stdout.write(
        `  ${line.leftPt.toFixed(1).padStart(7)} ${line.baselinePt.toFixed(1).padStart(7)} ${String(line.characters).padStart(4)}ch  |  ` +
          (line.drawn === null
            ? "not found"
            : `${line.drawn.leftPt.toFixed(1).padStart(7)} ${line.drawn.baselinePt.toFixed(1).padStart(7)}   dx ${(line.leftPt - line.drawn.leftPt).toFixed(1)}  dy ${(line.baselinePt - line.drawn.baselinePt).toFixed(1)}`) +
          "\n",
      );
    }
    return;
  }
  process.stdout.write("no such document\n");
}

void main();
