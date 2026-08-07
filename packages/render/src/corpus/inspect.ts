import { readFileSync } from "node:fs";

import {
  layOutDocument,
  openDocx,
  substitutingMetrics,
  WORD_FALLBACK_FACES,
} from "@docx-pages/core";
import { readTextPlacements } from "../pdf/text.js";
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
//     packages/render/src/corpus/inspect.ts <id>
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
    process.stdout.write("\nlines per page: ours vs Word's\n");
    for (const page of laid.pages) {
      let mine = 0;
      for (const box of page.body)
        for (const line of box.lines)
          if (line.line.segments.some((s) => s.kind === "text" && s.text.trim() !== "")) mine += 1;
      const theirs = drawn.filter((d) => d.pageIndex === page.index && d.text.trim() !== "").length;
      process.stdout.write(
        `  page ${String(page.index + 1)}: ours ${String(mine)}, Word's ${String(theirs)}\n`,
      );
    }
    for (const page of laid.pages.slice(0, 1)) {
      process.stdout.write(`\ncells on page 1: ${String(page.cells.length)}\n`);
      for (const cell of page.cells.slice(0, 6)) {
        process.stdout.write(
          `  cell at ${cell.leftPt.toFixed(1)},${cell.topPt.toFixed(1)} ${cell.widthPt.toFixed(1)}x${cell.heightPt.toFixed(1)}\n`,
        );
      }
    }
    process.stdout.write(
      "\nours (page, left, baseline, chars)  |  Word's nearest by text length\n",
    );
    for (const page of laid.pages.slice(0, 2)) {
      for (const box of page.body.slice(0, 12)) {
        for (const line of box.lines) {
          const text = line.line.segments
            .map((s) => (s.kind === "text" ? s.text : ""))
            .join("")
            .trim();
          if (text === "") continue;
          const theirs = drawn.filter(
            (d) => d.pageIndex === page.index && d.text.trim().startsWith(text.slice(0, 12)),
          );
          const near = theirs[0];
          process.stdout.write(
            `  p${String(page.index + 1)} ${line.leftPt.toFixed(1).padStart(6)} ${line.baselinePt.toFixed(1).padStart(7)} ${String(text.length).padStart(4)}ch  |  ` +
              (near === undefined
                ? "not found"
                : `${near.leftPt.toFixed(1).padStart(6)} ${near.baselinePt.toFixed(1).padStart(7)}   dx ${(line.leftPt - near.leftPt).toFixed(1)}  dy ${(line.baselinePt - near.baselinePt).toFixed(1)}`) +
              "\n",
          );
        }
      }
    }
    return;
  }
  process.stdout.write("no such document\n");
}

void main();
