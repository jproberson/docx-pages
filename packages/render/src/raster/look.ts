import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { renderedPath } from "../corpus/render.js";
import { CORPUS_DIRECTORY, documentsIn, identityOf } from "../corpus/sweep.js";
import { ourPages, wordPages, workspaceIn } from "./compare.js";
import { differenceBetween, gridOf, shareOf } from "./difference.js";
import { canDraw } from "./draw.js";
import { overlayOf } from "./overlay.js";
import { writePng } from "./png.js";

// One corpus document's pages laid over Word's, page by page, to be looked at.
//
// This is to the raster sweep what `inspect.ts` is to the line comparison: the
// ranking says which document, and this says what about it. **Neither replaces
// the other.** The raster says a page is wrong and where on it; the lines say by
// how much and in which direction, which is what names the rule.
//
// Nothing here prints a word out of a document. What is written is a picture of
// where the ink is, under the first twelve characters of the hash of the bytes.

const DIRECTORY = process.env["DOCX_PAGES_RASTER"] ?? "samples/corpus/raster";

async function main(): Promise<void> {
  const wanted = new Set(process.argv.slice(2));
  if (wanted.size === 0 || CORPUS_DIRECTORY === null || !canDraw()) {
    process.stdout.write("usage: look.ts <id>..., with DOCX_PAGES_CORPUS set\n");
    return;
  }

  const workspace = workspaceIn(DIRECTORY, false);

  for (const path of documentsIn(CORPUS_DIRECTORY)) {
    if (wanted.size === 0) break;
    const bytes = new Uint8Array(readFileSync(path));
    const id = identityOf(bytes);
    if (!wanted.has(id)) continue;
    wanted.delete(id);

    const ours = (await ourPages(bytes, id, workspace)).pages;
    const theirs = await wordPages(renderedPath(id), id, workspace);

    for (let at = 0; at < Math.max(ours.length, theirs.length); at += 1) {
      const mine = ours[at] ?? null;
      const yours = theirs[at] ?? null;
      const difference = differenceBetween(
        mine === null ? null : gridOf(mine),
        yours === null ? null : gridOf(yours),
      );
      const written = resolve(DIRECTORY, `${id}.page-${String(at + 1)}.png`);
      writeFileSync(written, writePng(overlayOf(mine, yours)));
      process.stdout.write(
        `  page ${String(at + 1).padStart(3)}  ` +
          `${(shareOf(difference) * 100).toFixed(1).padStart(6)}% of ` +
          `${String(difference.interesting).padStart(5)} cells  ${written}\n`,
      );
    }
  }
}

if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}
