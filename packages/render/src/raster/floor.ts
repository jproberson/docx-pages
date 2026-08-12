import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { referenceCases } from "../testing/cases.js";
import { looksOf, shareOfLooks, workspaceIn } from "./compare.js";
import { canDraw } from "./draw.js";
import { ourWrittenPages } from "./written.js";

// What two rasterisers cost when nothing is wrong.
//
// **A pixel count never reaches nought**: Chrome and Word hint and antialias
// differently, so the same glyph drawn in the same place at the same size differs
// pixel for pixel. Any ranking read off a raster is therefore read off a floor,
// and a floor nobody has measured is a floor nobody can subtract.
//
// So this asks the documents whose pages are already known to be right, line by
// line, against Word's own pdf. Whatever they score is what the two rasterisers
// cost, and a corpus document scoring the same is saying nothing at all.

const DIRECTORY = process.env["DOCX_PAGES_RASTER"] ?? "samples/corpus/raster";

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
  // rasteriser Word's goes through, which is what leaves the floor with nothing in
  // it but the two drawings differing.
  const written = process.argv.includes("--written");
  const drawOurs = written ? ourWrittenPages : null;

  process.stdout.write(
    `${String(cases.length)} documents whose pages are already right, ` +
      `ours drawn by ${written ? "the pdf writer, through poppler" : "the browser"}\n\n`,
  );

  let interesting = 0;
  let differing = 0;

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
    interesting += looks.interesting;
    differing += looks.differing;
    process.stdout.write(
      `  ${each.id.padEnd(24)} ${(shareOfLooks(looks) * 100).toFixed(1).padStart(6)}% ` +
        `of ${String(looks.interesting).padStart(6)} cells, ` +
        `${String(looks.pagesOurs)}/${String(looks.pagesWord)} pages ${looks.detail}\n`,
    );
  }

  process.stdout.write(
    `\nthe floor: ${((differing / (interesting || 1)) * 100).toFixed(1)}% of ${String(interesting)} cells\n`,
  );
}

if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}
