import { basename } from "node:path";

import { authoredCases } from "../authored/cases.js";
import { referenceCases, type ReferenceCase } from "../testing/cases.js";
import {
  authoredFaceSet,
  FRAME_STYLES,
  wordFaceSet,
  writeBrowser,
  writeFonts,
  writePreview,
} from "./pages.js";

const OUTPUT_DIRECTORY = "samples/preview";

const titleOf = (each: ReferenceCase): string =>
  `${each.id}. ${basename(each.documentPath).replace(/\.docx$/i, "")}`;

function main(): void {
  // The documents written for the suite are previewed beside the real ones, so a
  // gap the numbers name can be looked at as well as counted.
  const cases = [...referenceCases(), ...authoredCases()];

  if (cases.length === 0) {
    process.stdout.write("no reference cases; point DOCX_PAGES_REFERENCE_MANIFEST at a manifest\n");
    return;
  }

  const authored = authoredFaceSet();
  const sets = authored === null ? [wordFaceSet()] : [wordFaceSet(), authored];

  for (const set of sets) {
    process.stdout.write(`${writeFonts(OUTPUT_DIRECTORY, set)}\n`);
    for (const each of cases) {
      for (const frames of FRAME_STYLES)
        process.stdout.write(`${writePreview(OUTPUT_DIRECTORY, each, set, frames)}\n`);
    }
  }
  process.stdout.write(`${writeBrowser(OUTPUT_DIRECTORY, cases, sets, titleOf)}\n`);
}

main();
