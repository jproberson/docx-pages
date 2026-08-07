import { execFileSync } from "node:child_process";
import { resolve } from "node:path";

import { corpusFaces } from "../corpus/faces.js";
import { fallbackFacePath } from "../fonts/fallback.js";
import { referenceFonts } from "../testing/cases.js";
import { caseOf, gatherDocuments } from "./documents.js";
import {
  FRAME_STYLES,
  writeBrowser,
  writeFonts,
  writePreview,
  type FaceFile,
  type FaceSet,
} from "./pages.js";

// Any document at all, beside Word's own pdf of it.
//
// The suites answer for the documents they already know. This answers for one
// handed over on the command line, which is the shape every question about a
// corpus document takes: draw it, ask Word to draw it, and put the two side by
// side. Nothing here is a test and nothing is recorded; it is for looking.

const OUTPUT_DIRECTORY = "samples/check";
const RENDER_SCRIPT = resolve("packages/render/src/authored/render.applescript");

// One script for all of them, since Word leaves a document open unless it is asked
// to close and the script closes behind itself. Word being absent or busy is not
// fatal: the rendering is still worth looking at, only without its answer beside
// it.
function askWordToRender(paths: readonly string[]): void {
  try {
    execFileSync("osascript", [RENDER_SCRIPT, ...paths], {
      encoding: "utf8",
      maxBuffer: 64 * 1024 * 1024,
    });
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    process.stdout.write(`Word did not render: ${detail.split("\n")[0] ?? ""}\n`);
    process.stdout.write("  a -1728 or a timeout means a document is still open; quit Word\n");
  }
}

// The widest set of faces the machine can offer, which is what a sweep measures
// with: a document drawn in a poor set shows font trouble where there is none. The
// files behind them are what the browser draws with, and a collection holds
// several faces in one file, which no browser will unpick.
function machineFaceSet(): FaceSet {
  const files = new Map<string, FaceFile>();
  const keyOf = (name: string, bold: boolean, italic: boolean): string =>
    `${name.toLowerCase()}|${bold ? "b" : ""}|${italic ? "i" : ""}`;

  for (const font of referenceFonts()) {
    if (font.filePath === null) continue;
    files.set(keyOf(font.name, font.bold, font.italic), { ...font, filePath: font.filePath });
  }

  const faces = corpusFaces();
  for (const face of faces) {
    const key = keyOf(face.name, face.bold, face.italic);
    if (files.has(key)) continue;
    const path = fallbackFacePath(face.name);
    if (path === null || /\.ttc$/i.test(path)) continue;
    files.set(key, { name: face.name, bold: face.bold, italic: face.italic, filePath: path });
  }

  return { id: "word", label: "As this machine has it", faces, files: [...files.values()] };
}

function main(): void {
  const paths = process.argv.slice(2).filter((each) => !each.startsWith("--"));
  const serving = !process.argv.includes("--no-serve");

  if (paths.length === 0) {
    process.stdout.write("usage: pnpm check <document.docx> [more.docx...] [--no-serve]\n");
    process.exitCode = 1;
    return;
  }

  const documents = gatherDocuments(OUTPUT_DIRECTORY, paths);
  askWordToRender(documents.map((each) => each.path));

  const set = machineFaceSet();
  process.stdout.write(`${writeFonts(OUTPUT_DIRECTORY, set)}\n`);

  const cases = documents.flatMap(({ id, path }) => {
    const each = caseOf(id, path);
    if (each.renderedPath === null) process.stdout.write(`  ${id}: no pdf from Word\n`);
    try {
      for (const frames of FRAME_STYLES) writePreview(OUTPUT_DIRECTORY, each, set, frames);
      return [each];
    } catch (error) {
      // A document this project refuses has no page to put beside Word's, which is
      // an answer in itself rather than a failure of the run.
      process.stdout.write(`  ${id}: ${error instanceof Error ? error.message : String(error)}\n`);
      return [];
    }
  });

  if (cases.length === 0) {
    process.stdout.write("nothing to show\n");
    process.exitCode = 1;
    return;
  }

  writeBrowser(OUTPUT_DIRECTORY, cases, [set], (each) => each.id);

  if (!serving) {
    process.stdout.write(`${resolve(OUTPUT_DIRECTORY, "index.html")}\n`);
    return;
  }

  process.env["DOCX_PAGES_PREVIEW_START"] = "/check/index.html";
  void import("./serve.js");
}

main();
