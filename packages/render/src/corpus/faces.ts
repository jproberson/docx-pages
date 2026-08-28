import { existsSync } from "node:fs";

import type { SuppliedFace } from "@docx-pages/core";

import { authoredFace } from "../authored/faces.js";
import { fallbackFaces } from "../fonts/fallback.js";
import { installedFaces } from "../fonts/installed.js";
import { suppliedFaces } from "../testing/cases.js";

// The faces a corpus sweep measures with.
//
// A sweep is only worth as much as its fonts. Measured on the eight documents to
// hand: with Cambria alone every one of them is refused as unmeasurable, and with
// the reference manifest's set every one of them lays out. So a sweep gathers the
// widest set the machine can offer and says how many it found, because a run with
// a poor set counts font trouble where there is none.

const MANIFEST_PATH =
  process.env["DOCX_PAGES_REFERENCE_MANIFEST"] ?? "samples/reference-cases.json";

const keyOf = (face: SuppliedFace): string =>
  `${face.name.trim().toLowerCase()}|${face.bold ? "b" : ""}|${face.italic ? "i" : ""}`;

// The manifest is a private list of where this machine keeps its fonts, and is not
// always there. Everything else is found by looking.
function manifestFaces(): readonly SuppliedFace[] {
  if (!existsSync(MANIFEST_PATH)) return [];
  try {
    return suppliedFaces();
  } catch {
    return [];
  }
}

// Gathered once a run. `ourWrittenPages` and `ourPages` both ask for this per document,
// so a sweep of the 718 gathered the same faces 718 times and spent about a quarter of an
// hour doing it.
let gatheredOnce: readonly SuppliedFace[] | null = null;

export function corpusFaces(): readonly SuppliedFace[] {
  gatheredOnce ??= gatherCorpusFaces();
  return gatheredOnce;
}

function gatherCorpusFaces(): readonly SuppliedFace[] {
  const authored = authoredFace();
  // **A measured face beats a found one, and the order here is the whole of how
  // that is said.** The manifest is a person naming the file behind a name; the
  // fallback pack is a handful of faces whose answers were measured against Word
  // one at a time; scanning the disk is neither, and it goes last.
  //
  // Measured on 2026-08-10, by putting the scan first and watching 89 documents
  // that already had every face they needed fall from 93.4% placed to 87.1%. Two
  // files on this machine answer to Times New Roman: the system's states a line gap
  // of 87 units and Word's own states none, and `fallback.ts` merges them because
  // Word makes a Times New Roman line 13.80pt tall at 12pt, which is the 87. A scan
  // that reaches Word's file first hands back the same face with no gap at all and
  // every Times New Roman line comes out half a point short.
  const gathered = [
    ...manifestFaces(),
    ...fallbackFaces(),
    ...(authored === null ? [] : [authored]),
    ...installedFaces(),
  ];

  const found = new Map<string, SuppliedFace>();
  for (const face of gathered) if (!found.has(keyOf(face))) found.set(keyOf(face), face);
  return [...found.values()];
}
