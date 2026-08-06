import { existsSync } from "node:fs";

import type { SuppliedFace } from "@docx-pages/core";

import { authoredFace } from "../authored/faces.js";
import { fallbackFaces } from "../fonts/fallback.js";
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

export function corpusFaces(): readonly SuppliedFace[] {
  const authored = authoredFace();
  const gathered = [
    ...manifestFaces(),
    ...fallbackFaces(),
    ...(authored === null ? [] : [authored]),
  ];

  const found = new Map<string, SuppliedFace>();
  for (const face of gathered) if (!found.has(keyOf(face))) found.set(keyOf(face), face);
  return [...found.values()];
}
