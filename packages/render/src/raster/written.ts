import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import {
  layOutDocument,
  openDocx,
  substitutingMetrics,
  writePdf,
  WORD_FALLBACK_FACES,
  type PdfFont,
  type SubstitutingMetrics,
} from "@docx-pages/core";
import { facesUsed } from "@docx-pages/core/internal";

import { corpusFaces } from "../corpus/faces.js";
import { drawPdf } from "./draw.js";
import { offeredFaces } from "./faces.js";
import type { OurPages, Workspace } from "./compare.js";
import type { RasterImage } from "./png.js";

// Our page drawn the way Word's is: written out as a pdf and handed to the same
// rasteriser.
//
// **This is what takes the floor out of a raster comparison.** Our side used to be
// drawn by a browser because a browser was the only thing that drew what the viewer
// describes; Word's has always come out of poppler. Two rasterisers hint and
// antialias differently, so the same glyph in the same place at the same size
// differs pixel for pixel, and `floor.ts` exists to price what that costs before
// any ranking is read. Both sides through one rasteriser is a floor made of nothing
// but the two drawings actually differing.
//
// It is the same layout either way. `writePdf` walks `drawablesOf`, which is what
// the viewer walks, so this is not a second reading of the page: it is the one
// reading, put through a different pen.

const keyOf = (name: string, bold: boolean, italic: boolean): string =>
  `${name.trim().toLowerCase()}|${bold ? "b" : ""}|${italic ? "i" : ""}`;

// Read once for a whole sweep. This machine offers about twelve hundred faces and
// a document draws in a handful, so the files are opened on demand and kept.
const opened = new Map<string, Uint8Array | null>();

const bytesOfFile = (filePath: string): Uint8Array | null => {
  const held = opened.get(filePath);
  if (held !== undefined) return held;
  let bytes: Uint8Array | null;
  try {
    bytes = new Uint8Array(readFileSync(filePath));
  } catch {
    bytes = null;
  }
  opened.set(filePath, bytes);
  return bytes;
};

/**
 * The bytes of every face this page might be drawn in, under the names the layout
 * measured them as.
 *
 * **A face stood in for is carried under the name the document asked for**, which
 * is the same rule the viewer's export follows and for the same reason: that is the
 * name the layout measured it as, and a pdf carrying it under the stand-in's own
 * name is refused for a face it is in fact holding.
 *
 * Generous on purpose. A face named here that the page never draws in costs the
 * reading of one file; a face missing refuses the whole document.
 */
function facesToEmbed(
  bytes: Uint8Array,
  measuring: SubstitutingMetrics,
): { readonly fonts: readonly PdfFont[]; readonly missing: readonly string[] } {
  const files = new Map<string, string>();
  for (const face of offeredFaces()) {
    if (face.filePath === null) continue;
    files.set(keyOf(face.name, face.bold, face.italic), face.filePath);
  }

  const fonts: PdfFont[] = [];
  const missing: string[] = [];
  const taken = new Set<string>();

  const carry = (name: string, bold: boolean, italic: boolean, from: string): void => {
    const key = keyOf(name, bold, italic);
    if (name.trim() === "" || taken.has(key)) return;
    const filePath = files.get(from);
    if (filePath === undefined) return;
    const read = bytesOfFile(filePath);
    if (read === null) return;
    taken.add(key);
    fonts.push({ name, bold, italic, bytes: read });
  };

  // What the document names, under its own name.
  for (const face of facesUsed(openDocx(bytes))) {
    if (face.name === null) continue;
    const key = keyOf(face.name, face.bold, face.italic);
    if (files.has(key)) carry(face.name, face.bold, face.italic, key);
    else missing.push(face.name);
  }

  // What stood in for a name nothing answered to, under the asked-for name.
  for (const each of measuring.substitutions()) {
    carry(
      each.requested.name,
      each.requested.bold,
      each.requested.italic,
      keyOf(each.used.name, each.used.bold, each.used.italic),
    );
  }

  // The faces a character is borrowed from, which a run never names and the page
  // still draws out of.
  for (const each of measuring.fallbackCharacters()) {
    carry(
      each.used.name,
      each.used.bold,
      each.used.italic,
      keyOf(each.used.name, each.used.bold, each.used.italic),
    );
  }
  for (const name of WORD_FALLBACK_FACES) {
    for (const bold of [false, true]) {
      for (const italic of [false, true]) {
        carry(name, bold, italic, keyOf(name, bold, italic));
      }
    }
  }

  return { fonts, missing };
}

/**
 * Every page of one document, drawn out of a pdf we wrote ourselves.
 *
 * Answers the same shape `ourPages` does, so a comparison can take either and the
 * only difference between them is which pen drew our side.
 */
export async function ourWrittenPages(
  bytes: Uint8Array,
  id: string,
  workspace: Workspace,
): Promise<OurPages> {
  const measuring = substitutingMetrics(corpusFaces(), WORD_FALLBACK_FACES);
  const pkg = openDocx(bytes);
  const laid = layOutDocument(pkg, measuring);
  if (laid.kind !== "laid-out") throw new Error(`blocked: ${laid.blocker.kind}`);

  const { fonts, missing } = facesToEmbed(bytes, measuring);
  if (missing.length > 0 && fonts.length === 0) {
    throw new Error(`blocked: no face on this machine for ${missing.join(", ")}`);
  }

  const pdfPath = resolve(workspace.directory, `${id}.ours.pdf`);
  const written = writePdf(laid, {
    fonts,
    imageBytes: (part) => pkg.parts.get(part),
    metricsFor: measuring.metricsFor,
  });

  const { writeFileSync, rmSync } = await import("node:fs");
  writeFileSync(pdfPath, written);

  let pages: readonly RasterImage[];
  try {
    pages = await drawPdf(pdfPath, workspace.directory, `${id}.ours`, !workspace.keep);
  } finally {
    if (!workspace.keep) rmSync(pdfPath, { force: true });
  }

  return {
    pages,
    facesStoodIn: measuring.substitutions().length,
    asks: laid.unhonoured.map((each) => each.kind),
  };
}
