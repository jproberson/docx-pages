import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import {
  layOutDocument,
  openDocx,
  isDocxPagesError,
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

// The one face a refusal names, or null where the refusal was about something else.
const shortOfFace = (
  thrown: unknown,
): { readonly name: string; readonly bold: boolean; readonly italic: boolean } | null => {
  if (!isDocxPagesError(thrown) || thrown.code !== "font-not-supplied") return null;
  const { fontName, bold, italic } = thrown.context;
  if (typeof fontName !== "string" || fontName.trim() === "") return null;
  return { name: fontName, bold: bold === true, italic: italic === true };
};

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
): {
  readonly fonts: readonly PdfFont[];
  readonly missing: readonly string[];
  // Adds one face the writer says it needs, answering whether anything was found.
  // The writer names what it is short of, which beats guessing where a run got a
  // name the document never states: a theme resolves one, and a style cascade
  // another.
  readonly alsoCarry: (name: string, bold: boolean, italic: boolean) => boolean;
} {
  const offered = new Map<string, { readonly filePath: string; readonly fullName: string }>();
  for (const face of offeredFaces()) {
    if (face.filePath === null) continue;
    offered.set(keyOf(face.name, face.bold, face.italic), {
      filePath: face.filePath,
      fullName: face.fullName,
    });
  }

  const fonts: PdfFont[] = [];
  const missing: string[] = [];
  const taken = new Set<string>();

  // **`name` is what the face answers to and `faceName` is what it is called inside
  // its file.** They come apart at every stand-in: the file behind one is a
  // collection holding its own faces and none by the name the document asked for.
  const carry = (name: string, bold: boolean, italic: boolean, from: string): void => {
    const key = keyOf(name, bold, italic);
    if (name.trim() === "" || taken.has(key)) return;
    const face = offered.get(from);
    if (face === undefined) return;
    const read = bytesOfFile(face.filePath);
    if (read === null) return;
    taken.add(key);
    fonts.push({ name, bold, italic, bytes: read, faceName: face.fullName });
  };

  // **The stand-ins are claimed first**, because a name is claimed once and the
  // layout has already settled which face answers for it. A name the machine also
  // offers under something of its own would otherwise win here and be drawn in a
  // face the layout never measured.
  for (const each of measuring.substitutions()) {
    carry(
      each.requested.name,
      each.requested.bold,
      each.requested.italic,
      keyOf(each.used.name, each.used.bold, each.used.italic),
    );
  }

  // What the document names, under its own name.
  for (const face of facesUsed(openDocx(bytes))) {
    if (face.name === null) continue;
    const key = keyOf(face.name, face.bold, face.italic);
    if (offered.has(key)) carry(face.name, face.bold, face.italic, key);
    else if (!taken.has(keyOf(face.name, face.bold, face.italic))) missing.push(face.name);
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

  const alsoCarry = (name: string, bold: boolean, italic: boolean): boolean => {
    const before = fonts.length;
    for (const [wantBold, wantItalic] of [
      [bold, italic],
      [bold, false],
      [false, italic],
      [false, false],
    ] as const) {
      const from = keyOf(name, wantBold, wantItalic);
      if (!offered.has(from)) continue;
      carry(name, bold, italic, from);
      break;
    }
    return fonts.length > before;
  };

  return { fonts, missing, alsoCarry };
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

  const { fonts, missing, alsoCarry } = facesToEmbed(bytes, measuring);
  if (missing.length > 0 && fonts.length === 0) {
    throw new Error(`blocked: no face on this machine for ${missing.join(", ")}`);
  }

  // **The writer is asked what it is short of rather than guessed at.** A run can
  // reach a face name through a theme or a style cascade that nothing here walks,
  // so gathering the faces up front cannot be complete. `font-not-supplied` names
  // exactly one face; carry it and write again. It ends either way: each turn adds
  // a face or gives up.
  const pdfPath = resolve(workspace.directory, `${id}.ours.pdf`);
  let written: Uint8Array | null = null;
  for (let tries = 0; written === null && tries < 8; tries += 1) {
    try {
      written = writePdf(laid, {
        fonts,
        imageBytes: (part) => pkg.parts.get(part),
        metricsFor: measuring.metricsFor,
      });
    } catch (thrown) {
      const short = shortOfFace(thrown);
      if (short === null || !alsoCarry(short.name, short.bold, short.italic)) throw thrown;
    }
  }
  if (written === null) throw new Error("blocked: the writer kept asking for faces");

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
