import { existsSync } from "node:fs";
import { pathToFileURL } from "node:url";

import { fallbackFacePath } from "../fonts/fallback.js";
import { installedFaceFiles } from "../fonts/installed.js";
import { referenceFonts } from "../testing/cases.js";

// A stylesheet offering the browser the same faces the layout was measured in.
//
// This is not a nicety. A line is measured with the advances of the face it is
// laid out in, and a browser that has not got that face draws the text at those
// advances in whatever it does have, which crushes the spaces between the words
// and changes every glyph on the page. A raster taken through the wrong faces
// says nothing about the layout at all.
//
// A face is offered under the two names `installedFaces` measures it under: the
// family a document names in `w:rFonts`, with the weight and slope beside it, and
// the whole of its own name as a regular.

export type Offered = {
  readonly name: string;
  // What the face calls itself inside its file, which is how one is picked out of a
  // collection. A face is offered under the family a document names as well as
  // under the whole of its own name, so the two are often not the same.
  readonly fullName: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly filePath: string | null;
};

const keyOf = (name: string, bold: boolean, italic: boolean): string =>
  `${name.trim().toLowerCase()}|${bold ? "b" : ""}|${italic ? "i" : ""}`;

// The order is `corpusFaces`': a face somebody measured beats one found by
// looking, and the first to claim a name keeps it.
export function offeredFaces(): readonly Offered[] {
  const found = new Map<string, Offered>();
  const offer = (
    name: string,
    bold: boolean,
    italic: boolean,
    path: string | null,
    fullName = name,
  ): void => {
    if (name.trim() === "") return;
    const key = keyOf(name, bold, italic);
    if (found.has(key)) return;
    found.set(key, {
      name,
      fullName,
      bold,
      italic,
      filePath: path !== null && existsSync(path) ? path : null,
    });
  };

  try {
    for (const font of referenceFonts()) offer(font.name, font.bold, font.italic, font.filePath);
  } catch {
    // The manifest is a private list of where this machine keeps its fonts, and
    // is not always there. Everything else is found by looking.
  }

  for (const face of installedFaceFiles()) {
    // **The internal name only describes the installed file.** Where the pack's
    // file stands in for it the name goes back to the family, which a file holding
    // one face ignores anyway; naming the installed file's face against the pack's
    // file would look the wrong face up in the wrong place.
    const fromPack = fallbackFacePath(face.family);
    offer(
      face.family,
      face.bold,
      face.italic,
      fromPack ?? face.filePath,
      fromPack === null ? face.fullName : face.family,
    );
    if (face.fullName !== face.family) offer(face.fullName, false, false, face.filePath);
  }

  return [...found.values()];
}

const escaped = (name: string): string => name.replace(/["\\]/g, "");

/**
 * A rule per face the browser might be asked for, naming the file it came out of
 * first and the machine's own copy behind that.
 *
 * **A collection is the reason the second source is there.** `Cambria.ttc` holds
 * two faces at once and no browser will unpick one, so the file cannot be handed
 * over and `local()` is the only way to the face. Every face installed on this
 * machine can be reached that way; the ones that cannot are the ones kept in a
 * directory of their own, which is what the file is for.
 */
export function faceStylesheet(): string {
  const rules = offeredFaces().map((face) => {
    const local = `local("${escaped(face.name)}")`;
    const path = face.filePath;
    const source =
      path === null || /\.ttc$/i.test(path)
        ? local
        : `url("${pathToFileURL(path).href}"), ${local}`;
    return (
      `@font-face { font-family: "${escaped(face.name)}";` +
      ` font-weight: ${face.bold ? "bold" : "normal"};` +
      ` font-style: ${face.italic ? "italic" : "normal"};` +
      ` src: ${source}; }`
    );
  });

  return `${rules.join("\n")}\n`;
}
