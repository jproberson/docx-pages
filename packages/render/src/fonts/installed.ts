import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { extname, join } from "node:path";

import { readFontFaces, readFontFile, type SuppliedFace } from "@docx-pages/core";

/**
 * Every face this machine can offer a document, found by opening the fonts rather
 * than by naming them.
 *
 * The sweep used to gather its faces from three places that name what they hold:
 * the reference manifest, the fallback pack and the one face the authored suite is
 * written in. That came to forty, while Word alone ships 280 files a few
 * directories away, so 484 of the 715 corpus documents were laid out in Cambria
 * standing in for a Verdana that was on the disk the whole time. **A face named in
 * a list is a face somebody remembered**, and nobody was going to remember 280.
 *
 * A file cannot be matched to a document by its own name: `seguisb.ttf` is Segoe
 * UI Semibold, `calibril.ttf` is Calibri Light, and `Cambria.ttc` is two faces at
 * once. So each is opened and asked what it calls itself.
 */
const DIRECTORIES: readonly string[] = [
  "/Applications/Microsoft Word.app/Contents/Resources/DFonts",
  "/Library/Fonts/Microsoft",
  "/Library/Fonts",
  "/System/Library/Fonts/Supplemental",
];

// Directories of this machine's own, which is where a face that ships with neither
// Word nor the system goes. It is an environment variable rather than a path here
// for the same reason the reference manifest is one: what a machine has and where
// it keeps it is nobody else's business, and some of it may not be redistributable.
const STATED = process.env["DOCX_PAGES_FONT_DIRECTORIES"];

const READABLE = new Set([".ttf", ".ttc", ".otf", ".woff"]);

export function fontDirectories(): readonly string[] {
  const stated = STATED === undefined ? [] : STATED.split(":").filter((each) => each !== "");
  return [...stated, ...DIRECTORIES].filter((each) => existsSync(each));
}

function filesIn(directory: string): readonly string[] {
  let entries: readonly string[];
  try {
    entries = readdirSync(directory);
  } catch {
    return [];
  }
  return entries
    .filter((each) => READABLE.has(extname(each).toLowerCase()))
    .map((each) => join(directory, each))
    .filter((each) => {
      try {
        return statSync(each).isFile();
      } catch {
        return false;
      }
    });
}

const keyOf = (name: string, bold: boolean, italic: boolean): string =>
  `${name.trim().toLowerCase()}|${bold ? "b" : ""}|${italic ? "i" : ""}`;

/**
 * The faces found in the directories this machine keeps its fonts in.
 *
 * A face is offered under two names: the family a document names in `w:rFonts`,
 * with the weight and slope beside it, and the whole of its own name where that
 * differs. Both are how a real document asks. `Calibri Light` is the family
 * `Calibri` to some of its own name records and a family of its own to a document,
 * and a document that says `w:rFonts w:ascii="Calibri Light"` with no `w:b` beside
 * it wants the file, not Calibri emboldened.
 *
 * The first face to claim a name keeps it, so a directory earlier in the list wins
 * and a font's own regular cut is never shadowed by a bold one from elsewhere.
 */
export type InstalledFaceFile = {
  readonly family: string;
  readonly fullName: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly filePath: string;
};

/**
 * The same faces named beside the file each came out of, which is what a browser
 * asked to draw a page needs and what measuring it never did. Only the name
 * tables are read, so this costs a fraction of what `installedFaces` costs.
 *
 * A collection holds several faces in one file and no browser will unpick one, so
 * a face out of a `.ttc` is named here and has to be reached some other way.
 */
export function installedFaceFiles(
  directories: readonly string[] = fontDirectories(),
): readonly InstalledFaceFile[] {
  const found: InstalledFaceFile[] = [];

  for (const directory of directories) {
    for (const filePath of filesIn(directory)) {
      try {
        for (const face of readFontFaces(new Uint8Array(readFileSync(filePath)))) {
          found.push({ ...face, filePath });
        }
      } catch {
        // A file this reader cannot take is a face this machine cannot offer.
      }
    }
  }

  return found;
}

export function installedFaces(
  directories: readonly string[] = fontDirectories(),
): readonly SuppliedFace[] {
  const found = new Map<string, SuppliedFace>();

  const offer = (name: string, bold: boolean, italic: boolean, build: () => SuppliedFace): void => {
    if (name === "") return;
    const key = keyOf(name, bold, italic);
    if (found.has(key)) return;
    try {
      found.set(key, { ...build(), name, bold, italic });
    } catch {
      // A file this reader cannot take is a face this machine cannot offer, which
      // is what it was before it was looked for. A directory of fonts is not a
      // thing to refuse a whole sweep over.
    }
  };

  for (const directory of directories) {
    for (const path of filesIn(directory)) {
      let bytes: Uint8Array;
      let faces: ReturnType<typeof readFontFaces>;
      try {
        bytes = new Uint8Array(readFileSync(path));
        faces = readFontFaces(bytes);
      } catch {
        continue;
      }

      for (const face of faces) {
        // A collection is read by the name of the face wanted out of it, so the
        // whole name is what asks; a file of one face ignores it either way.
        const build = (): SuppliedFace => {
          const read = readFontFile(bytes, face.fullName === "" ? undefined : face.fullName);
          return {
            name: face.family,
            bold: face.bold,
            italic: face.italic,
            metrics: read.metrics,
            advances: read.advances,
            // The pairs the file states, so a run that asks to kern kerns in the face
            // Word drew it in rather than in a face measured without them.
            kerning: read.kerning,
            // What each glyph draws and what the face says about setting
            // mathematics, without which a face carrying a MATH table cannot set an
            // equation at all.
            ink: read.ink,
            math: read.math,
            sansSerif: read.sansSerif,
          };
        };

        offer(face.family, face.bold, face.italic, build);
        if (face.fullName !== face.family) offer(face.fullName, false, false, build);
      }
    }
  }

  return [...found.values()];
}
