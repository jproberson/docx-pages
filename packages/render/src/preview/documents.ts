import { copyFileSync, existsSync, mkdirSync } from "node:fs";
import { basename, resolve } from "node:path";

import type { ReferenceCase } from "../testing/cases.js";

// The documents `check.tsx` was handed on the command line, dressed as reference
// cases so that everything built for the suite's own draws them too.

export type GatheredDocument = {
  readonly id: string;
  readonly path: string;
};

// Word will not export into a directory it has not been granted, and it has been
// granted this project's `samples`. A document is copied in rather than exported
// where it lies, which also leaves a corpus untouched by having been looked at.
export function gatherDocuments(
  directory: string,
  paths: readonly string[],
): readonly GatheredDocument[] {
  mkdirSync(directory, { recursive: true });

  const taken = new Set<string>();
  return paths.map((each) => {
    const source = resolve(each);
    if (!existsSync(source)) throw new Error(`no such document: ${source}`);

    const stem =
      basename(source)
        .replace(/\.docx$/i, "")
        .replace(/[^a-z0-9]+/gi, "-")
        .replace(/^-|-$/g, "")
        .toLowerCase() || "document";

    // Two documents in different directories can share a name, and the second
    // would otherwise draw over the first's pages without saying so.
    let id = stem;
    for (let n = 2; taken.has(id); n += 1) id = `${stem}-${String(n)}`;
    taken.add(id);

    const path = resolve(directory, `${id}.docx`);
    if (path !== source) copyFileSync(source, path);
    return { id, path };
  });
}

const NOTHING_MEASURED = {
  tolerancePt: 0.5,
  textTolerancePt: 1,
  bodyTopPt: null,
  headerTopsPt: [],
  bodyTopsPt: [],
  headerFloatCount: null,
  leastBodyFloatCount: null,
  floatsPt: [],
  inlinesPt: [],
  disjointFloatPairs: [],
  renderedImagesPt: null,
  renderedPageIndexes: null,
  picturesWordDrewWithoutAnImage: 0,
  imagesWordDrewOutsideAPicture: 0,
  textLinesMatched: null,
  textLinesPlaced: null,
  textRunsMatched: null,
  textRunsPlaced: null,
  unrenderablePictures: 0,
  unknownDrawings: 0,
  metafileFills: null,
  metafileRuns: null,
  numbersMatched: null,
  numbersPlaced: null,
  unhonoured: [],
} as const;

// Nothing is pinned against a document handed over on the command line, so every
// number a reference case carries is left empty. Word's pdf is beside it or it is
// not, and that is the whole of what is known.
export const caseOf = (id: string, path: string): ReferenceCase => {
  const rendered = path.replace(/\.docx$/i, ".pdf");
  return {
    ...NOTHING_MEASURED,
    id,
    documentPath: path,
    renderedPath: existsSync(rendered) ? rendered : null,
  };
};
