import { existsSync } from "node:fs";

import type { ReferenceCase } from "../testing/cases.js";
import { authoredDocuments } from "./documents.js";
import { authoredPath } from "./write.js";

// The authored documents, dressed as reference cases so that everything already
// built for the seven real ones works on them too: the preview beside Word's own
// pdf, the text comparison, the fills and the pictures.
//
// They carry no expectations of their own here. What each one is worth is pinned
// in `layout.reference.test.ts` against Word's answers about the document rather
// than against its rendering, and this is for looking at: a number that is out and
// a page that looks wrong are the same fault seen twice, and seeing it is what says
// which one to believe.

const renderedPath = (id: string): string | null => {
  const path = authoredPath(id).replace(/\.docx$/, ".pdf");
  return existsSync(path) ? path : null;
};

const EMPTY = {
  bodyTopPt: null,
  headerTopsPt: [],
  bodyTopsPt: [],
  headerFloatCount: null,
  leastBodyFloatCount: null,
  floatsPt: [],
  inlinesPt: [],
  disjointFloatPairs: [],
  renderedImagesPt: [],
  renderedPageIndexes: [],
  textLinesMatched: null,
  textLinesPlaced: null,
  textRunsMatched: null,
  textRunsPlaced: null,
  unrenderablePictures: 0,
  metafileFills: null,
  metafileRuns: null,
  numbersMatched: null,
  numbersPlaced: null,
} as const;

export function authoredCases(): readonly ReferenceCase[] {
  return authoredDocuments().flatMap((each) => {
    const documentPath = authoredPath(each.id);
    if (!existsSync(documentPath)) return [];
    return [
      {
        ...EMPTY,
        id: `authored-${each.id}`,
        documentPath,
        renderedPath: renderedPath(each.id),
        tolerancePt: 0.5,
        textTolerancePt: 1,
      },
    ];
  });
}
