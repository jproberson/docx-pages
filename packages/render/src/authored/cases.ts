import { existsSync } from "node:fs";

import type { ReferenceCase } from "../testing/cases.js";
import { authoredDocuments } from "./documents.js";
import { authoredPath } from "./write.js";

// The authored documents, dressed as reference cases so that everything already
// built for the seven real ones works on them too: the preview beside Word's own
// pdf, the text comparison, the fills and the pictures.
//
// Where every line of one landed is pinned here, against Word's own rendering of
// it. That is not what `layout.reference.test.ts` asks, which is where Word says
// it put each paragraph: a table is the plainest case of the difference, since
// Word will only answer for the row and its pdf draws every cell.

const renderedPath = (id: string): string | null => {
  const path = authoredPath(id).replace(/\.docx$/, ".pdf");
  return existsSync(path) ? path : null;
};

// How much of each document Word's pdf agrees with: every line of it, every run
// of every line, and every number in front of one. A document left out of this
// draws the same text over and over in boxes standing on each other, and which
// line of it a run of items in the pdf belongs to is then anyone's guess.
type Drawn = {
  readonly lines: number;
  readonly runs: number;
  readonly numbers: number;
  // How many of those lines and runs land where Word drew them. Short of all of
  // them names a rule the layout has measured and does not answer, which is then
  // a number that cannot quietly grow.
  readonly placed?: number;
  readonly runsPlaced?: number;
};

const DRAWN: Readonly<Record<string, Drawn>> = {
  tabs: { lines: 13, runs: 37, numbers: 0 },
  tables: { lines: 27, runs: 27, numbers: 0 },
  spacing: { lines: 19, runs: 19, numbers: 0 },
  wrapping: { lines: 7, runs: 7, numbers: 0 },
  numbering: { lines: 10, runs: 10, numbers: 7 },
  pages: { lines: 24, runs: 24, numbers: 0 },
  // Three lines a case, five where the paragraph above the object holds three,
  // and two where it holds none at all. Two of them stand beside the narrow
  // object, which is centred in the column and wrapped on its largest side, so
  // which of the two equal sides they take is the whole of what the two documents
  // say differently here.
  "legacy-wrapping": { lines: 21, runs: 21, numbers: 0 },
  "modern-wrapping": { lines: 21, runs: 21, numbers: 0 },
  // Four lines a case: one naming it and three beside the object.
  "wrap-sides": { lines: 28, runs: 28, numbers: 0 },
  "legacy-wrap-sides": { lines: 28, runs: 28, numbers: 0 },
  // Every line but the seven under the wave borders, which Word draws neither at
  // the width they state nor at any width read off it: the rows they line come out
  // shorter here and everything below them stands too high.
  borders: { lines: 104, runs: 104, numbers: 0, placed: 97, runsPlaced: 102 },
};

// What each authored document is expected to say it passed over. A document that
// asks about a feature and states one it does not read says so here; the rest say
// nothing.
const UNHONOURED: Readonly<Record<string, readonly string[]>> = {
  // A bar stop draws a line down the page, which nothing here draws.
  tabs: ["bar-tab-stop"],
  // The wave borders, which are drawn as plain lines of the stated width.
  borders: ["approximated-border"],
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
  unrenderablePictures: 0,
  unknownDrawings: 0,
  metafileFills: null,
  metafileRuns: null,
  unhonoured: [],
} as const;

const drawnBy = (
  id: string,
): Pick<
  ReferenceCase,
  | "textLinesMatched"
  | "textLinesPlaced"
  | "textRunsMatched"
  | "textRunsPlaced"
  | "numbersMatched"
  | "numbersPlaced"
> => {
  const drawn = DRAWN[id];
  if (drawn === undefined) {
    return {
      textLinesMatched: null,
      textLinesPlaced: null,
      textRunsMatched: null,
      textRunsPlaced: null,
      numbersMatched: null,
      numbersPlaced: null,
    };
  }
  return {
    textLinesMatched: drawn.lines,
    textLinesPlaced: drawn.placed ?? drawn.lines,
    textRunsMatched: drawn.runs,
    textRunsPlaced: drawn.runsPlaced ?? drawn.runs,
    numbersMatched: drawn.numbers,
    numbersPlaced: drawn.numbers,
  };
};

export function authoredCases(): readonly ReferenceCase[] {
  return authoredDocuments().flatMap((each) => {
    const documentPath = authoredPath(each.id);
    if (!existsSync(documentPath)) return [];
    return [
      {
        ...EMPTY,
        ...drawnBy(each.id),
        ...(UNHONOURED[each.id] === undefined ? {} : { unhonoured: UNHONOURED[each.id] }),
        id: `authored-${each.id}`,
        documentPath,
        renderedPath: renderedPath(each.id),
        tolerancePt: 0.5,
        textTolerancePt: 1,
      },
    ];
  });
}
