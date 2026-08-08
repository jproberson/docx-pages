import { existsSync } from "node:fs";

import type { ReferenceCase } from "../testing/cases.js";
import { authoredDocuments } from "./documents.js";
import { hasStatedFaces } from "./faces.js";
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
  // them names a rule the layout has measured and does not answer. Every number
  // here is held to exactly, so it can no more quietly grow than quietly fall:
  // answering one line more is a rule now met and is worth saying so.
  readonly placed?: number;
  readonly runsPlaced?: number;
};

const DRAWN: Readonly<Record<string, Drawn>> = {
  // Every line, tabs and all. Word draws a tab as a space stretched to the width
  // of the gap it opened, though the line holds no character for it, which is why
  // a line here is spelled out of what carries ink rather than of every character
  // drawn on it.
  tabs: { lines: 16, runs: 42, numbers: 0 },
  tables: { lines: 27, runs: 27, numbers: 0 },
  spacing: { lines: 26, runs: 26, numbers: 0 },
  wrapping: { lines: 7, runs: 7, numbers: 0 },
  numbering: { lines: 10, runs: 10, numbers: 7 },
  pages: { lines: 24, runs: 24, numbers: 0 },
  // Fifteen cases of four rows and two cells, and the line either side of each
  // table: 150 lines in both documents, since the compatibility mode changes
  // nothing about any of them.
  "lined-rows": { lines: 150, runs: 150, numbers: 0 },
  "legacy-lined-rows": { lines: 150, runs: 150, numbers: 0 },
  // Every line of all eight cases, the fillers holding them down the page
  // included: a case is nine lines of its own and the row it asks about.
  tearing: { lines: 154, runs: 154, numbers: 0 },
  // Twenty lines out of thirty two paragraphs: the twelve holding nothing but a
  // space or a tab draw no ink at all, which is the whole of what the document is
  // asking about and the reason Word's report is the oracle for it. The three runs
  // over twenty lines are the word each side of the wide space in case c.
  "trailing-space": { lines: 20, runs: 23, numbers: 0 },
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
  // Every line, including the three Word drew in pieces around a space of its own
  // and the one where it drew a glyph the pdf can put no character to. What each
  // of those characters is remains the other suite's question, asked of where the
  // letters either side of it sit.
  "unmapped-characters": { lines: 24, runs: 26, numbers: 0 },
  // Nine tables and the nine lines between them, in both documents: an indent is
  // measured to the text in an old document and to the table's edge in a modern
  // one, and every case of it now lands where Word drew it either way.
  "legacy-table-indent": { lines: 18, runs: 18, numbers: 0 },
  "table-indent": { lines: 18, runs: 18, numbers: 0 },
  // Every line but the seven under the wave borders, which Word draws neither at
  // the width they state nor at any width read off it: the rows they line come out
  // shorter here and everything below them stands too high.
  borders: { lines: 104, runs: 104, numbers: 0, placed: 97, runsPlaced: 102 },
  "section-pages": { lines: 18, runs: 12, numbers: 0, placed: 12 },
  // Fifty four lines over eleven cases: a marker, a shim, the paragraph anchoring
  // the object, the line under it, and the object's own line, which is what says
  // which page the object landed on. The case with no line under it is the one
  // asking what becomes of an object anchored to the last paragraph there is.
  //
  // The one left over is the line under the object that was drawn up to the foot of
  // the text, which Word puts beside the object where it now stands and this project
  // puts beside where it was asked for. **The room a line is left is settled while
  // the story is measured, and an object is drawn up to a foot no page has yet**:
  // measuring is one column with no bottom, so nothing there can know how far past
  // the foot of a page an object 2000pt down the column reaches. Word's own report
  // cannot see it either, since it answers nought for the left of any line an object
  // narrowed, so only this catches it.
  "objects-past-the-foot": { lines: 54, runs: 54, numbers: 0, placed: 53, runsPlaced: 53 },
  // Every line of all eleven cases: a marker, four rows of two cells and a line
  // under the table. Word's own pdf is the oracle this document is read by, since
  // its report puts the last row of four of the cases 0.55 to 0.7pt below where its
  // own drawing has it.
  "stated-row-heights": { lines: 110, runs: 110, numbers: 0 },
  // Eighteen lines out of twenty two paragraphs: the four closing a section with
  // nothing in them draw no ink at all, which is the whole of what the document is
  // asking about. Every drawn line of it lands where Word drew it, which is worth
  // more here than Word's own report: the report answers for a closer with the top
  // of the paragraph above it, and gives the same answer whether the closer fits on
  // its page or not.
  "section-closer": { lines: 18, runs: 18, numbers: 0 },
  // The three documents whose sections make different pages, and the oracle they
  // were written to be read by: each line of them is drawn where its own section's
  // margins put it, which Word's report cannot say at all. The report answers for a
  // paragraph's left from its own section's text boundary, so it is nought down
  // every one of these however far across the page Word drew the line.
  sections: { lines: 8, runs: 8, numbers: 0 },
  "section-flow": { lines: 8, runs: 8, numbers: 0 },
  // Two pages of a continuous section's own text past the page it opened on, and
  // the two lines of the body's own section under them.
  "overflowing-section": { lines: 66, runs: 66, numbers: 0 },
  // Seven cases of a marker, four cells and the lines the table left standing, which
  // the last case writes long enough to reach past the table and take two lines each.
  //
  // Three of the 61 are the case whose table leaves a run of the frame wide enough
  // for a line on both sides of itself: Word draws one line and fills both runs with
  // it, and a line here stands in one run or the other. Nothing in the wild leaves a
  // run that wide on the near side.
  "positioned-table": { lines: 61, runs: 62, numbers: 0, placed: 58, runsPlaced: 60 },
  // Eight cases: a marker and the case written out three times, and ten lines more
  // in each of the two whose page is opened by the one below it filling up. This is
  // the oracle the document was written for, since where a line was drawn is the
  // whole of what it asks.
  "space-above-a-break": { lines: 52, runs: 52, numbers: 0 },
  // Seven cases of four lines each, and the line inside each case's own box.
  "space-under-a-wrap": { lines: 35, runs: 35, numbers: 0 },
  // Twenty cases of five lines: a marker, the case written out three times and a
  // plain line under it. The thirty six runs over the hundred lines are the moved
  // run of each case that stands beside a plain one, which Word draws as an item of
  // its own. This is the oracle the document was written for, since a raise is a
  // position and Word's report answers for the paragraph rather than the run.
  "raised-text": { lines: 100, runs: 136, numbers: 0 },
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

// Images in Word's pdf that are no picture of the document's. Word draws a glyph
// its face has no outline for through a bitmap font of its own making, and a
// bitmap glyph reaches the pdf as an image: the three here are one character
// written out three times, each 12x12 standing on its line's baseline, and Word
// writes the character as text in the same place as well.
const IMAGES_THAT_ARE_GLYPHS: Readonly<Record<string, number>> = {
  "unmapped-in-a-text-face": 3,
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
  renderedImagesPt: null,
  renderedPageIndexes: null,
  picturesWordDrewWithoutAnImage: 0,
  imagesWordDrewOutsideAPicture: 0,
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
    // A document this project refuses cannot be compared against anything. It is
    // still written and still asked about; it is only kept out of the suites that
    // need a page to compare, until the rule it asks about is built.
    if (each.refuses !== undefined) return [];
    // A document naming a face this machine has not got would be laid out in
    // another one, which answers a different question.
    if (!hasStatedFaces(each.statedFaces ?? [])) return [];
    return [
      {
        ...EMPTY,
        ...drawnBy(each.id),
        ...(UNHONOURED[each.id] === undefined ? {} : { unhonoured: UNHONOURED[each.id] }),
        imagesWordDrewOutsideAPicture: IMAGES_THAT_ARE_GLYPHS[each.id] ?? 0,
        id: `authored-${each.id}`,
        documentPath,
        renderedPath: renderedPath(each.id),
        tolerancePt: 0.5,
        textTolerancePt: 1,
      },
    ];
  });
}
