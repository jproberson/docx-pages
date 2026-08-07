import { type Block, type Paragraph } from "@docx-pages/core";
import { readAnchors } from "@docx-pages/core/internal";

// Which paragraph of ours each of Word's answers is about.
//
// Word counts more paragraphs than a document has `w:p` elements: every row of a
// table ends with a mark of its own, which Word counts and this project never
// lays out. And a paragraph whose range takes in the mark ending its cell does
// not answer for itself. Asked where it is, Word gives the origin of the row
// that holds it, so the last paragraph of every cell reports where the row's
// first cell began and the row's own mark reports the same. A paragraph with
// another after it in the same cell keeps the mark out of its range and answers
// for itself.
//
// Word measures the horizontal answer from the cell rather than from the text
// column, which makes it nought for every paragraph in a table whatever the cell
// holds, so there is nothing there to compare against. Word's own pdf is the
// oracle for where a cell put its text.
//
// The horizontal answer is where the paragraph's own text starts, which is not
// where a line an object narrowed began: Word reports nought for a line it drew
// three hundred points across. So a story holding anything text wraps round is not
// compared on the left at all, and the pdf answers for those too.

export type Answer = {
  // The paragraph of ours whose place Word is reporting.
  readonly paragraph: number;
  // The paragraph of ours whose page Word is reporting, which is not always the
  // same one: Word answers for a paragraph's start and for its active end at once,
  // and where the two ends fell in different places the answer mixes them. The
  // last paragraph of a cell reports the row's origin and its own page, so a row a
  // break was torn through answers from the page it opened on with the page it
  // finished on.
  readonly endsAt: number;
  // Whether the horizontal answer is measured from anywhere this project knows.
  readonly comparesLeft: boolean;
};

export function answeringParagraphs(blocks: readonly Block[]): readonly Answer[] {
  const answers: Answer[] = [];
  walk(blocks, answers, null);
  if (!wrapsText(blocks)) return answers;
  return answers.map((answer) => ({ ...answer, comparesLeft: false }));
}

const wrapsText = (blocks: readonly Block[]): boolean =>
  blocks.some((block) =>
    block.kind === "paragraph"
      ? anchorsWrap(block.paragraph)
      : block.rows.some((row) => row.cells.some((cell) => wrapsText(cell.blocks))),
  );

const anchorsWrap = (paragraph: Paragraph): boolean =>
  readAnchors(paragraph).some((anchor) => anchor.wrap !== "none");

// `endingAnswer` is the paragraph the cell's own last paragraph answers with, or
// null in a run of blocks no cell holds.
function walk(blocks: readonly Block[], answers: Answer[], endingAnswer: number | null): void {
  const last = lastParagraphIndex(blocks);

  for (const block of blocks) {
    if (block.kind === "paragraph") {
      const index = block.paragraph.index;
      const answering = index === last ? (endingAnswer ?? index) : index;
      answers.push({ paragraph: answering, endsAt: index, comparesLeft: endingAnswer === null });
      continue;
    }

    for (const row of block.rows) {
      const cells = row.cells.flatMap((cell) => cell.blocks);
      const opener = firstParagraphIndex(cells);
      for (const cell of row.cells) walk(cell.blocks, answers, opener);
      // The mark that ends a row stands where the row's text starts, and ends where
      // the row's last paragraph does. A row with nothing in it to answer for still
      // counts, so that the answers after it stay lined up with the paragraphs they
      // are about.
      answers.push({
        paragraph: opener ?? -1,
        endsAt: lastParagraphIndex(cells) ?? -1,
        comparesLeft: false,
      });
    }
  }
}

const paragraphIndexes = (blocks: readonly Block[]): readonly number[] =>
  blocks.flatMap((block) =>
    block.kind === "paragraph"
      ? [block.paragraph.index]
      : block.rows.flatMap((row) => row.cells.flatMap((cell) => paragraphIndexes(cell.blocks))),
  );

const firstParagraphIndex = (blocks: readonly Block[]): number | null =>
  paragraphIndexes(blocks)[0] ?? null;

const lastParagraphIndex = (blocks: readonly Block[]): number | null =>
  paragraphIndexes(blocks).at(-1) ?? null;
