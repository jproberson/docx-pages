import type { Block } from "@docx-pages/core";

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

export type Answer = {
  // The paragraph of ours whose place Word is reporting.
  readonly paragraph: number;
  // Whether the horizontal answer is measured from anywhere this project knows.
  readonly comparesLeft: boolean;
};

export function answeringParagraphs(blocks: readonly Block[]): readonly Answer[] {
  const answers: Answer[] = [];
  walk(blocks, answers, null);
  return answers;
}

// `endingAnswer` is the paragraph the cell's own last paragraph answers with, or
// null in a run of blocks no cell holds.
function walk(blocks: readonly Block[], answers: Answer[], endingAnswer: number | null): void {
  const last = lastParagraphIndex(blocks);

  for (const block of blocks) {
    if (block.kind === "paragraph") {
      const index = block.paragraph.index;
      const answering = index === last ? (endingAnswer ?? index) : index;
      answers.push({ paragraph: answering, comparesLeft: endingAnswer === null });
      continue;
    }

    for (const row of block.rows) {
      const opener = firstParagraphIndex(row.cells.flatMap((cell) => cell.blocks));
      for (const cell of row.cells) walk(cell.blocks, answers, opener);
      // The mark that ends a row stands where the row's text starts. A row with
      // nothing in it to answer for still counts, so that the answers after it
      // stay lined up with the paragraphs they are about.
      answers.push({ paragraph: opener ?? -1, comparesLeft: false });
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
