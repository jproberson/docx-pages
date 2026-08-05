import type { ParagraphBox, PlacedCell } from "./stack.js";

export type PageStack = {
  readonly index: number;
  readonly boxes: readonly ParagraphBox[];
  readonly cells: readonly PlacedCell[];
};

export type BreakStackInput = {
  // A stack measured from `topPt` with no bottom, as `measureStack` produces it.
  readonly boxes: readonly ParagraphBox[];
  readonly cells: readonly PlacedCell[];
  readonly topPt: number;
  readonly bottomPt: number;
};

// Every page repeats the one body box, so a page is the stack shifted back up by
// however much of it came before: what carried on to the next page starts again at
// the body's top. What a paragraph put above its first line, its space before, is
// what the break leaves behind.
export function breakStack(input: BreakStackInput): readonly PageStack[] {
  const pages: ParagraphBox[][] = [[]];
  // Where in the stack each page started, which is what the cells are cut by
  // once the text has said where the pages fall.
  const shifts: number[] = [0];
  let shiftPt = 0;

  const put = (box: ParagraphBox): void => {
    pages[pages.length - 1]?.push(box);
  };

  const open = (): void => {
    shifts.push(shiftPt);
    pages.push([]);
  };

  const overflows = (topPt: number, heightPt: number): boolean =>
    topPt - shiftPt + heightPt > input.bottomPt && topPt - shiftPt > input.topPt;

  // The page is left where the break asked, so what comes after it starts again at
  // the body's top. A break with nothing left to move makes no page: a paragraph
  // asking for one of its own while already standing at the top of one stays where
  // it is, which is what keeps the first paragraph of a document off page two.
  const leave = (topPt: number): boolean => {
    if (topPt - shiftPt <= input.topPt) return false;
    shiftPt = topPt - input.topPt;
    open();
    return true;
  };

  // Whether the paragraph before this one ended on a page break, which draws no
  // line of its own and so has nothing here to be seen at.
  let broken = false;

  for (const box of input.boxes) {
    if (broken || box.startsPage) leave(box.lines[0]?.topPt ?? box.topPt);
    broken = box.endsPage;

    // A paragraph with nothing in it is judged by the room its mark stands in, as
    // one with lines is judged by its lines: the room it keeps below itself hangs
    // past the foot of the page rather than moving it on.
    if (box.lines.length === 0) {
      if (overflows(box.topPt, box.contentBottomPt - box.topPt)) {
        shiftPt = box.topPt - input.topPt;
        open();
      }
      put(partOf(box, 0, 0, shiftPt));
      continue;
    }

    let from = 0;
    let at = 0;
    while (at < box.lines.length) {
      const line = box.lines[at];
      if (line === undefined) {
        at += 1;
        continue;
      }
      // A line a break of the paragraph's own put at the head of a page goes there
      // whatever room was left below, and widow control has no say in where a break
      // the document asked for falls.
      const asked = line.startsPage;
      if (!asked && !overflows(line.topPt, line.heightPt)) {
        at += 1;
        continue;
      }

      // The lines between the cut and the line that overflowed are on the next
      // page now, so they are looked at again from there.
      const cut = asked ? at : cutFor(box, { from, at, shiftPt, topPt: input.topPt });
      if (cut > from) put(partOf(box, from, cut, shiftPt));
      shiftPt = (box.lines[cut]?.topPt ?? line.topPt) - input.topPt;
      open();
      from = cut;
      at = cut + 1;
    }

    put(partOf(box, from, box.lines.length, shiftPt));
  }

  return pages.map((boxes, index) => ({
    index,
    boxes,
    cells: cellsOn(input, shifts, index),
  }));
}

// A cell is cut by the pages the text broke into rather than breaking them: the
// piece of one standing on a page is what it covers between where that page
// started in the stack and where the next page did.
function cellsOn(
  input: BreakStackInput,
  shifts: readonly number[],
  index: number,
): readonly PlacedCell[] {
  const shiftPt = shifts[index] ?? 0;
  const fromPt = input.topPt + shiftPt;
  const next = shifts[index + 1];
  const toPt = Math.min(
    next === undefined ? Number.POSITIVE_INFINITY : input.topPt + next,
    fromPt + (input.bottomPt - input.topPt),
  );

  return input.cells.flatMap((cell) => {
    const topPt = Math.max(cell.topPt, fromPt);
    const bottomPt = Math.min(cell.topPt + cell.heightPt, toPt);
    if (bottomPt - topPt <= 0) return [];
    return [{ ...cell, topPt: topPt - shiftPt, heightPt: bottomPt - topPt }];
  });
}

type Cut = {
  // The first of the paragraph's lines still to be placed, and the one whose own
  // height took it over the bottom of the page.
  readonly from: number;
  readonly at: number;
  readonly shiftPt: number;
  readonly topPt: number;
};

// Word will not leave a paragraph's first line alone at the foot of a page, nor
// its last line alone at the top of the next one: `w:widowControl` moves a break
// that would do either back a line, and takes the whole paragraph over when a
// single line is all that would be left above it. A break with nowhere to move to
// is left where it fell, since a page that carries nothing forward never ends.
function cutFor(box: ParagraphBox, cut: Cut): number {
  if (!box.widowControl) return cut.at;

  const alone = cut.at === box.lines.length - 1 ? cut.at - 1 : cut.at;
  const moved = cut.from === 0 && alone < 2 ? 0 : alone;
  const shiftPt = (box.lines[moved]?.topPt ?? 0) - cut.topPt;
  return moved < cut.from || shiftPt <= cut.shiftPt ? cut.at : moved;
}

// A paragraph the break ran through keeps its index on both pages: it is the one
// paragraph, drawn in two places. Only the part holding its first line carries the
// number the list put in front of it.
function partOf(box: ParagraphBox, from: number, to: number, shiftPt: number): ParagraphBox {
  const lines = box.lines.slice(from, to);
  const first = lines[0];
  const last = lines[lines.length - 1];
  const topPt = (from === 0 || first === undefined ? box.topPt : first.topPt) - shiftPt;
  const bottomPt =
    to === box.lines.length || last === undefined
      ? box.topPt + box.heightPt - shiftPt
      : last.topPt + last.heightPt - shiftPt;

  return {
    index: box.index,
    topPt,
    anchorTopPt: box.anchorTopPt - shiftPt,
    heightPt: bottomPt - topPt,
    lines: lines.map((line) => ({
      ...line,
      topPt: line.topPt - shiftPt,
      baselinePt: line.baselinePt - shiftPt,
    })),
    marker:
      box.marker === null || from > 0
        ? null
        : { ...box.marker, baselinePt: box.marker.baselinePt - shiftPt },
    // The mark stands at the end of the paragraph, so it goes with the part that
    // holds the last of it.
    markTopPt: to === box.lines.length ? box.markTopPt - shiftPt : bottomPt,
    contentBottomPt:
      to === box.lines.length || last === undefined
        ? box.contentBottomPt - shiftPt
        : last.topPt + last.heightPt - shiftPt,
    widowControl: box.widowControl,
    // What the paragraph asked of the pages either side of it belongs to the part
    // of it that stands there.
    startsPage: from === 0 && box.startsPage,
    endsPage: to === box.lines.length && box.endsPage,
    contentWidthPt: box.contentWidthPt,
    clipTo: box.clipTo === null ? null : { ...box.clipTo, topPt: box.clipTo.topPt - shiftPt },
    paint: box.paint,
  };
}
