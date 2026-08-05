import type { ParagraphBox } from "./stack.js";

export type PageStack = {
  readonly index: number;
  readonly boxes: readonly ParagraphBox[];
};

export type BreakStackInput = {
  // A stack measured from `topPt` with no bottom, as `measureStack` produces it.
  readonly boxes: readonly ParagraphBox[];
  readonly topPt: number;
  readonly bottomPt: number;
};

// Every page repeats the one body box, so a page is the stack shifted back up by
// however much of it came before: what carried on to the next page starts again at
// the body's top. What a paragraph put above its first line, its space before, is
// what the break leaves behind.
export function breakStack(input: BreakStackInput): readonly PageStack[] {
  const pages: ParagraphBox[][] = [[]];
  let shiftPt = 0;

  const put = (box: ParagraphBox): void => {
    pages[pages.length - 1]?.push(box);
  };

  const overflows = (topPt: number, heightPt: number): boolean =>
    topPt - shiftPt + heightPt > input.bottomPt && topPt - shiftPt > input.topPt;

  for (const box of input.boxes) {
    if (box.lines.length === 0) {
      if (overflows(box.topPt, box.heightPt)) {
        shiftPt = box.topPt - input.topPt;
        pages.push([]);
      }
      put(partOf(box, 0, 0, shiftPt));
      continue;
    }

    let from = 0;
    let at = 0;
    while (at < box.lines.length) {
      const line = box.lines[at];
      if (line === undefined || !overflows(line.topPt, line.heightPt)) {
        at += 1;
        continue;
      }

      // The lines between the cut and the line that overflowed are on the next
      // page now, so they are looked at again from there.
      const cut = cutFor(box, { from, at, shiftPt, topPt: input.topPt });
      if (cut > from) put(partOf(box, from, cut, shiftPt));
      shiftPt = (box.lines[cut]?.topPt ?? line.topPt) - input.topPt;
      pages.push([]);
      from = cut;
      at = cut + 1;
    }

    put(partOf(box, from, box.lines.length, shiftPt));
  }

  return pages.map((boxes, index) => ({ index, boxes }));
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
    widowControl: box.widowControl,
    contentWidthPt: box.contentWidthPt,
  };
}
