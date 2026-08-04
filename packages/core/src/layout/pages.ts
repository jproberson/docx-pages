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
    box.lines.forEach((line, at) => {
      if (!overflows(line.topPt, line.heightPt)) return;
      if (at > from) put(partOf(box, from, at, shiftPt));
      shiftPt = line.topPt - input.topPt;
      pages.push([]);
      from = at;
    });

    put(partOf(box, from, box.lines.length, shiftPt));
  }

  return pages.map((boxes, index) => ({ index, boxes }));
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
  };
}
