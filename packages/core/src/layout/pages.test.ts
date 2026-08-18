import { describe, expect, it } from "vitest";

import { breakStack } from "./pages.js";
import type { ParagraphBox, ParagraphMarker, PlacedLine } from "./stack.js";

const MARK = {
  font: { kind: "named", name: "Meridian Sans" },
  fontSizePt: 10,
  bold: false,
  italic: false,
  underline: false,
  raisePt: 0,
  lineSizePt: 10,
  lineRaisePt: 0,
  color: null,
  characterSpacingPt: 0,
  characterScale: 1,
  kernFromHalfPoints: null,
  highlight: null,
  capitals: "none",
} as const;

const line = (topPt: number, heightPt: number, text: string): PlacedLine => ({
  line: {
    segments: [{ kind: "text", mark: MARK, text, widthPt: 20, offsetPt: 0 }],
    widthPt: 20,
    heightPt,
    ascentPt: heightPt * 0.8,
    seatPt: 0,
    fontHeightPt: heightPt,
    heldOpenPt: null,
  },
  leftPt: 0,
  topPt,
  heightPt,
  seatPt: 0,
  fittingHeightPt: heightPt,
  baselinePt: topPt + heightPt * 0.8,
  startsPage: false,
});

const MARKER: ParagraphMarker = {
  text: "1.",
  mark: MARK,
  widthPt: 6,
  leftPt: 0,
  baselinePt: 0,
};

// Paragraphs are laid down one after another, each as tall as the lines it holds,
// and each willing to be broken through unless the test says otherwise.
function stack(
  shape: readonly (readonly number[])[],
  topPt = 100,
  widowControl = false,
): readonly ParagraphBox[] {
  const boxes: ParagraphBox[] = [];
  let top = topPt;

  shape.forEach((heights, index) => {
    const lines: PlacedLine[] = [];
    let lineTop = top;
    for (const heightPt of heights) {
      lines.push(line(lineTop, heightPt, `p${String(index)}l${String(lines.length)}`));
      lineTop += heightPt;
    }
    boxes.push({
      index,
      topPt: top,
      anchorTopPt: top,
      resumesUnderPt: 0,
      heightPt: lineTop - top,
      lines,
      marker: null,
      markTopPt: lines[lines.length - 1]?.topPt ?? top,
      contentBottomPt: lineTop,
      widowControl,
      keepNext: false,
      startsPage: false,
      endsPage: false,
      endsPageAtASection: false,
      contentWidthPt: 0,
      clipTo: null,
      paint: null,
    });
    top = lineTop;
  });

  return boxes;
}

const cellAt = (topPt: number, heightPt: number) => ({
  leftPt: 100,
  topPt,
  widthPt: 72,
  heightPt,
  fillColor: null,
  borders: { top: null, left: null, bottom: null, right: null },
});

const linesOn = (
  page: { readonly boxes: readonly ParagraphBox[] } | undefined,
): readonly number[] => (page?.boxes ?? []).map((box) => box.lines.length);

const indexesOn = (page: { readonly boxes: readonly ParagraphBox[] }): readonly number[] =>
  page.boxes.map((box) => box.index);

// What a document asking for its own page breaks hands the stack: a paragraph
// that asked to start a page, one whose text ran on past a break of its own, and
// one that ended on a break with nothing after it to carry over.
const asking = (
  boxes: readonly ParagraphBox[],
  at: number,
  asks: {
    readonly startsPage?: boolean;
    readonly endsPage?: boolean;
    readonly endsPageAtASection?: boolean;
    readonly line?: number;
  },
): readonly ParagraphBox[] =>
  boxes.map((box, index) =>
    index !== at
      ? box
      : {
          ...box,
          startsPage: asks.startsPage ?? box.startsPage,
          endsPage: asks.endsPage ?? box.endsPage,
          endsPageAtASection: asks.endsPageAtASection ?? box.endsPageAtASection,
          lines: box.lines.map((each, line) =>
            line === asks.line ? { ...each, startsPage: true } : each,
          ),
        },
  );

describe("breakStack", () => {
  it("leaves a stack that fits on the one page", () => {
    const pages = breakStack({
      cells: [],
      boxes: stack([[10], [10], [10]]),
      topPt: 100,
      bottomPt: 200,
    });

    expect(pages).toHaveLength(1);
    expect(indexesOn(pages[0] ?? { boxes: [] })).toStrictEqual([0, 1, 2]);
  });

  // Measured on 2026-08-11 by the authored `twip-grid` document: one of its cases
  // keeps 39 lines whose boxes come to 720.46 in a body of 720, and its last line
  // ends 8.8pt above the foot with the room its multiple opened hanging past it.
  it("keeps a line whose text fits though the room its rule opened below does not", () => {
    const boxes = stack([[10], [10]]).map((box, at) =>
      at === 1
        ? {
            ...box,
            lines: box.lines.map((each) => ({ ...each, heightPt: 20, fittingHeightPt: 10 })),
          }
        : box,
    );
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 125 });

    expect(pages).toHaveLength(1);
    expect(indexesOn(pages[0] ?? { boxes: [] })).toStrictEqual([0, 1]);
  });

  it("moves a line whose own text crosses the foot however little the rule opened", () => {
    const boxes = stack([[10], [10]]).map((box, at) =>
      at === 1
        ? {
            ...box,
            lines: box.lines.map((each) => ({ ...each, heightPt: 20, fittingHeightPt: 20 })),
          }
        : box,
    );
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 125 });

    expect(pages).toHaveLength(2);
  });

  it("moves the paragraph that would cross the bottom onto the next page", () => {
    const pages = breakStack({
      cells: [],
      boxes: stack([[10], [10], [10]]),
      topPt: 100,
      bottomPt: 125,
    });

    expect(pages).toHaveLength(2);
    expect(indexesOn(pages[0] ?? { boxes: [] })).toStrictEqual([0, 1]);
    expect(indexesOn(pages[1] ?? { boxes: [] })).toStrictEqual([2]);
  });

  it("starts the next page at the top of the body rather than where the flow was", () => {
    const pages = breakStack({
      cells: [],
      boxes: stack([[10], [10], [10]]),
      topPt: 100,
      bottomPt: 125,
    });

    expect(pages[1]?.boxes[0]?.topPt).toBe(100);
    expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(100);
    expect(pages[1]?.boxes[0]?.lines[0]?.baselinePt).toBe(108);
  });

  it("breaks a paragraph between its own lines and carries the rest over", () => {
    const pages = breakStack({
      cells: [],
      boxes: stack([[10, 10, 10]]),
      topPt: 100,
      bottomPt: 120,
    });

    expect(pages).toHaveLength(2);
    expect(pages[0]?.boxes[0]?.lines).toHaveLength(2);
    expect(pages[1]?.boxes[0]?.index).toBe(0);
    expect(pages[1]?.boxes[0]?.lines).toHaveLength(1);
    expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(100);
  });

  it("keeps carrying paragraphs onto further pages for as long as the stack runs", () => {
    const pages = breakStack({
      cells: [],
      boxes: stack([[10], [10], [10], [10], [10], [10]]),
      topPt: 100,
      bottomPt: 120,
    });

    expect(pages.map(indexesOn)).toStrictEqual([
      [0, 1],
      [2, 3],
      [4, 5],
    ]);
  });

  it("gives the number to the page that holds the line it belongs in front of", () => {
    const boxes = stack([[10, 10]]).map((box) => ({
      ...box,
      marker: { ...MARKER, baselinePt: 108 },
    }));
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 105 });

    expect(pages[0]?.boxes[0]?.marker?.baselinePt).toBe(108);
    expect(pages[1]?.boxes[0]?.marker).toBeNull();
  });

  it("draws a line taller than the page rather than looking for a page it fits", () => {
    const pages = breakStack({ cells: [], boxes: stack([[400]]), topPt: 100, bottomPt: 200 });

    expect(pages).toHaveLength(1);
    expect(pages[0]?.boxes[0]?.lines).toHaveLength(1);
  });

  it("takes a paragraph over whole rather than leaving its first line at the foot", () => {
    const boxes = stack([[10], [10, 10]], 100, true);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 125 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1]]);
    expect(pages.map(linesOn)).toStrictEqual([[1], [2]]);
    expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(100);
  });

  it("moves the line above the break down with it rather than leaving the last line alone", () => {
    const boxes = stack([[10], [10, 10, 10, 10]], 100, true);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 145 });

    expect(pages.map(indexesOn)).toStrictEqual([[0, 1], [1]]);
    expect(pages.map(linesOn)).toStrictEqual([[1, 2], [2]]);
  });

  it("leaves the break where it fell when neither end of the paragraph is left alone", () => {
    const boxes = stack([[10], [10, 10, 10, 10, 10]], 100, true);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 135 });

    expect(pages.map(linesOn)).toStrictEqual([[1, 2], [3]]);
  });

  it("breaks a paragraph that starts at the top of its page, having nowhere to move it", () => {
    const pages = breakStack({
      cells: [],
      boxes: stack([[10, 10]], 100, true),
      topPt: 100,
      bottomPt: 115,
    });

    expect(pages.map(linesOn)).toStrictEqual([[1], [1]]);
  });

  it("splits a paragraph whose file asks for no widow control", () => {
    const boxes = stack([[10], [10, 10]], 100, false);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 125 });

    expect(pages.map(linesOn)).toStrictEqual([[1, 1], [1]]);
  });

  // An empty paragraph as tall as the room its mark stands in, and one that keeps
  // room below itself as well.
  const emptyAt = (index: number, topPt: number, markPt: number, afterPt = 0): ParagraphBox => ({
    index,
    topPt,
    anchorTopPt: topPt,
    resumesUnderPt: 0,
    heightPt: markPt + afterPt,
    lines: [],
    marker: null,
    markTopPt: topPt,
    contentBottomPt: topPt + markPt,
    widowControl: false,
    keepNext: false,
    startsPage: false,
    endsPage: false,
    endsPageAtASection: false,
    contentWidthPt: 0,
    clipTo: null,
    paint: null,
  });

  it("counts an empty paragraph's own height against the page", () => {
    const boxes = [emptyAt(0, 100, 20), emptyAt(1, 120, 20)];
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 130 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1]]);
    expect(pages[1]?.boxes[0]?.topPt).toBe(100);
  });

  it("leaves an empty paragraph where the room below it is all that overflows", () => {
    const boxes = [emptyAt(0, 100, 20), emptyAt(1, 120, 8, 40)];
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 130 });

    expect(pages.map(indexesOn)).toStrictEqual([[0, 1]]);
  });

  it("gives a paragraph that asked for a page of its own one, with room to spare", () => {
    const boxes = asking(stack([[10], [10], [10]]), 1, { startsPage: true });
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 200 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1, 2]]);
    expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(100);
  });

  it("makes no page for a paragraph that asked for one and already stands at a top", () => {
    const boxes = asking(stack([[10], [10]]), 0, { startsPage: true });
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 200 });

    expect(pages.map(indexesOn)).toStrictEqual([[0, 1]]);
  });

  it("breaks a paragraph at a line of its own that asked to start a page", () => {
    const boxes = asking(stack([[10, 10, 10]]), 0, { line: 1 });
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 200 });

    expect(pages.map(linesOn)).toStrictEqual([[1], [2]]);
    expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(100);
  });

  // A break the paragraph ends on draws no line of its own, so there is nothing in
  // the stack to be seen at: the paragraph after it is what carries the page over.
  it("puts what follows a paragraph that ended on a break on a page of its own", () => {
    const boxes = asking(stack([[10], [10], [10]]), 0, { endsPage: true });
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 200 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1, 2]]);
  });

  // Widow control moves a break it would otherwise strand a line with; a break the
  // document asked for is not one of those.
  it("holds a line asked onto a page there whatever widow control would say", () => {
    const boxes = asking(stack([[10, 10, 10, 10]], 100, true), 0, { line: 3 });
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 200 });

    expect(pages.map(linesOn)).toStrictEqual([[3], [1]]);
  });
});

// Word's own answers, measured by the authored `space-above-a-break` document:
// four kinds of break open a page there, and the paragraph opening the page asks
// for 18pt above itself in each. Only the section break drew its first line that
// far below the top of the page.
describe("breakStack over the room a paragraph asks for above itself", () => {
  // The room a paragraph keeps above its first line, which leaves its top standing
  // that far above the line and moves everything under it down.
  const roomAbove = (
    boxes: readonly ParagraphBox[],
    at: number,
    roomPt: number,
  ): readonly ParagraphBox[] =>
    boxes.map((box, index) => {
      if (index < at) return box;
      const lower = (topPt: number): number => topPt + roomPt;
      return {
        ...box,
        topPt: index === at ? box.topPt : lower(box.topPt),
        heightPt: index === at ? box.heightPt + roomPt : box.heightPt,
        markTopPt: lower(box.markTopPt),
        contentBottomPt: lower(box.contentBottomPt),
        lines: box.lines.map((line) => ({
          ...line,
          topPt: lower(line.topPt),
          baselinePt: lower(line.baselinePt),
        })),
      };
    });

  it("keeps that room where a section break opened the page", () => {
    const boxes = roomAbove(
      asking(stack([[10], [10]]), 0, { endsPage: true, endsPageAtASection: true }),
      1,
      18,
    );
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 200 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1]]);
    expect(pages[1]?.boxes[0]?.topPt).toBe(100);
    expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(118);
  });

  it("leaves it behind where a break in the text opened the page", () => {
    const boxes = roomAbove(asking(stack([[10], [10]]), 0, { endsPage: true }), 1, 18);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 200 });

    expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(100);
  });

  it("leaves it behind where the paragraph asked for a page of its own", () => {
    const boxes = roomAbove(asking(stack([[10], [10]]), 1, { startsPage: true }), 1, 18);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 200 });

    expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(100);
  });

  it("leaves it behind where the foot of the page carried the paragraph over", () => {
    const boxes = roomAbove(stack([[10], [10]]), 1, 18);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 125 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1]]);
    expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(100);
  });
});

// Every answer here is Word's own, measured by the authored `keeping` document:
// nine cases, each a page's worth of paragraphs told exactly how tall to be so that
// the room left at the foot is arithmetic, and each on a page of its own so that
// what one case did to the flow cannot reach the next.
describe("breakStack over paragraphs held to the one after them", () => {
  const holding = (
    boxes: readonly ParagraphBox[],
    ...at: readonly number[]
  ): readonly ParagraphBox[] =>
    boxes.map((box, index) => (at.includes(index) ? { ...box, keepNext: true } : box));

  it("leaves a paragraph where there is room for the one it holds", () => {
    const boxes = holding(stack([[10], [10], [10]]), 1);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 130 });

    expect(pages.map(indexesOn)).toStrictEqual([[0, 1, 2]]);
  });

  it("moves one onto the page the paragraph it holds begins", () => {
    const boxes = holding(stack([[10], [10], [10]]), 1);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 125 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1, 2]]);
    expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(100);
  });

  it("pulls a whole chain of them back rather than the last one alone", () => {
    const boxes = holding(stack([[10], [10], [10], [10]]), 1, 2);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 130 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1, 2, 3]]);
  });

  // Word gives up rather than chasing a paragraph it can never catch, and gives up
  // after moving once rather than where the pair started.
  it("moves one once where what it holds is taller than any page", () => {
    const boxes = holding(stack([[10], [10], [50]]), 1);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 130 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1], [2]]);
  });

  // The case that says the giving up is a paragraph moving at most once and not a
  // paragraph running out of room to move into: the second of the chain is left
  // mid page with a whole page below it.
  it("leaves the second of such a chain where the first move put it", () => {
    const boxes = holding(stack([[10], [10], [10], [50]]), 1, 2);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 140 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1, 2], [3]]);
    expect(pages[1]?.boxes[1]?.lines[0]?.topPt).toBe(110);
  });

  it("does not follow the paragraph it holds onto a page that one asked for", () => {
    const boxes = asking(holding(stack([[10], [10], [10]]), 1), 2, { startsPage: true });
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 200 });

    expect(pages.map(indexesOn)).toStrictEqual([[0, 1], [2]]);
  });

  it("does not follow it over a break of its own either", () => {
    const boxes = asking(holding(stack([[10], [10], [10]]), 1), 1, { endsPage: true });
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 200 });

    expect(pages.map(indexesOn)).toStrictEqual([[0, 1], [2]]);
  });

  // It is the paragraph's end that is held: one the break already runs through has
  // its last line on the page the paragraph it holds begins, so there is nothing to
  // move and widow control's break stands.
  it("leaves a paragraph the break already runs through where it broke", () => {
    const boxes = holding(stack([[10], [10, 10, 10, 10], [10]], 100, true), 1);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 130 });

    expect(pages.map(indexesOn)).toStrictEqual([
      [0, 1],
      [1, 2],
    ]);
    expect(pages.map(linesOn)).toStrictEqual([
      [1, 2],
      [2, 1],
    ]);
  });

  // **The paragraph it holds need not overflow at its own start to have been
  // carried off the page.** Read off `bd42bfc93fdf`, whose ninth page Word opens
  // with two headings this project left at the foot of the eighth: the paragraph
  // they hold has its first line inside the page by a quarter of a point and its
  // second past the foot, so widow control takes the whole of it forward and
  // nothing ever overflows where it starts.
  it("follows a paragraph widow control carried forward whole", () => {
    const boxes = holding(stack([[10], [10], [10, 10]], 100, true), 1);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 135 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1, 2]]);
    expect(pages.map(linesOn)).toStrictEqual([[1], [1, 2]]);
  });

  it("moves nothing for the last paragraph of the story, which holds nothing", () => {
    const boxes = holding(stack([[10], [10]]), 1);
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 115 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1]]);
  });
});

// A cell is not broken by anything of its own: it is cut where the text it holds
// broke, and the piece of it on each page is what it covered there.
describe("breakStack over the cells of a table", () => {
  const boxes = stack([[20], [20], [20]], 100);

  it("keeps a cell on the page its row landed on, shifted with it", () => {
    const pages = breakStack({
      boxes,
      cells: [cellAt(140, 20)],
      topPt: 100,
      bottomPt: 140,
    });
    expect(pages[0]?.cells).toStrictEqual([]);
    expect(pages[1]?.cells).toStrictEqual([cellAt(100, 20)]);
  });

  it("cuts one the break ran through, and starts the rest again at the top", () => {
    const pages = breakStack({
      boxes,
      cells: [cellAt(120, 40)],
      topPt: 100,
      bottomPt: 140,
    });
    expect(pages[0]?.cells).toStrictEqual([cellAt(120, 20)]);
    expect(pages[1]?.cells).toStrictEqual([cellAt(100, 20)]);
  });
});

// A row that will not come apart is moved whole to the page under it, and one
// still too tall for a whole page is torn there. Measured on 2026-08-07 by the
// authored `tearing` document.
describe("breakStack over a row a break may not run through", () => {
  const untorn = (topPt: number, bottomPt: number, opensAt: number) => ({
    topPt,
    bottomPt,
    opensAt,
  });

  it("tears a row nothing has spoken for, as it does any other run of lines", () => {
    const pages = breakStack({
      boxes: stack([[20], [20], [20], [20]], 100),
      cells: [cellAt(140, 40)],
      topPt: 100,
      bottomPt: 160,
    });

    expect(linesOn(pages[0])).toStrictEqual([1, 1, 1]);
    expect(linesOn(pages[1])).toStrictEqual([1]);
  });

  it("moves one that refuses whole, and takes its cells with it", () => {
    const pages = breakStack({
      boxes: stack([[20], [20], [20], [20]], 100),
      cells: [cellAt(140, 40)],
      untornRows: [untorn(140, 180, 2)],
      topPt: 100,
      bottomPt: 160,
    });

    expect(linesOn(pages[0])).toStrictEqual([1, 1]);
    expect(linesOn(pages[1])).toStrictEqual([1, 1]);
    expect(pages[0]?.cells).toStrictEqual([]);
    expect(pages[1]?.cells).toStrictEqual([cellAt(100, 40)]);
  });

  it("leaves one alone that fits where it stands", () => {
    const pages = breakStack({
      boxes: stack([[20], [20], [20], [20]], 100),
      cells: [cellAt(140, 40)],
      untornRows: [untorn(140, 180, 2)],
      topPt: 100,
      bottomPt: 180,
    });

    expect(pages).toHaveLength(1);
    expect(linesOn(pages[0])).toStrictEqual([1, 1, 1, 1]);
  });

  // A document states its page size and margins per section, so what a page keeps
  // for the body is the section's rather than the document's. **A page belongs to
  // the section whose text opened it**, which is what these hold: the section a
  // paragraph stands in decides the page that paragraph opens and nothing about the
  // page it was already standing on.
  describe("where the sections make different pages", () => {
    const roomier = (from: number) => (box: ParagraphBox) =>
      box.index < from ? { topPt: 100, bottomPt: 160 } : { topPt: 40, bottomPt: 300 };

    it("opens a page at the top the section running on to it keeps", () => {
      const pages = breakStack({
        cells: [],
        boxes: stack([[20], [20], [20], [20], [20]], 100),
        topPt: 100,
        bottomPt: 160,
        bodyOf: roomier(3),
      });

      expect(pages.map(indexesOn)).toStrictEqual([
        [0, 1, 2],
        [3, 4],
      ]);
      expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(40);
    });

    it("holds a page to the foot the section that opened it keeps", () => {
      const pages = breakStack({
        cells: [],
        boxes: stack([[20], [20], [20], [20], [20]], 100),
        topPt: 100,
        bottomPt: 160,
        bodyOf: roomier(2),
      });

      expect(pages.map(indexesOn)).toStrictEqual([
        [0, 1, 2],
        [3, 4],
      ]);
    });

    it("says which paragraph opened each page", () => {
      const pages = breakStack({
        cells: [],
        boxes: stack([[20], [20], [20], [20], [20]], 100),
        topPt: 100,
        bottomPt: 160,
      });

      expect(pages.map((page) => page.openedBy)).toStrictEqual([0, 3]);
    });
  });

  // Moved once, its top stands at the top of a page and there is nowhere left to
  // move it to, so the break falls through it as it would through anything else.
  it("tears one no page has room for, once it has been moved", () => {
    const pages = breakStack({
      boxes: stack([[20], [20], [20], [20], [20]], 100),
      cells: [cellAt(120, 80)],
      untornRows: [untorn(120, 200, 1)],
      topPt: 100,
      bottomPt: 160,
    });

    expect(linesOn(pages[0])).toStrictEqual([1]);
    expect(linesOn(pages[1])).toStrictEqual([1, 1, 1]);
    expect(linesOn(pages[2])).toStrictEqual([1]);
  });
});

// **An object text wraps round may be drawn up until its own top reaches the foot
// of the line anchoring it, and moves the paragraph on when even that leaves it
// past the bottom of the page.** Measured on 2026-08-08 by the authored
// `objects-and-the-footer` document, which is where the foot it may rise to was
// found: a box hung 100pt below its 24pt line with 172pt of room under it moved,
// though drawing it up to the foot of the text would have left its top 22pt below
// the paragraph's own and 2pt above the foot of its line.
describe("an object anchored to a paragraph", () => {
  // The band an object stands in the way with, which these cases ask nothing about:
  // they are about the room under an object, and the break pass reads a band nowhere
  // yet.
  const anchored = (topPt: number, bottomPt: number) => ({
    topPt,
    bottomPt,
    anchoredAt: 1,
    band: { leftPt: 0, rightPt: 0, topPt, bottomPt },
  });

  // Three paragraphs of one 24pt line each from 100, so the second is anchored at
  // 124 and the foot of its line stands at 148.
  const three = () => stack([[24], [24], [24]], 100);

  it("leaves the page where the object fits standing at the foot of its own line", () => {
    const pages = breakStack({
      cells: [],
      boxes: three(),
      anchoredObjects: [anchored(200, 340)],
      topPt: 100,
      bottomPt: 300,
    });

    expect(pages.map(indexesOn)).toStrictEqual([[0, 1, 2]]);
  });

  it("moves the paragraph on where the object would have to rise above its line", () => {
    const pages = breakStack({
      cells: [],
      boxes: three(),
      anchoredObjects: [anchored(200, 360)],
      topPt: 100,
      bottomPt: 300,
    });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1, 2]]);
  });

  // A paragraph answers for every object at once: of the three the corpus
  // template anchors to one paragraph, only the first fails to fit and Word takes
  // the whole paragraph on to the next page.
  it("moves the paragraph on for one object of several", () => {
    const pages = breakStack({
      cells: [],
      boxes: three(),
      anchoredObjects: [anchored(200, 340), anchored(200, 360)],
      topPt: 100,
      bottomPt: 300,
    });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1, 2]]);
  });
});

// **A page break's own line never has to fit.** Seven cases were put to Word on
// 2026-08-15 by `break-foot-probe`, three repeats each, over a body running 36 to 756
// filled to 660, and it left the break's line at the foot of the page it started on in
// every one. Each case below is that page: a filler, then a paragraph ending on a
// break whose line stands at 110, then what follows the break. The foot of the body is
// what moves from case to case, so the line has room for itself twice over, exactly
// enough, a twip too little, half its own height too little, and a single twip against
// a line of 24. The break ends the page whatever it did with the line.
describe("a paragraph whose page break ends the page", () => {
  const breaking = (heightPt = 24): readonly ParagraphBox[] =>
    asking(stack([[10], [heightPt], [10]], 100), 1, { endsPage: true });

  const pagesOf = (bottomPt: number, heightPt = 24) =>
    breakStack({ cells: [], boxes: breaking(heightPt), topPt: 100, bottomPt });

  it("keeps its line where it stands with room for it twice over", () => {
    const pages = pagesOf(158);

    expect(indexesOn(pages[0] ?? { boxes: [] })).toStrictEqual([0, 1]);
    expect(indexesOn(pages[1] ?? { boxes: [] })).toStrictEqual([2]);
  });

  it("keeps its line ending exactly at the foot", () => {
    expect(indexesOn(pagesOf(134)[0] ?? { boxes: [] })).toStrictEqual([0, 1]);
  });

  it("keeps its line missing the foot by a twip", () => {
    expect(indexesOn(pagesOf(133.95)[0] ?? { boxes: [] })).toStrictEqual([0, 1]);
  });

  it("keeps its line missing the foot by half of itself", () => {
    expect(indexesOn(pagesOf(122)[0] ?? { boxes: [] })).toStrictEqual([0, 1]);
  });

  it("keeps its line left a single twip of room", () => {
    expect(indexesOn(pagesOf(110.05)[0] ?? { boxes: [] })).toStrictEqual([0, 1]);
  });

  // The shape two corpus documents have, which is what sent anyone looking: an
  // ordinary line of 12.207 left 8 points of room.
  it("keeps an ordinary line left less room than it takes", () => {
    expect(indexesOn(pagesOf(118, 12.207)[0] ?? { boxes: [] })).toStrictEqual([0, 1]);
  });

  // And what follows the break opens the next page however much room was left.
  it("opens the next page for what follows the break", () => {
    const pages = pagesOf(400);

    expect(pages).toHaveLength(2);
    expect(indexesOn(pages[1] ?? { boxes: [] })).toStrictEqual([2]);
  });

  // A paragraph closing a section is left as it was, since none of the seven asked
  // about one and the page a section opens keeps its own room above itself.
  it("moves a line that will not fit where the break closes a section", () => {
    const boxes = asking(breaking(), 1, { endsPageAtASection: true });
    const pages = breakStack({ cells: [], boxes, topPt: 100, bottomPt: 122 });

    expect(indexesOn(pages[0] ?? { boxes: [] })).toStrictEqual([0]);
    expect(indexesOn(pages[1] ?? { boxes: [] })).toStrictEqual([1]);
  });
});
