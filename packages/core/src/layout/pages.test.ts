import { describe, expect, it } from "vitest";

import { breakStack } from "./pages.js";
import type { ParagraphBox, ParagraphMarker, PlacedLine } from "./stack.js";

const MARK = {
  font: { kind: "named", name: "Meridian Sans" },
  fontSizePt: 10,
  bold: false,
  italic: false,
  raisePt: 0,
  color: null,
} as const;

const line = (topPt: number, heightPt: number, text: string): PlacedLine => ({
  line: {
    segments: [{ kind: "text", mark: MARK, text, widthPt: 20, offsetPt: 0 }],
    widthPt: 20,
    heightPt,
    ascentPt: heightPt * 0.8,
  },
  leftPt: 0,
  topPt,
  heightPt,
  baselinePt: topPt + heightPt * 0.8,
});

const MARKER: ParagraphMarker = {
  text: "1.",
  mark: MARK,
  widthPt: 6,
  leftPt: 0,
  baselinePt: 0,
};

// Paragraphs are laid down one after another, each as tall as the lines it holds.
function stack(shape: readonly (readonly number[])[], topPt = 100): readonly ParagraphBox[] {
  const boxes: ParagraphBox[] = [];
  let top = topPt;

  shape.forEach((heights, index) => {
    const lines: PlacedLine[] = [];
    let lineTop = top;
    for (const heightPt of heights) {
      lines.push(line(lineTop, heightPt, `p${String(index)}l${String(lines.length)}`));
      lineTop += heightPt;
    }
    boxes.push({ index, topPt: top, heightPt: lineTop - top, lines, marker: null });
    top = lineTop;
  });

  return boxes;
}

const indexesOn = (page: { readonly boxes: readonly ParagraphBox[] }): readonly number[] =>
  page.boxes.map((box) => box.index);

describe("breakStack", () => {
  it("leaves a stack that fits on the one page", () => {
    const pages = breakStack({ boxes: stack([[10], [10], [10]]), topPt: 100, bottomPt: 200 });

    expect(pages).toHaveLength(1);
    expect(indexesOn(pages[0] ?? { boxes: [] })).toStrictEqual([0, 1, 2]);
  });

  it("moves the paragraph that would cross the bottom onto the next page", () => {
    const pages = breakStack({ boxes: stack([[10], [10], [10]]), topPt: 100, bottomPt: 125 });

    expect(pages).toHaveLength(2);
    expect(indexesOn(pages[0] ?? { boxes: [] })).toStrictEqual([0, 1]);
    expect(indexesOn(pages[1] ?? { boxes: [] })).toStrictEqual([2]);
  });

  it("starts the next page at the top of the body rather than where the flow was", () => {
    const pages = breakStack({ boxes: stack([[10], [10], [10]]), topPt: 100, bottomPt: 125 });

    expect(pages[1]?.boxes[0]?.topPt).toBe(100);
    expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(100);
    expect(pages[1]?.boxes[0]?.lines[0]?.baselinePt).toBe(108);
  });

  it("breaks a paragraph between its own lines and carries the rest over", () => {
    const pages = breakStack({ boxes: stack([[10, 10, 10]]), topPt: 100, bottomPt: 120 });

    expect(pages).toHaveLength(2);
    expect(pages[0]?.boxes[0]?.lines).toHaveLength(2);
    expect(pages[1]?.boxes[0]?.index).toBe(0);
    expect(pages[1]?.boxes[0]?.lines).toHaveLength(1);
    expect(pages[1]?.boxes[0]?.lines[0]?.topPt).toBe(100);
  });

  it("keeps carrying paragraphs onto further pages for as long as the stack runs", () => {
    const pages = breakStack({
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
    const pages = breakStack({ boxes, topPt: 100, bottomPt: 105 });

    expect(pages[0]?.boxes[0]?.marker?.baselinePt).toBe(108);
    expect(pages[1]?.boxes[0]?.marker).toBeNull();
  });

  it("draws a line taller than the page rather than looking for a page it fits", () => {
    const pages = breakStack({ boxes: stack([[400]]), topPt: 100, bottomPt: 200 });

    expect(pages).toHaveLength(1);
    expect(pages[0]?.boxes[0]?.lines).toHaveLength(1);
  });

  it("counts an empty paragraph's own height against the page", () => {
    const boxes: readonly ParagraphBox[] = [
      { index: 0, topPt: 100, heightPt: 20, lines: [], marker: null },
      { index: 1, topPt: 120, heightPt: 20, lines: [], marker: null },
    ];
    const pages = breakStack({ boxes, topPt: 100, bottomPt: 130 });

    expect(pages.map(indexesOn)).toStrictEqual([[0], [1]]);
    expect(pages[1]?.boxes[0]?.topPt).toBe(100);
  });
});
