import { describe, expect, it } from "vitest";

import { NO_BORDERS, type Border, type Borders } from "../docx/borders.js";
import { paintOfCell, paintOfParagraph, PARAGRAPH_PAINT_PT } from "./painting.js";
import type { PlacedCell } from "./stack.js";

const line = (widthPt: number, over: Partial<Border> = {}): Border => ({
  style: "single",
  widthPt,
  color: "#FF0000",
  spacePt: 0,
  ...over,
});

const cellWith = (borders: Partial<Borders>, fillColor: string | null = null): PlacedCell => ({
  leftPt: 100,
  topPt: 200,
  widthPt: 72,
  heightPt: 20,
  fillColor,
  borders: { ...NO_BORDERS, ...borders },
  holds: [],
});

describe("paintOfCell", () => {
  it("centres each line on the edge it runs along", () => {
    const { lines } = paintOfCell(cellWith({ left: line(6), top: line(2) }));
    expect(lines.map((each) => [each.vertical, each.atPt, each.widthPt])).toStrictEqual([
      [false, 200, 2],
      [true, 100, 6],
    ]);
  });

  it("reaches into the corners so that two lines meet there", () => {
    const { lines } = paintOfCell(cellWith({ left: line(6), top: line(2) }));
    const [top] = lines;
    expect([top?.fromPt, top?.toPt]).toStrictEqual([97, 172]);
  });

  it("draws a double line as the two bands it is", () => {
    const { lines } = paintOfCell(cellWith({ top: line(2, { style: "double" }) }));
    expect(lines.map((each) => each.atPt)).toStrictEqual([198, 202]);
  });

  it("stops the fill at the inner side of each line", () => {
    const { fills } = paintOfCell(cellWith({ left: line(6), right: line(2) }, "#DEEBF7"));
    expect(fills).toStrictEqual([
      { color: "#DEEBF7", leftPt: 103, topPt: 200, widthPt: 68, heightPt: 20 },
    ]);
  });

  it("lays a cell that asks for no colour none", () => {
    expect(paintOfCell(cellWith({})).fills).toStrictEqual([]);
  });
});

const paintWith = (borders: Partial<Borders>, fillColor: string | null = null) =>
  paintOfParagraph(
    { leftPt: 100, rightPt: 400, fillColor, borders: { ...NO_BORDERS, ...borders } },
    200,
    220,
  );

describe("paintOfParagraph", () => {
  it("reaches past the text area at each side and not above or below it", () => {
    const { fills } = paintWith({}, "#FBE5D6");
    expect(fills).toStrictEqual([
      {
        color: "#FBE5D6",
        leftPt: 100 - PARAGRAPH_PAINT_PT,
        topPt: 200,
        widthPt: 300 + PARAGRAPH_PAINT_PT * 2,
        heightPt: 20,
      },
    ]);
  });

  it("stands a line off by the room it asks for, past that same reach", () => {
    const { lines } = paintWith({ bottom: line(1.5, { spacePt: 12 }) });
    expect(lines.map((each) => each.atPt)).toStrictEqual([220 + 12 + 0.75]);
  });

  it("stands a line at the side off by the reach as well", () => {
    const { lines } = paintWith({ left: line(1.5) });
    expect(lines.map((each) => each.atPt)).toStrictEqual([100 - PARAGRAPH_PAINT_PT - 0.75]);
  });

  it("draws nothing for a paragraph that asks for neither", () => {
    expect(paintWith({})).toStrictEqual({ fills: [], lines: [] });
  });
});
