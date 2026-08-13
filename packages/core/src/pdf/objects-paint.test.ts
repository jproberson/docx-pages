import { describe, expect, it } from "vitest";

import { ROUNDED_CORNER_FRACTION } from "../docx/drawing.js";
import type { PlacedPaint } from "../layout/floats.js";

import type { Content } from "./content.js";
import type { PdfPage } from "./coordinates.js";
import { paintedObject, type ObjectDrawable } from "./objects-paint.js";

// What each preset is actually laid down as, which is the half of a shape that the
// pdf reader cannot answer for: every one of these comes back from a reader as the
// bounds it fits in, and an ellipse drawn as its own box has exactly those bounds.
//
// So the operators are read instead. The numbers below are in a pdf's coordinates,
// which count up the page, and the page is 200pt tall so that a box laid out 50pt
// down stands 100pt up.

const PAGE: PdfPage = { widthPt: 300, heightPt: 200 };

// A 60 by 40 box, 50pt down the page: 20 to 80 across, 110 to 150 up.
const BOX = { leftPt: 20, topPt: 50, widthPt: 60, heightPt: 40 };
const LEFT = 20;
const RIGHT = 80;
const BOTTOM = 110;
const TOP = 150;

type Call = { readonly op: string; readonly numbers: readonly number[] };

const recorder = (): { readonly out: Content; readonly calls: Call[] } => {
  const calls: Call[] = [];
  const put =
    (op: string) =>
    (...numbers: number[]): void => {
      calls.push({ op, numbers });
    };

  const out: Content = {
    save: put("save"),
    restore: put("restore"),
    fillColor: () => calls.push({ op: "fillColor", numbers: [] }),
    strokeColor: () => calls.push({ op: "strokeColor", numbers: [] }),
    lineWidth: put("lineWidth"),
    dash: () => calls.push({ op: "dash", numbers: [] }),
    rectangle: put("rectangle"),
    line: put("line"),
    moveTo: put("moveTo"),
    lineTo: put("lineTo"),
    curveTo: put("curveTo"),
    closePath: put("closePath"),
    fill: put("fill"),
    stroke: put("stroke"),
    fillAndStroke: put("fillAndStroke"),
    clip: put("clip"),
    transform: () => calls.push({ op: "transform", numbers: [] }),
    drawObject: () => calls.push({ op: "drawObject", numbers: [] }),
    beginText: put("beginText"),
    endText: put("endText"),
    font: () => calls.push({ op: "font", numbers: [] }),
    characterSpacing: put("characterSpacing"),
    textPosition: put("textPosition"),
    textMatrix: () => calls.push({ op: "textMatrix", numbers: [] }),
    showGlyphs: () => calls.push({ op: "showGlyphs", numbers: [] }),
    bytes: () => new Uint8Array(),
  };

  return { out, calls };
};

const UNFLIPPED = { horizontal: false, vertical: false };

const object = (flip = UNFLIPPED): ObjectDrawable => ({
  kind: "object",
  key: "one",
  name: "Shape 1",
  content: { kind: "unknown" },
  ...BOX,
  flip,
  turnDegrees: 0,
});

const FILLED: PlacedPaint = {
  fillColor: "ff0000",
  outline: null,
  geometry: "rectangle",
  path: null,
};

const drawn = (paint: PlacedPaint, flip = UNFLIPPED): readonly Call[] => {
  const { out, calls } = recorder();
  paintedObject(out, PAGE, object(flip), paint);
  return calls;
};

const of = (calls: readonly Call[], op: string): readonly Call[] =>
  calls.filter((call) => call.op === op);

const points = (calls: readonly Call[], op: string): readonly (readonly number[])[] =>
  of(calls, op).map((call) => call.numbers);

describe("what a shape is drawn as", () => {
  it("lays a rectangle down as its own box", () => {
    const calls = drawn(FILLED);

    expect(points(calls, "rectangle")).toStrictEqual([[LEFT, BOTTOM, 60, 40]]);
    expect(of(calls, "fill")).toHaveLength(1);
  });

  // The one that mattered enough to be written down where the geometry is read: a
  // corpus document rules a whole page with a path, and the box that path fits in
  // is a filled rectangle over everything else the page holds.
  it("draws a path it cannot play as nothing at all, not as the box it fits in", () => {
    const calls = drawn({ ...FILLED, geometry: "custom", path: null });

    expect(calls).toStrictEqual([]);
  });

  // A path in shares of its own box, drawn where those shares land in the box the
  // object was given. A pdf counts y up the page, so a share of nought is the top.
  it("draws a path the file drew point by point", () => {
    const calls = drawn({
      ...FILLED,
      geometry: "custom",
      path: [
        { kind: "move", to: { x: 0.5, y: 0 } },
        { kind: "line", to: { x: 1, y: 1 } },
        { kind: "curve", first: { x: 0.5, y: 1 }, second: { x: 0, y: 0.5 }, to: { x: 0, y: 0 } },
        { kind: "close" },
      ],
    });

    expect(points(calls, "moveTo")).toStrictEqual([[50, TOP]]);
    expect(points(calls, "lineTo")).toStrictEqual([[RIGHT, BOTTOM]]);
    expect(points(calls, "curveTo")).toStrictEqual([[50, BOTTOM, LEFT, 130, LEFT, TOP]]);
    expect(of(calls, "closePath")).toHaveLength(1);
    expect(of(calls, "fill")).toHaveLength(1);
    expect(of(calls, "rectangle")).toHaveLength(0);
  });

  it("mirrors a path the shape was flipped", () => {
    const path = [{ kind: "move", to: { x: 0.25, y: 0 } }] as const;
    const flipped = drawn(
      { ...FILLED, geometry: "custom", path: [...path] },
      { horizontal: true, vertical: true },
    );

    expect(points(flipped, "moveTo")).toStrictEqual([[LEFT + 45, BOTTOM]]);
  });

  it("draws an ellipse as four curves through the middle of each edge", () => {
    const calls = drawn({ ...FILLED, geometry: "ellipse", path: null });
    const curves = points(calls, "curveTo");

    expect(of(calls, "rectangle")).toHaveLength(0);
    expect(curves).toHaveLength(4);
    // Started at the top of the ellipse, and each curve ends at the next quarter
    // round: the right, the foot, the left, and back to the top.
    expect(points(calls, "moveTo")).toStrictEqual([[50, TOP]]);
    expect(curves.map((curve) => curve.slice(4))).toStrictEqual([
      [RIGHT, 130],
      [50, BOTTOM],
      [LEFT, 130],
      [50, TOP],
    ]);
  });

  it("rounds a rectangle's corner by a sixth of its shorter side", () => {
    const calls = drawn({ ...FILLED, geometry: "rounded-rectangle", path: null });
    const radius = 40 * ROUNDED_CORNER_FRACTION;

    expect(radius).toBeCloseTo(6.667, 3);
    expect(of(calls, "curveTo")).toHaveLength(4);
    // Four straight edges, each held off the corner by the radius, and a curve
    // round every corner between them.
    expect(points(calls, "moveTo")).toStrictEqual([[LEFT + radius, BOTTOM]]);
    expect(points(calls, "lineTo")).toStrictEqual([
      [RIGHT - radius, BOTTOM],
      [RIGHT, TOP - radius],
      [LEFT + radius, TOP],
      [LEFT, BOTTOM + radius],
    ]);
  });

  it("stands a triangle on its base, apex at the head of the box", () => {
    const calls = drawn({ ...FILLED, geometry: "triangle", path: null });

    expect(points(calls, "moveTo")).toStrictEqual([[50, TOP]]);
    expect(points(calls, "lineTo")).toStrictEqual([
      [RIGHT, BOTTOM],
      [LEFT, BOTTOM],
    ]);
  });

  it("turns a triangle over where the shape was flipped", () => {
    const calls = drawn(
      { ...FILLED, geometry: "triangle", path: null },
      { horizontal: false, vertical: true },
    );

    expect(points(calls, "moveTo")).toStrictEqual([[50, BOTTOM]]);
    expect(points(calls, "lineTo")).toStrictEqual([
      [RIGHT, TOP],
      [LEFT, TOP],
    ]);
  });
});

// A connector is a box with a line across it, so which corners it joins is the
// whole of what a flip decides.
describe("a line shape", () => {
  const STROKED: PlacedPaint = {
    fillColor: null,
    outline: { color: "000000", widthPt: 1 },
    geometry: "line",
    path: null,
  };

  it("runs from the head of its box down to the foot", () => {
    expect(points(drawn(STROKED), "line")).toStrictEqual([[LEFT, TOP, RIGHT, BOTTOM]]);
  });

  it("runs the other way across where the shape was flipped across", () => {
    const flipped = drawn(STROKED, { horizontal: true, vertical: false });

    expect(points(flipped, "line")).toStrictEqual([[RIGHT, TOP, LEFT, BOTTOM]]);
  });

  it("runs from the foot up where the shape was flipped over", () => {
    const flipped = drawn(STROKED, { horizontal: false, vertical: true });

    expect(points(flipped, "line")).toStrictEqual([[LEFT, BOTTOM, RIGHT, TOP]]);
  });

  // It has no inside, so a colour behind it would fill nothing at all.
  it("is stroked and never filled", () => {
    const calls = drawn({ ...STROKED, fillColor: "ff0000" });

    expect(of(calls, "stroke")).toHaveLength(1);
    expect(of(calls, "fill")).toHaveLength(0);
    expect(of(calls, "fillAndStroke")).toHaveLength(0);
  });
});
