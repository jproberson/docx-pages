import { describe, expect, it } from "vitest";

import { fitLine, freeSpans, type WrapBand } from "./wrapping.js";

const band = (leftPt: number, rightPt: number, topPt: number, bottomPt: number): WrapBand => ({
  leftPt,
  rightPt,
  topPt,
  bottomPt,
});

const fit = (options: {
  widthPt: number;
  bands?: readonly WrapBand[];
  topPt?: number;
  heightPt?: number;
}) =>
  fitLine({
    topPt: options.topPt ?? 100,
    heightPt: options.heightPt ?? 12,
    leftPt: 36,
    rightPt: 576,
    widthPt: options.widthPt,
    bands: options.bands ?? [],
  });

describe("freeSpans", () => {
  it("leaves the whole frame free when nothing crosses it", () => {
    expect(freeSpans(36, 576, [])).toStrictEqual([{ leftPt: 36, rightPt: 576 }]);
  });

  it("opens a run on each side of a band standing in the middle", () => {
    expect(freeSpans(36, 576, [band(200, 300, 0, 10)])).toStrictEqual([
      { leftPt: 36, rightPt: 200 },
      { leftPt: 300, rightPt: 576 },
    ]);
  });

  it("takes overlapping bands as one", () => {
    expect(freeSpans(36, 576, [band(200, 300, 0, 10), band(250, 400, 0, 10)])).toStrictEqual([
      { leftPt: 36, rightPt: 200 },
      { leftPt: 400, rightPt: 576 },
    ]);
  });

  it("reads bands in frame order however they arrive", () => {
    expect(freeSpans(36, 576, [band(400, 500, 0, 10), band(100, 200, 0, 10)])).toStrictEqual([
      { leftPt: 36, rightPt: 100 },
      { leftPt: 200, rightPt: 400 },
      { leftPt: 500, rightPt: 576 },
    ]);
  });

  it("leaves no run at all when a band spans the frame", () => {
    expect(freeSpans(36, 576, [band(0, 600, 0, 10)])).toStrictEqual([]);
  });

  it("ignores a band standing outside the frame", () => {
    expect(freeSpans(36, 576, [band(0, 20, 0, 10)])).toStrictEqual([{ leftPt: 36, rightPt: 576 }]);
  });
});

describe("fitLine", () => {
  it("leaves a line where it started when nothing is in its way", () => {
    expect(fit({ widthPt: 400 })).toStrictEqual({ topPt: 100, leftPt: 36, rightPt: 576 });
  });

  it("moves a line's start past an object hanging over its left edge", () => {
    const slot = fit({ widthPt: 200, bands: [band(16, 76, 90, 130)] });
    expect(slot).toStrictEqual({ topPt: 100, leftPt: 76, rightPt: 576 });
  });

  // The line starts at 100 and is 12 tall.
  it("falls a line to the foot of the object that blocked it", () => {
    const slot = fit({ widthPt: 500, bands: [band(300, 580, 90, 130)] });
    expect(slot).toStrictEqual({ topPt: 130, leftPt: 36, rightPt: 576 });
  });

  // Measured against Word by sweeping a box's bottom edge down through a line: the
  // line moved as soon as the edge reached it, half a point in.
  it("moves a line an object reaches only the top of", () => {
    const slot = fit({ widthPt: 500, bands: [band(300, 580, 90, 104)] });
    expect(slot).toStrictEqual({ topPt: 104, leftPt: 36, rightPt: 576 });
  });

  // Word answers an outline for the middle of a line alone: two documents leave a
  // line beside an outline that reaches 4pt into the top of it.
  it("leaves a line an outline reaches only the top of", () => {
    const slot = fit({ widthPt: 500, bands: [{ ...band(300, 580, 90, 104), outlined: true }] });
    expect(slot).toStrictEqual({ topPt: 100, leftPt: 36, rightPt: 576 });
  });

  // And a line an outline does block steps down by its own height rather than
  // falling to the foot of it, which lands it below the outline's own bottom.
  it("steps a line down by its own height past an outline", () => {
    const slot = fit({ widthPt: 500, bands: [{ ...band(300, 580, 90, 130), outlined: true }] });
    expect(slot).toStrictEqual({ topPt: 124, leftPt: 36, rightPt: 576 });
  });

  it("keeps falling past each object in turn", () => {
    const slot = fit({
      widthPt: 500,
      bands: [band(300, 580, 90, 130), band(280, 580, 125, 160)],
    });
    expect(slot).toStrictEqual({ topPt: 160, leftPt: 36, rightPt: 576 });
  });

  it("stops falling once no object reaches the line at all", () => {
    const slot = fit({
      widthPt: 500,
      bands: [band(300, 580, 90, 130), band(280, 580, 145, 200)],
    });
    expect(slot).toStrictEqual({ topPt: 130, leftPt: 36, rightPt: 576 });
  });

  it("sits beside an object rather than under it when the room is there", () => {
    const slot = fit({ widthPt: 150, bands: [band(300, 580, 90, 130)] });
    expect(slot).toStrictEqual({ topPt: 100, leftPt: 36, rightPt: 300 });
  });

  it("ignores an object that ends exactly where the line begins", () => {
    const slot = fit({ widthPt: 500, bands: [band(300, 580, 60, 100)] });
    expect(slot).toStrictEqual({ topPt: 100, leftPt: 36, rightPt: 576 });
  });

  // Word refuses a gap this narrow even to a line that needs no room at all.
  it("will not put a line in a gap too narrow to be worth having", () => {
    const slot = fit({
      widthPt: 0,
      bands: [band(0, 200, 90, 130), band(210, 600, 90, 130)],
    });
    expect(slot.topPt).toBe(130);
  });

  it("takes a gap wide enough to be worth having", () => {
    const slot = fit({
      widthPt: 0,
      bands: [band(0, 200, 90, 130), band(230, 600, 90, 130)],
    });
    expect(slot).toStrictEqual({ topPt: 100, leftPt: 200, rightPt: 230 });
  });

  // A wedge inside the same band: nothing of it at the top, all 50pt of it across
  // the foot, with the band 4pt wider on the left and 6pt on the right for the
  // distances text is kept off it by.
  const wedge: WrapBand = {
    ...band(16, 76, 90, 130),
    outline: [
      { xPt: 20, yPt: 90 },
      { xPt: 70, yPt: 130 },
      { xPt: 20, yPt: 130 },
    ],
  };

  it("keeps a line off the part of an outline beside it, not the whole of it", () => {
    const slot = fit({ widthPt: 200, topPt: 100, bands: [wedge] });
    expect(slot).toStrictEqual({ topPt: 100, leftPt: 53.5, rightPt: 576 });
  });

  it("keeps a line off the whole of an outline where the whole of it is beside it", () => {
    const slot = fit({ widthPt: 200, topPt: 118, bands: [wedge] });
    expect(slot).toStrictEqual({ topPt: 118, leftPt: 76, rightPt: 576 });
  });

  // A line with no height of its own has no step to take, so an outline it cannot
  // sit beside leaves it where it started rather than looking for room for ever.
  it("leaves a line that cannot step where it started", () => {
    const slot = fit({
      widthPt: 500,
      heightPt: 0,
      bands: [{ ...band(0, 600, 90, 130), outlined: true }],
    });
    expect(slot).toStrictEqual({ topPt: 100, leftPt: 36, rightPt: 576 });
  });

  it("leaves a line nothing stands in the way of where the frame cannot hold it", () => {
    const slot = fit({ widthPt: 800 });
    expect(slot).toStrictEqual({ topPt: 100, leftPt: 36, rightPt: 576 });
  });
});
