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

  it("drops a line below an object it is too wide to sit beside", () => {
    const slot = fit({ widthPt: 500, bands: [band(300, 580, 90, 130)] });
    expect(slot).toStrictEqual({ topPt: 130, leftPt: 36, rightPt: 576 });
  });

  it("keeps falling until it clears every object in turn", () => {
    const slot = fit({
      widthPt: 500,
      bands: [band(300, 580, 90, 130), band(280, 580, 125, 160)],
    });
    expect(slot).toStrictEqual({ topPt: 160, leftPt: 36, rightPt: 576 });
  });

  it("stops falling once an object no longer reaches the line", () => {
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

  it("leaves a line that fits nowhere in the frame it started in", () => {
    const slot = fit({ widthPt: 500, bands: [band(0, 600, 90, 130)] });
    expect(slot).toStrictEqual({ topPt: 130, leftPt: 36, rightPt: 576 });
  });
});
