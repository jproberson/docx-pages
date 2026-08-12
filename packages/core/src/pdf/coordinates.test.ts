import { describe, expect, it } from "vitest";

import { turnedAboutInPdf } from "./coordinates.js";

// Which way a turn goes round, which is the one thing about it that can be wrong
// without looking wrong anywhere but on the page.
//
// A turn is stated clockwise, the way a reader sees one and the way Word writes
// one. A pdf's own angles run the other way, so every case here is written in what
// a reader would see and not in what the matrix holds: a page that turns the wrong
// way is a matrix that is right about everything except its sign.

// Where a point lands once the matrix has had it. A pdf maps `(x, y)` through
// `[a b c d e f]` as `(a x + c y + e, b x + d y + f)`.
const through = (
  matrix: readonly number[],
  xPt: number,
  yPt: number,
): { readonly xPt: number; readonly yPt: number } => {
  const [a = 0, b = 0, c = 0, d = 0, e = 0, f = 0] = matrix;
  return { xPt: a * xPt + c * yPt + e, yPt: b * xPt + d * yPt + f };
};

// Everything below turns about here, and reads in a pdf's coordinates: up the page
// is a larger y.
const CENTRE = { xPt: 100, yPt: 500 };

const turned = (degrees: number, xPt: number, yPt: number) =>
  through(turnedAboutInPdf(degrees, CENTRE.xPt, CENTRE.yPt), xPt, yPt);

describe("turnedAboutInPdf", () => {
  it("leaves everything where it stands where nothing was turned", () => {
    const put = turned(0, 130, 560);

    expect(put.xPt).toBeCloseTo(130, 9);
    expect(put.yPt).toBeCloseTo(560, 9);
  });

  it("leaves the point it turns about where it stands, whatever the angle", () => {
    for (const degrees of [17, 90, 180, 270, -45]) {
      const put = turned(degrees, CENTRE.xPt, CENTRE.yPt);

      expect(put.xPt).toBeCloseTo(CENTRE.xPt, 9);
      expect(put.yPt).toBeCloseTo(CENTRE.yPt, 9);
    }
  });

  // The case the whole file exists for. Turned a quarter clockwise, what stood
  // above the middle stands to the right of it, and a matrix that took layout's
  // angle unchanged would put it to the left.
  it("takes what stood above the middle round to the right of it, a quarter turn on", () => {
    const put = turned(90, CENTRE.xPt, CENTRE.yPt + 40);

    expect(put.xPt).toBeCloseTo(CENTRE.xPt + 40, 9);
    expect(put.yPt).toBeCloseTo(CENTRE.yPt, 9);
  });

  it("takes what stood to the right of the middle round beneath it, the same turn on", () => {
    const put = turned(90, CENTRE.xPt + 40, CENTRE.yPt);

    expect(put.xPt).toBeCloseTo(CENTRE.xPt, 9);
    expect(put.yPt).toBeCloseTo(CENTRE.yPt - 40, 9);
  });

  it("turns a half turn onto the far side of the middle", () => {
    const put = turned(180, CENTRE.xPt + 30, CENTRE.yPt + 40);

    expect(put.xPt).toBeCloseTo(CENTRE.xPt - 30, 9);
    expect(put.yPt).toBeCloseTo(CENTRE.yPt - 40, 9);
  });

  // Word states a turn anticlockwise as a negative one rather than as its
  // reflection, so both have to arrive at the same place.
  it("turns anticlockwise where the angle is negative", () => {
    const back = turned(-90, CENTRE.xPt, CENTRE.yPt + 40);
    const round = turned(270, CENTRE.xPt, CENTRE.yPt + 40);

    expect(back.xPt).toBeCloseTo(CENTRE.xPt - 40, 9);
    expect(back.yPt).toBeCloseTo(CENTRE.yPt, 9);
    expect(round.xPt).toBeCloseTo(back.xPt, 9);
    expect(round.yPt).toBeCloseTo(back.yPt, 9);
  });

  // A turn moves a drawing and never resizes one, so the distance from the middle
  // is the same before and after however odd the angle.
  it("keeps everything as far from the middle as it was", () => {
    const away = (put: { readonly xPt: number; readonly yPt: number }): number =>
      Math.hypot(put.xPt - CENTRE.xPt, put.yPt - CENTRE.yPt);

    for (const degrees of [7, 33, 90, 128, 355]) {
      expect(away(turned(degrees, CENTRE.xPt + 30, CENTRE.yPt + 40))).toBeCloseTo(50, 9);
    }
  });
});
