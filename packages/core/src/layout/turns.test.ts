import { describe, expect, it } from "vitest";

import { boundsOfTurn, roomForTurn, turnedAbout, turnsOnItsSide } from "./turns.js";

// Every number here was measured off Word's own pdf of the authored
// `rotated-drawings` and `rotated-drawing-ties` documents, whose 72 x 24pt picture
// is turned nine ways and each way written out three times.
describe("what a turn costs the flow", () => {
  it("keeps the extent the way it was stored for a turn nearer no turn at all", () => {
    for (const degrees of [0, 30, 26.7, 180, 135, 359]) {
      expect(turnsOnItsSide(degrees)).toBe(false);
      expect(roomForTurn({ widthPt: 72, heightPt: 24 }, degrees)).toStrictEqual({
        widthPt: 72,
        heightPt: 24,
      });
    }
  });

  it("lays the extent on its side for a turn nearer a quarter", () => {
    for (const degrees of [90, 270, 45, 225, -90]) {
      expect(turnsOnItsSide(degrees)).toBe(true);
      expect(roomForTurn({ widthPt: 72, heightPt: 24 }, degrees)).toStrictEqual({
        widthPt: 24,
        heightPt: 72,
      });
    }
  });

  // The boundary is an eighth of a turn and it rounds away from the way the box
  // was stored: 45 is read as a quarter and 135 as a half. Word gave the three
  // repeats of the first 72pt of line each and the three of the second 24pt.
  it("reads a turn standing exactly between two quarters as the further one", () => {
    expect(turnsOnItsSide(45)).toBe(true);
    expect(turnsOnItsSide(135)).toBe(false);
    expect(turnsOnItsSide(225)).toBe(true);
    expect(turnsOnItsSide(315)).toBe(false);
  });
});

describe("where a turned object is painted", () => {
  // A quarter turn of a 72 x 24 picture in the 24 x 72 room its line kept: the
  // paint fills that room exactly. Word drew it at the left of the text frame,
  // 24 wide and 72 tall.
  it("fills the room it was kept in where the turn is a quarter", () => {
    const painted = boundsOfTurn({ leftPt: 12, topPt: 100, widthPt: 72, heightPt: 24 }, 90);

    expect(painted.leftPt).toBeCloseTo(36, 9);
    expect(painted.topPt).toBeCloseTo(76, 9);
    expect(painted.widthPt).toBeCloseTo(24, 9);
    expect(painted.heightPt).toBeCloseTo(72, 9);
  });

  // 30 degrees, which Word drew 74.35 x 56.78 with its middle where the middle of
  // the 72 x 24 box was: at a left of 34.82 against the box's 36.
  it("hangs out of the room on every side where the turn is no quarter at all", () => {
    const painted = boundsOfTurn({ leftPt: 36, topPt: 100, widthPt: 72, heightPt: 24 }, 30);

    expect(painted.widthPt).toBeCloseTo(74.354, 2);
    expect(painted.heightPt).toBeCloseTo(56.785, 2);
    expect(painted.leftPt).toBeCloseTo(34.823, 2);
    expect(painted.topPt).toBeCloseTo(83.608, 2);
  });

  it("paints a turn of no degrees where the object already stood", () => {
    const rect = { leftPt: 36, topPt: 100, widthPt: 72, heightPt: 24 };

    expect(boundsOfTurn(rect, 0)).toStrictEqual(rect);
  });
});

describe("a group turned as a whole", () => {
  it("swings each shape inside it about the group's own middle", () => {
    const swung = turnedAbout({ xPt: 20, yPt: 10 }, { xPt: 10, yPt: 10 }, 90);

    expect(swung.xPt).toBeCloseTo(10, 9);
    expect(swung.yPt).toBeCloseTo(20, 9);
  });

  it("leaves a shape where it stands where the group was never turned", () => {
    expect(turnedAbout({ xPt: 20, yPt: 10 }, { xPt: 10, yPt: 10 }, 0)).toStrictEqual({
      xPt: 20,
      yPt: 10,
    });
  });
});
