// What a drawing turned after it was drawn costs the flow around it, and where it
// is painted once the flow has kept room for it.
//
// Measured on 2026-08-11 off Word's own pdf of the authored `rotated-drawings`
// document, whose 72 x 24pt inline picture is turned nine ways, each written out
// three times so the room its line kept is the distance from one repeat to the
// next.
//
// **Word does not keep the room the turn actually needs.** It rounds the turn to
// the nearest quarter and keeps the extent that way round, so a picture turned by
// a quarter is held in a box 24 wide and 72 tall and one turned by 30 degrees in
// the box it was stored in, though it is painted 74.35 x 56.78 and hangs out of it
// on every side. What Word kept, by the turn: 0 and 180 gave 24pt of line and 72
// of advance, 90 and 270 gave 72 and 24, 45 and 225 gave 72 and 24, and 135, 30
// and 26.7 gave 24 and 72. So the boundary is at the eighth and it rounds away
// from the turn the box was stored at: 45 is a quarter and 135 is a half.

export type TurnedSize = {
  readonly widthPt: number;
  readonly heightPt: number;
};

export type TurnedRect = TurnedSize & {
  readonly leftPt: number;
  readonly topPt: number;
};

const QUARTER_TURN = 90;
const HALF_TURN = 180;

const clockwise = (degrees: number): number => ((degrees % HALF_TURN) + HALF_TURN) % HALF_TURN;

// Whether the flow reads the turn as one that lays the object on its side.
export const turnsOnItsSide = (degrees: number): boolean => {
  const turn = clockwise(degrees);
  return turn >= QUARTER_TURN / 2 && turn < QUARTER_TURN + QUARTER_TURN / 2;
};

// The room the flow keeps for an object turned this far, which is the extent the
// way round the turn rounds to.
export const roomForTurn = (size: TurnedSize, degrees: number): TurnedSize =>
  turnsOnItsSide(degrees) ? { widthPt: size.heightPt, heightPt: size.widthPt } : size;

/**
 * Where an object stands before it is turned, given the room the flow kept for it.
 *
 * A turn is about the middle of that room, and the middle is the one point the two
 * boxes share: the object keeps the size it was stored at and is drawn from the
 * middle outwards, whatever the flow held open around it.
 */
export function unturnedRect(room: TurnedRect, size: TurnedSize): TurnedRect {
  return {
    leftPt: room.leftPt + (room.widthPt - size.widthPt) / 2,
    topPt: room.topPt + (room.heightPt - size.heightPt) / 2,
    widthPt: size.widthPt,
    heightPt: size.heightPt,
  };
}

// The rectangle a turned object's paint actually reaches, which is what a
// rendering of it can be compared against. Nothing in the layout is measured from
// this: the flow keeps `roomForTurn` and no more, and the paint hangs out of it.
export function boundsOfTurn(rect: TurnedRect, degrees: number): TurnedRect {
  const radians = (degrees * Math.PI) / HALF_TURN;
  const across = Math.abs(Math.cos(radians));
  const down = Math.abs(Math.sin(radians));
  const widthPt = rect.widthPt * across + rect.heightPt * down;
  const heightPt = rect.widthPt * down + rect.heightPt * across;
  return {
    leftPt: rect.leftPt + (rect.widthPt - widthPt) / 2,
    topPt: rect.topPt + (rect.heightPt - heightPt) / 2,
    widthPt,
    heightPt,
  };
}

// Where a point stands once the box holding it is turned about its own middle,
// which is what a group turned as a whole does to each shape inside it.
export function turnedAbout(
  point: { readonly xPt: number; readonly yPt: number },
  middle: { readonly xPt: number; readonly yPt: number },
  degrees: number,
): { readonly xPt: number; readonly yPt: number } {
  const radians = (degrees * Math.PI) / HALF_TURN;
  const across = Math.cos(radians);
  const down = Math.sin(radians);
  const x = point.xPt - middle.xPt;
  const y = point.yPt - middle.yPt;
  return {
    xPt: middle.xPt + x * across - y * down,
    yPt: middle.yPt + x * down + y * across,
  };
}
