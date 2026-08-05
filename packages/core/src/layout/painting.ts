import {
  borderExtentPt,
  SIDES,
  type Border,
  type BorderStyle,
  type Borders,
} from "../docx/borders.js";
import type { ParagraphPaint, PlacedCell } from "./stack.js";

// Where the colour behind a thing stops and where each line round it runs. Word
// paints both as filled rectangles and its own pdf reports them that way, so this
// is what the rendering can be held against: see `how-word-draws-a-border`.

export type PaintedFill = {
  readonly color: string;
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

// One band of a border. A double border is two of them, a band's width either
// side of the edge; a dashed or dotted one is a single band the pattern is cut
// out of, which is the drawing's business rather than the geometry's.
export type PaintedLine = {
  readonly color: string | null;
  readonly style: BorderStyle;
  readonly widthPt: number;
  readonly vertical: boolean;
  // Where the band's own centre lies across the direction it runs in, and how far
  // it runs.
  readonly atPt: number;
  readonly fromPt: number;
  readonly toPt: number;
};

export type Painted = {
  readonly fills: readonly PaintedFill[];
  readonly lines: readonly PaintedLine[];
};

const EMPTY: Painted = { fills: [], lines: [] };

type Rect = {
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

const halfPt = (border: Border | null): number => borderExtentPt(border) / 2;

// The bands one border makes, centred on the edge it runs along and reaching into
// the corners at both ends so that the lines round a cell meet.
function linesOf(borders: Borders, rect: Rect, corners: Rect): readonly PaintedLine[] {
  return SIDES.flatMap((side) => {
    const border = borders[side];
    if (border === null) return [];

    const vertical = side === "left" || side === "right";
    const edgePt =
      side === "left"
        ? rect.leftPt
        : side === "right"
          ? rect.leftPt + rect.widthPt
          : side === "top"
            ? rect.topPt
            : rect.topPt + rect.heightPt;
    const offsets = border.style === "double" ? [-border.widthPt, border.widthPt] : [0];

    return offsets.map((offset) => ({
      color: border.color,
      style: border.style,
      widthPt: border.widthPt,
      vertical,
      atPt: edgePt + offset,
      fromPt: vertical ? corners.topPt : corners.leftPt,
      toPt: vertical ? corners.topPt + corners.heightPt : corners.leftPt + corners.widthPt,
    }));
  });
}

// A cell's own colour and lines. The fill stops at the inner edge of each border,
// since a border is centred on the cell's own edge and half of it falls inside.
export function paintOfCell(cell: PlacedCell): Painted {
  const { borders } = cell;
  const left = halfPt(borders.left);
  const right = halfPt(borders.right);
  const top = halfPt(borders.top);
  const bottom = halfPt(borders.bottom);

  const fills =
    cell.fillColor === null
      ? []
      : [
          {
            color: cell.fillColor,
            leftPt: cell.leftPt + left,
            topPt: cell.topPt + top,
            widthPt: Math.max(0, cell.widthPt - left - right),
            heightPt: Math.max(0, cell.heightPt - top - bottom),
          },
        ];

  return {
    fills,
    lines: linesOf(borders, cell, {
      leftPt: cell.leftPt - left,
      topPt: cell.topPt - top,
      widthPt: cell.widthPt + left + right,
      heightPt: cell.heightPt + top + bottom,
    }),
  };
}

// How far a paragraph's own fill and lines stand outside its text area to either
// side. Measured against Word at 1.44pt, which is a fiftieth of an inch, and
// nothing at all above or below.
export const PARAGRAPH_PAINT_PT = 1.44;

// A paragraph's own colour and lines, given the room its lines took. A border
// stands off by whatever `w:space` it asks for on top of that, and is centred on
// its own edge like any other.
export function paintOfParagraph(paint: ParagraphPaint, topPt: number, bottomPt: number): Painted {
  const leftPt = paint.leftPt - PARAGRAPH_PAINT_PT;
  const rightPt = paint.rightPt + PARAGRAPH_PAINT_PT;
  if (rightPt <= leftPt || bottomPt < topPt) return EMPTY;

  const fills =
    paint.fillColor === null
      ? []
      : [
          {
            color: paint.fillColor,
            leftPt,
            topPt,
            widthPt: rightPt - leftPt,
            heightPt: bottomPt - topPt,
          },
        ];

  // Each side stands off by its own room, so each is drawn round a rectangle of
  // its own rather than round one shared with the others.
  const lines = SIDES.flatMap((side) => {
    const border = paint.borders[side];
    if (border === null) return [];
    const roomPt = border.spacePt + borderExtentPt(border) / 2;
    const rect = {
      leftPt: leftPt - roomPt,
      topPt: topPt - roomPt,
      widthPt: rightPt - leftPt + roomPt * 2,
      heightPt: bottomPt - topPt + roomPt * 2,
    };
    return linesOf(
      { top: null, left: null, bottom: null, right: null, [side]: border },
      rect,
      rect,
    );
  });

  return { fills, lines };
}
