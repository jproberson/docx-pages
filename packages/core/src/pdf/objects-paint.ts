import { ROUNDED_CORNER_FRACTION, type DrawingFlip } from "../docx/drawing.js";
import type { Drawable } from "../layout/drawables.js";
import type { PlacedPaint } from "../layout/floats.js";

import { bottomOf, type PdfPage } from "./coordinates.js";
import type { Content } from "./content.js";

// A shape's own paint: the colour behind it and the line round it, which mirrors
// the viewer's `painted`.
//
// Word centres an outline on the edge it runs along, which is where a stroked path
// sits, so the geometry is drawn at the object's own bounds and the stroke falls
// half outside them of its own accord. The viewer has to grow its layer to leave
// room for that half; a pdf has no layer to grow, and the half simply falls where
// it falls.
//
// **The presets are drawn here in a pdf's own operators and in the viewer's in
// svg, and that is the one thing both backends state twice.** Neither notation can
// be written in the other, so what is shared instead is every number the shapes are
// built out of: the corner a rounded rectangle is rounded by comes from core, and
// nothing here rounds one by a figure of its own.

export type ObjectDrawable = Extract<Drawable, { kind: "object" }>;

// The box a shape is drawn in, in a pdf's coordinates, which count up the page.
type Box = {
  readonly leftPt: number;
  readonly rightPt: number;
  readonly bottomPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

const boxOf = (page: PdfPage, at: ObjectDrawable): Box => {
  const bottomPt = bottomOf(page, at.topPt, at.heightPt);
  return {
    leftPt: at.leftPt,
    rightPt: at.leftPt + at.widthPt,
    bottomPt,
    topPt: bottomPt + at.heightPt,
    widthPt: at.widthPt,
    heightPt: at.heightPt,
  };
};

// How far along a quarter of a circle its two control points stand, which is the
// closest a cubic curve comes to an arc. A pdf draws no arc of its own, so an
// ellipse is four of these and every rounded corner is one.
const ARC = 0.5522847498307936;

// A connector is stored as a box with a line across it, so the two corners it joins
// are the ones the flips choose. This is the same rule the viewer draws a `line`
// by, and a shape with a symmetry either way is drawn the same however it flipped.
function lineAcross(out: Content, box: Box, flip: DrawingFlip): void {
  const [fromXPt, toXPt] = flip.horizontal ? [box.rightPt, box.leftPt] : [box.leftPt, box.rightPt];
  // The viewer measures a line's ends down from the top of its box, so the
  // unflipped one runs from the top left corner to the bottom right.
  const [fromYPt, toYPt] = flip.vertical ? [box.bottomPt, box.topPt] : [box.topPt, box.bottomPt];

  out.line(fromXPt, fromYPt, toXPt, toYPt);
}

function ellipse(out: Content, box: Box): void {
  const middleXPt = box.leftPt + box.widthPt / 2;
  const middleYPt = box.bottomPt + box.heightPt / 2;
  const acrossPt = (box.widthPt / 2) * ARC;
  const upPt = (box.heightPt / 2) * ARC;

  out.moveTo(middleXPt, box.topPt);
  out.curveTo(
    middleXPt + acrossPt,
    box.topPt,
    box.rightPt,
    middleYPt + upPt,
    box.rightPt,
    middleYPt,
  );
  out.curveTo(
    box.rightPt,
    middleYPt - upPt,
    middleXPt + acrossPt,
    box.bottomPt,
    middleXPt,
    box.bottomPt,
  );
  out.curveTo(
    middleXPt - acrossPt,
    box.bottomPt,
    box.leftPt,
    middleYPt - upPt,
    box.leftPt,
    middleYPt,
  );
  out.curveTo(box.leftPt, middleYPt + upPt, middleXPt - acrossPt, box.topPt, middleXPt, box.topPt);
  out.closePath();
}

function roundedRectangle(out: Content, box: Box): void {
  const radiusPt = Math.min(box.widthPt, box.heightPt) * ROUNDED_CORNER_FRACTION;
  const pullPt = radiusPt * (1 - ARC);

  out.moveTo(box.leftPt + radiusPt, box.bottomPt);
  out.lineTo(box.rightPt - radiusPt, box.bottomPt);
  out.curveTo(
    box.rightPt - pullPt,
    box.bottomPt,
    box.rightPt,
    box.bottomPt + pullPt,
    box.rightPt,
    box.bottomPt + radiusPt,
  );
  out.lineTo(box.rightPt, box.topPt - radiusPt);
  out.curveTo(
    box.rightPt,
    box.topPt - pullPt,
    box.rightPt - pullPt,
    box.topPt,
    box.rightPt - radiusPt,
    box.topPt,
  );
  out.lineTo(box.leftPt + radiusPt, box.topPt);
  out.curveTo(
    box.leftPt + pullPt,
    box.topPt,
    box.leftPt,
    box.topPt - pullPt,
    box.leftPt,
    box.topPt - radiusPt,
  );
  out.lineTo(box.leftPt, box.bottomPt + radiusPt);
  out.curveTo(
    box.leftPt,
    box.bottomPt + pullPt,
    box.leftPt + pullPt,
    box.bottomPt,
    box.leftPt + radiusPt,
    box.bottomPt,
  );
  out.closePath();
}

// Flipped upright, the apex is at the foot of the box and the base along its head,
// which is the viewer's rule read into a pdf's coordinates.
function triangle(out: Content, box: Box, flip: DrawingFlip): void {
  const apexYPt = flip.vertical ? box.bottomPt : box.topPt;
  const baseYPt = flip.vertical ? box.topPt : box.bottomPt;

  out.moveTo(box.leftPt + box.widthPt / 2, apexYPt);
  out.lineTo(box.rightPt, baseYPt);
  out.lineTo(box.leftPt, baseYPt);
  out.closePath();
}

export function paintedObject(
  out: Content,
  page: PdfPage,
  at: ObjectDrawable,
  paint: PlacedPaint,
): void {
  const { outline } = paint;
  if (paint.fillColor === null && outline === null) return;
  // A path nothing here plays is drawn as nothing at all rather than as the box it
  // fits in. One corpus document rules a whole page with a custom path, and the box
  // it fits in is a filled rectangle over everything the page holds.
  if (paint.geometry === "custom") return;

  out.save();
  if (outline !== null) {
    out.strokeColor(outline.color);
    out.lineWidth(outline.widthPt);
    out.dash(null);
  }
  if (paint.fillColor !== null) out.fillColor(paint.fillColor);

  const box = boxOf(page, at);

  // A line shape is stored with no height, and is the one geometry whose paint is
  // the diagonal of its own box rather than the box. It has no inside, so a colour
  // behind it would fill nothing and it is stroked or it is not drawn.
  if (paint.geometry === "line") {
    if (outline !== null) {
      lineAcross(out, box, at.flip);
      out.stroke();
    }
    out.restore();
    return;
  }

  switch (paint.geometry) {
    case "ellipse":
      ellipse(out, box);
      break;
    case "rounded-rectangle":
      roundedRectangle(out, box);
      break;
    case "triangle":
      triangle(out, box, at.flip);
      break;
    case "rectangle":
      out.rectangle(box.leftPt, box.bottomPt, box.widthPt, box.heightPt);
      break;
  }

  // `B` paints a path and then strokes it, which is the one operator for a shape
  // that has both; a path with only one of them takes the operator for that one
  // alone, so that nothing is drawn in a colour that was never set.
  if (paint.fillColor !== null && outline !== null) out.fillAndStroke();
  else if (paint.fillColor !== null) out.fill();
  else out.stroke();

  out.restore();
}
