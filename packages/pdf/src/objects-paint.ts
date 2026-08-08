import type { Drawable, PlacedPaint } from "@docx-pages/core";

import { bottomOf, upFromTop, type PdfPage } from "./coordinates.js";
import type { Content } from "./content.js";

// A shape's own paint: the colour behind it and the line round it, which mirrors
// the viewer's `painted`.
//
// Word centres an outline on the edge it runs along, which is where a stroked path
// sits, so the geometry is drawn at the object's own bounds and the stroke falls
// half outside them of its own accord. The viewer has to grow its layer to leave
// room for that half; a pdf has no layer to grow, and the half simply falls where
// it falls.

export type ObjectDrawable = Extract<Drawable, { kind: "object" }>;

export function paintedObject(
  out: Content,
  page: PdfPage,
  at: ObjectDrawable,
  paint: PlacedPaint,
): void {
  const { outline } = paint;
  if (paint.fillColor === null && outline === null) return;

  out.save();
  if (outline !== null) {
    out.strokeColor(outline.color);
    out.lineWidth(outline.widthPt);
    out.dash(null);
  }
  if (paint.fillColor !== null) out.fillColor(paint.fillColor);

  // A line shape is stored with no height, and is the one geometry whose paint is
  // the diagonal of its own box rather than the box.
  if (paint.geometry === "line") {
    if (outline !== null) {
      out.line(
        at.leftPt,
        upFromTop(page, at.topPt),
        at.leftPt + at.widthPt,
        upFromTop(page, at.topPt + at.heightPt),
      );
      out.stroke();
    }
    out.restore();
    return;
  }

  out.rectangle(at.leftPt, bottomOf(page, at.topPt, at.heightPt), at.widthPt, at.heightPt);
  // `B` paints a path and then strokes it, which is the one operator for a
  // rectangle that has both; a path with only one of them takes the operator for
  // that one alone, so that nothing is drawn in a colour that was never set.
  if (paint.fillColor !== null && outline !== null) out.fillAndStroke();
  else if (paint.fillColor !== null) out.fill();
  else out.stroke();

  out.restore();
}
