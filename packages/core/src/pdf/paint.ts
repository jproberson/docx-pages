import type { DrawnLine, DrawnPaint, HighlightPaint } from "../layout/drawables.js";
import type { PaintedFill } from "../layout/painting.js";

import { bottomOf, upFromTop, type PdfPage } from "./coordinates.js";
import type { Content } from "./content.js";

// Everything drawn behind a story's text, which mirrors the viewer's `paintLayer`
// and answers the same question in the other notation.
//
// **Nothing here decides anything.** `drawables.ts` turns a cell and a paragraph
// into the rectangles and the bands Word paints them as, and states how the dashes
// of each band fall; both backends draw exactly what it hands over. A border is a
// band centred on the edge it runs along, not a stroke round a box, which is why
// nothing here draws a rectangle's outline.

function paintedFill(out: Content, page: PdfPage, fill: PaintedFill | HighlightPaint): void {
  out.fillColor(fill.color);
  out.rectangle(
    fill.leftPt,
    bottomOf(page, fill.topPt, fill.heightPt),
    fill.widthPt,
    fill.heightPt,
  );
  out.fill();
}

// One band of a border, drawn as the line it is: `atPt` is where its own centre
// lies across the direction it runs in, which is where a stroke sits.
function paintedLine(out: Content, page: PdfPage, line: DrawnLine): void {
  out.strokeColor(line.color);
  out.lineWidth(line.widthPt);
  out.dash(line.dashes);

  if (line.vertical) {
    out.line(line.atPt, upFromTop(page, line.fromPt), line.atPt, upFromTop(page, line.toPt));
  } else {
    const atPt = upFromTop(page, line.atPt);
    out.line(line.fromPt, atPt, line.toPt, atPt);
  }
  out.stroke();
}

/**
 * Everything drawn behind a story's text: what each thing paints, in the order
 * `drawables.ts` put them in, and last the highlights, which Word draws over a
 * shaded paragraph.
 */
export function paintLayer(
  out: Content,
  page: PdfPage,
  painted: readonly DrawnPaint[],
  highlights: readonly HighlightPaint[],
): void {
  // The dash pattern and the line width outlive the band that set them, so the
  // whole layer is drawn inside a saved state rather than each band undoing what
  // it did.
  out.save();
  for (const each of painted) {
    for (const fill of each.fills) paintedFill(out, page, fill);
    for (const line of each.lines) paintedLine(out, page, line);
  }
  for (const each of highlights) paintedFill(out, page, each);
  out.restore();
}
