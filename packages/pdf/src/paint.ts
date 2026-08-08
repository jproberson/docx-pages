import {
  paintOfCell,
  paintOfParagraph,
  type BorderStyle,
  type Painted,
  type PaintedFill,
  type PaintedLine,
  type PaintedParagraph,
  type PlacedCell,
} from "@docx-pages/core";

import { bottomOf, upFromTop, type PdfPage } from "./coordinates.js";
import type { Content } from "./content.js";

// Everything drawn behind a story's text, which mirrors the viewer's `paintLayer`
// and answers the same question in the other notation.
//
// The geometry is core's: `paintOfCell` and `paintOfParagraph` turn a cell and a
// paragraph into the rectangles and the bands Word paints them as, measured
// against Word's own pdf, and both backends draw exactly what those hand back. A
// border is a band centred on the edge it runs along, not a stroke round a box,
// which is why nothing here draws a rectangle's outline.

// How Word draws each pattern, measured at a width of a point and a half: a dashed
// line runs four widths on and four off, where a dotted one runs one and one. A
// double line is two bands, which the geometry has already made of it.
const DASHES: Readonly<Record<BorderStyle, readonly number[] | null>> = {
  single: null,
  double: null,
  dashed: [4, 4],
  dotted: [1, 1],
};

// What a band with no colour of its own is drawn in. The viewer says
// `currentColor` and inherits whatever the page is set in; a pdf inherits nothing,
// and the colour Word draws an unstated border in is black.
const DEFAULT_COLOR = "000000";

function paintedFill(out: Content, page: PdfPage, fill: PaintedFill): void {
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
function paintedLine(out: Content, page: PdfPage, line: PaintedLine): void {
  const dashes = DASHES[line.style];

  out.strokeColor(line.color ?? DEFAULT_COLOR);
  out.lineWidth(line.widthPt);
  out.dash(dashes === null ? null : dashes.map((each) => each * line.widthPt));

  if (line.vertical) {
    out.line(line.atPt, upFromTop(page, line.fromPt), line.atPt, upFromTop(page, line.toPt));
  } else {
    const atPt = upFromTop(page, line.atPt);
    out.line(line.fromPt, atPt, line.toPt, atPt);
  }
  out.stroke();
}

const drawn = (out: Content, page: PdfPage, painted: Painted): void => {
  for (const fill of painted.fills) paintedFill(out, page, fill);
  for (const line of painted.lines) paintedLine(out, page, line);
};

/**
 * Everything drawn behind a story's text: the cells of its tables first, then what
 * each paragraph asks for, which Word draws over the cell holding it.
 */
export function paintLayer(
  out: Content,
  page: PdfPage,
  cells: readonly PlacedCell[],
  paragraphs: readonly PaintedParagraph[],
): void {
  // The dash pattern and the line width outlive the band that set them, so the
  // whole layer is drawn inside a saved state rather than each band undoing what
  // it did.
  out.save();
  for (const cell of cells) drawn(out, page, paintOfCell(cell));
  for (const each of paragraphs) {
    drawn(out, page, paintOfParagraph(each.paint, each.topPt, each.bottomPt));
  }
  out.restore();
}
