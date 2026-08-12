import type { SectionGeometry } from "../docx/section.js";
import { twipsToPoints } from "../layout/units.js";

// A pdf's user space is a point across and a point down by default, which is what
// layout measures in, so nothing here is ever scaled. The whole of the difference
// between the two is which way up they are: layout counts down from the top of the
// page and a pdf counts up from the bottom.
//
// Keeping that in one place is the point of this file. A flip applied twice, or
// applied to a height rather than to a position, is the kind of mistake that draws
// a page that looks almost right, and the numbers it produces agree with nothing.

// The page a section makes, in the points everything else here is measured in.
export type PdfPage = {
  readonly widthPt: number;
  readonly heightPt: number;
};

export const pdfPageOf = (geometry: SectionGeometry): PdfPage => ({
  widthPt: twipsToPoints(geometry.widthTwips),
  heightPt: twipsToPoints(geometry.heightTwips),
});

/**
 * Where a distance measured down from the top of the page stands in a pdf, which
 * measures up from the bottom.
 *
 * A width or a height is not a position and must not go through here: the flip
 * turns the top edge of a rectangle into its bottom one, so a rectangle is flipped
 * by its far edge and keeps the size it had.
 */
export const upFromTop = (page: PdfPage, downPt: number): number => page.heightPt - downPt;

// The bottom edge of a rectangle laid out from its top, which is what a pdf wants
// of one: `re` takes a corner, a width and a height, and the corner it takes is
// the low one.
export const bottomOf = (page: PdfPage, topPt: number, heightPt: number): number =>
  page.heightPt - topPt - heightPt;

/**
 * The matrix that turns whatever is drawn under it about the middle of the box it
 * stands in, as far clockwise as layout says the drawing was turned.
 *
 * **The angle goes in negated, and that is this file's flip once more rather than
 * a second rule.** Layout counts a turn clockwise as a reader sees it; a pdf's own
 * angles run the other way round, because its y counts up the page where layout's
 * counts down. A turn stated once therefore comes out on the page the way it was
 * stated, and a turn taken straight from layout would come out mirrored.
 *
 * The centre is given in a pdf's coordinates, not layout's, since everything a
 * matrix touches is already in them.
 */
export const turnedAboutInPdf = (
  turnDegrees: number,
  centreXPt: number,
  centreYPt: number,
): readonly number[] => {
  const radians = (-turnDegrees * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);

  // Turning about a point is a move to the origin, the turn, and the move back,
  // multiplied out: a pdf takes the six numbers of the product rather than three
  // operators.
  return [
    cos,
    sin,
    -sin,
    cos,
    centreXPt - cos * centreXPt + sin * centreYPt,
    centreYPt - sin * centreXPt - cos * centreYPt,
  ];
};
