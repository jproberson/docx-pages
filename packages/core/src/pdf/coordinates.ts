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
