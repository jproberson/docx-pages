import type { SectionColumns } from "../docx/section.js";
import { twipsToPoints } from "./units.js";

// Where each of a section's columns stands across the frame its margins leave.
//
// Measured on 2026-08-07 by the authored `columns` document, which is committed,
// and read off Word's own pdf: the first and last line of every case there is
// right-aligned, since a right-aligned line ends where its column ends and nothing
// else in a drawing says how wide a column is.
//
// **An equal-width section divides what its frame leaves.** The gaps come off the
// frame first and what is left is split: two columns of a 540pt frame with 36pt
// between them are 252pt each, drawn at 36 and at 324, and three with 18pt between
// them are 168pt, drawn at 36, 222 and 408. The last column's right edge is the
// frame's own, to the point.
//
// **A section stating its own widths is taken at its word**, each column's gap
// standing after it: widths of 3600 and 6480 twips with 720 between them were drawn
// 180pt wide at 36 and 324pt wide at 252.
export type Column = {
  readonly leftPt: number;
  readonly widthPt: number;
};

export function columnsAcross(
  columns: SectionColumns,
  frame: { readonly leftPt: number; readonly widthPt: number },
): readonly Column[] {
  if (columns.count <= 1) return [{ leftPt: frame.leftPt, widthPt: frame.widthPt }];

  if (columns.widthsTwips.length >= columns.count) {
    const across: Column[] = [];
    let leftPt = frame.leftPt;
    for (let at = 0; at < columns.count; at += 1) {
      const widthPt = twipsToPoints(columns.widthsTwips[at] ?? 0);
      across.push({ leftPt, widthPt });
      leftPt += widthPt + twipsToPoints(columns.gapsTwips[at] ?? 0);
    }
    return across;
  }

  const spacePt = twipsToPoints(columns.spaceTwips);
  const widthPt = (frame.widthPt - spacePt * (columns.count - 1)) / columns.count;
  return Array.from({ length: columns.count }, (_, at) => ({
    leftPt: frame.leftPt + at * (widthPt + spacePt),
    widthPt,
  }));
}
