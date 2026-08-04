import type { ParagraphFrame } from "../docx/styles.js";
import { twipsToPoints } from "./units.js";

// Word's own default, from w:defaultTabStop, which both reference documents set to
// exactly this.
export const DEFAULT_TAB_STOP_PT = 36;

// Positions are compared, not accumulated, so this only has to absorb the last
// bits of a conversion out of twips.
const EPSILON = 1e-9;

// Every stop the paragraph has, in points from the left edge of the text area. A
// hanging indent adds an implicit stop at the left indent, which is the one the
// number on the first line tabs across to.
export function tabStopsPt(frame: ParagraphFrame): readonly number[] {
  const declared = frame.tabStopsTwips.map(twipsToPoints);
  if (frame.indentFirstLineTwips >= 0) return declared;
  return [...declared, twipsToPoints(frame.indentLeftTwips)].sort((left, right) => left - right);
}

// Past the last stop a paragraph declares, Word falls back to its default ones.
export function nextTabStopPt(fromPt: number, stops: readonly number[]): number {
  const declared = stops.find((stop) => stop > fromPt + EPSILON);
  if (declared !== undefined) return declared;
  return (Math.floor(fromPt / DEFAULT_TAB_STOP_PT + EPSILON) + 1) * DEFAULT_TAB_STOP_PT;
}
