import type { ParagraphFrame, TabAlignment } from "../docx/styles.js";
import { twipsToPoints } from "./units.js";

// What the stops fall back to where nothing has read `w:defaultTabStop`, which is
// the 720 twips Word itself writes there.
export const DEFAULT_TAB_STOP_PT = 36;

// Positions are compared, not accumulated, so this only has to absorb the last
// bits of a conversion out of twips.
const EPSILON = 1e-9;

export type TabStopPt = {
  readonly positionPt: number;
  readonly alignment: TabAlignment;
};

// Every stop the paragraph has, in points from the left edge of the text area. A
// hanging indent adds an implicit stop at the left indent, which is the one the
// number on the first line tabs across to.
//
// A bar is not a place a tab ever lands, so it is not one of these: a tab passes it
// by and goes on to the stop or the default beyond it, which is what Word does.
// Word also draws a line down the page at a bar's position, on every line of the
// paragraph, which nothing here draws yet.
export function tabStopsPt(frame: ParagraphFrame): readonly TabStopPt[] {
  const declared = frame.tabStops
    .filter((stop) => stop.alignment !== "bar")
    .map((stop) => ({ positionPt: twipsToPoints(stop.positionTwips), alignment: stop.alignment }));
  if (frame.indentFirstLineTwips >= 0) return declared;

  const hanging = {
    positionPt: twipsToPoints(frame.indentLeftTwips),
    alignment: "left" as const,
  };
  return [...declared, hanging].sort((left, right) => left.positionPt - right.positionPt);
}

// Past the last stop a paragraph declares, Word falls back to the ones the document
// spaces evenly across the page, which start the text where they stand.
export function nextTabStop(
  fromPt: number,
  stops: readonly TabStopPt[],
  defaultStopPt: number = DEFAULT_TAB_STOP_PT,
): TabStopPt {
  const declared = stops.find((stop) => stop.positionPt > fromPt + EPSILON);
  if (declared !== undefined) return declared;
  if (defaultStopPt <= 0) return { positionPt: fromPt, alignment: "left" };
  return {
    positionPt: (Math.floor(fromPt / defaultStopPt + EPSILON) + 1) * defaultStopPt,
    alignment: "left",
  };
}
