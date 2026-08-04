import type { InlineDrawing } from "../docx/inlines.js";
import type { SectionGeometry } from "../docx/section.js";
import type { ParagraphFrame } from "../docx/styles.js";
import type { Theme } from "../docx/theme.js";
import { resolveContent, type PartResolver, type PlacedContent } from "./floats.js";
import { emuToPoints, twipsToPoints } from "./units.js";

export type PlacedInline = {
  readonly drawing: InlineDrawing;
  readonly content: PlacedContent;
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

export type PlaceInlinesInput = {
  readonly drawings: readonly InlineDrawing[];
  readonly page: SectionGeometry;
  readonly frame: ParagraphFrame;
  readonly paragraphTopPt: number;
  readonly resolvePart: PartResolver;
  readonly theme: Theme;
};

// Inline drawings sit in the paragraph's own line, so they run along it in order
// from wherever the paragraph's alignment starts the line.
export function placeInlines(input: PlaceInlinesInput): readonly PlacedInline[] {
  const { page, frame } = input;
  const startPt = twipsToPoints(page.margin.leftTwips + frame.indentLeftTwips);
  const endPt = twipsToPoints(page.widthTwips - page.margin.rightTwips - frame.indentRightTwips);

  const widths = input.drawings.map((drawing) => emuToPoints(drawing.widthEmu));
  const total = widths.reduce((sum, width) => sum + width, 0);

  let left = lineStart(frame, startPt, endPt, total);
  const placed: PlacedInline[] = [];

  input.drawings.forEach((drawing, at) => {
    const widthPt = widths[at] ?? 0;
    placed.push({
      drawing,
      content: resolveContent(drawing.content, input.resolvePart, input.theme),
      leftPt: left,
      topPt: input.paragraphTopPt,
      widthPt,
      heightPt: emuToPoints(drawing.heightEmu),
    });
    left += widthPt;
  });

  return placed;
}

function lineStart(
  frame: ParagraphFrame,
  startPt: number,
  endPt: number,
  contentPt: number,
): number {
  if (frame.alignment === "right") return endPt - contentPt;
  if (frame.alignment === "center") return startPt + (endPt - startPt - contentPt) / 2;
  return startPt;
}
