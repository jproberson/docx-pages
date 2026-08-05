import type { InlineDrawing } from "../docx/inlines.js";
import { resolveContent, type PartResolver, type PlacedContent } from "./floats.js";
import type { ParagraphBox } from "./stack.js";
import type { Theme } from "../docx/theme.js";

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
  // The paragraph as it was laid out, whose lines already hold each drawing.
  readonly box: ParagraphBox;
  readonly resolvePart: PartResolver;
  readonly theme: Theme;
};

/**
 * An inline drawing is part of its line rather than something placed beside one,
 * so it lands wherever breaking the line put it: within the paragraph's own
 * indents, at its alignment, past whatever tabs opened ahead of it, and inside
 * the room an object wrapping beside the line left. It stands on the baseline, as
 * a letter does.
 *
 * A paragraph the page break ran through keeps its drawings on the page it starts
 * on, so a drawing whose line fell to the next page is not placed at all; no
 * reference document breaks a page through a paragraph holding one.
 */
export function placeInlines(input: PlaceInlinesInput): readonly PlacedInline[] {
  const placed: PlacedInline[] = [];

  for (const line of input.box.lines) {
    for (const segment of line.line.segments) {
      if (segment.kind !== "drawing") continue;
      const drawing = input.drawings[placed.length];
      if (drawing === undefined) return placed;
      placed.push({
        drawing,
        content: resolveContent(drawing.content, input.resolvePart, input.theme),
        leftPt: line.leftPt + segment.offsetPt,
        topPt: line.baselinePt - segment.heightPt,
        widthPt: segment.widthPt,
        heightPt: segment.heightPt,
      });
    }
  }

  return placed;
}
