import type { DrawingFlip } from "../docx/drawing.js";
import type { InlineDrawing } from "../docx/inlines.js";
import { resolveContent, type PartResolver, type PlacedContent } from "./floats.js";
import type { ClipRect, ParagraphBox } from "./stack.js";
import { boundsOfTurn, unturnedRect } from "./turns.js";
import { emuToPoints } from "./units.js";
import type { Theme } from "../docx/theme.js";

export type PlacedInline = {
  readonly drawing: InlineDrawing;
  readonly content: PlacedContent;
  // Where the drawing stands before it is turned, which is the size it was stored
  // at seated in the middle of the room its line kept. The two are the same box
  // wherever nothing was turned.
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly turnDegrees: number;
  readonly flip: DrawingFlip;
  // What the paragraph lets through of a drawing hanging above it, and null wherever
  // the whole of it shows. See `cutToItsParagraph`.
  readonly clipTo: ClipRect | null;
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
/**
 * **A paragraph cuts off a drawing that hangs above it.** A line told exactly how tall
 * to be is a slot the drawing is dropped into, and what hangs out of the paragraph is
 * written into the page and painted nowhere, exactly as an exact row does to the text of
 * a cell.
 *
 * Measured on 2026-08-25 by `exact-line-clip-probe`, six cases read off Word's own pdf: a
 * picture 150 by 100 alone in paragraphs whose lines are told to be exactly 10, 20, 40,
 * 60, 100 and 120 showed **8.01, 16.01, 32.01, 48.01, 80.01 and 96.01** of its height,
 * every one of them starting at the paragraph's own top. The same picture under `atLeast`
 * and under no rule at all showed the whole of its 100, since those rules grow the line
 * to hold it.
 *
 * **It is the paragraph and not the line that cuts**, which those six could not tell
 * apart, since a picture alone in its paragraph makes the two the same box. Measured the
 * same day by `exact-line-paragraph-probe`: the same picture with one empty line above it
 * in the same paragraph showed **18.01**, which is that paragraph's two lines of 10 less
 * the 2 below its last baseline, and with twelve lines above it showed **all 100**. At an
 * exact 40 with one line above it, **72.01**.
 *
 * `10a0a7948de5` and `832cee47ded7` are the reading in the wild, and `91bb74cea83b` is
 * the reason this is the paragraph's box: a picture 531 by 108.75 hanging over eleven
 * empty lines of its own paragraph, which Word draws whole.
 *
 * The cut is up and down alone. Nothing has asked Word what a paragraph does to a drawing
 * wider than the text area, so the rectangle keeps the drawing's own left and width.
 *
 * **A turned drawing is measured by the paint it reaches and not by the box it was
 * stored in.** The flow keeps a turned drawing the room its turn rounds to, so a picture
 * on its side stands 307 by 167 in a paragraph 167 tall while the box it was stored in
 * is 167 by 307: read that stored box and every such drawing looks as though it hangs
 * out. `e199f3435eaf` and the two documents beside it are the reading, whose photograph
 * came out cut down its sides.
 */
const cutToItsParagraph = (box: ParagraphBox, drawing: ClipRect): ClipRect | null => {
  const first = box.lines[0];
  if (first === undefined) return null;
  const topPt = first.topPt + first.seatPt;
  const heightPt = Math.max(0, box.contentBottomPt - topPt);
  if (
    drawing.topPt >= topPt - EPSILON &&
    drawing.topPt + drawing.heightPt <= topPt + heightPt + EPSILON
  ) {
    return null;
  }
  return { leftPt: drawing.leftPt, topPt, widthPt: drawing.widthPt, heightPt };
};

const EPSILON = 1e-9;

export function placeInlines(input: PlaceInlinesInput): readonly PlacedInline[] {
  const placed: PlacedInline[] = [];

  for (const line of input.box.lines) {
    for (const segment of line.line.segments) {
      if (segment.kind !== "drawing") continue;
      const drawing = input.drawings[placed.length];
      if (drawing === undefined) return placed;
      const standing = unturnedRect(
        {
          leftPt: line.leftPt + segment.offsetPt,
          topPt: line.baselinePt - segment.heightPt,
          widthPt: segment.widthPt,
          heightPt: segment.heightPt,
        },
        {
          widthPt: emuToPoints(drawing.widthEmu),
          heightPt: emuToPoints(drawing.heightEmu),
        },
      );
      placed.push({
        drawing,
        content: resolveContent(drawing.content, input.resolvePart, input.theme),
        ...standing,
        turnDegrees: drawing.turnDegrees,
        flip: drawing.flip,
        clipTo: cutToItsParagraph(input.box, boundsOfTurn(standing, drawing.turnDegrees)),
      });
    }
  }

  return placed;
}
