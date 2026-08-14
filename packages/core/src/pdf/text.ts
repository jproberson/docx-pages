import type { ParagraphMark } from "../docx/styles.js";
import {
  drawnColor,
  faceAskedFor,
  type Drawable,
  type UnderlinePaint,
} from "../layout/drawables.js";
import type { ParagraphMarker, PlacedLine } from "../layout/stack.js";

import { bottomOf, upFromTop, type PdfPage } from "./coordinates.js";
import type { PdfFonts } from "./fonts.js";
import type { Content } from "./content.js";

// The text-showing half of a page, which mirrors the viewer's `textLayer` and
// answers the same question in the other notation.
//
// **Every run is written at the place layout put it** rather than being let run on
// from the one before. The viewer does the same, and for the same reason: a tab's
// gap and an inline picture leave a hole along the line that the runs after them
// carry, and a line laid end to end closes it. The advances inside a run are the
// face's own, and the face is the very one the line was measured with, so within a
// run the two agree exactly.

export type TextOptions = {
  readonly page: PdfPage;
  readonly fonts: PdfFonts;
};

// **What a run shows is decided in `drawables.ts`**, which is where a run in a
// stood-in symbol face has its positions turned into what they mean, and what an
// unstated colour comes to. The text and the colour reaching here are what is
// drawn.

function shownRun(
  out: Content,
  options: TextOptions,
  mark: ParagraphMark,
  text: string,
  leftPt: number,
  baselinePt: number,
): void {
  if (text === "") return;

  const face = options.fonts.faceFor(faceAskedFor(mark));
  const glyphs = face.glyphsFor(text);

  out.fillColor(drawnColor(mark.color));
  out.beginText();
  out.font(face.resource, mark.fontSizePt);
  // Laid after every character of the run, the last one included, which is how
  // layout measured it and how a pdf's own character spacing behaves.
  out.characterSpacing(mark.characterSpacingPt);
  // Every glyph drawn as wide a share of itself as the run states, which is what
  // the measurer multiplied its advance by.
  out.characterScale(mark.characterScale);
  out.textPosition(leftPt, upFromTop(options.page, baselinePt));
  out.showGlyphs(glyphs);
  out.endText();
}

// **Where an underline goes is decided in `drawables.ts`**, out of the metrics the
// drawn face states, and reaches here as the rectangle to fill. A pdf has no such
// thing as an underline and neither has Word: the line is filled.
function underlines(out: Content, page: PdfPage, drawn: readonly UnderlinePaint[]): void {
  for (const line of drawn) {
    out.fillColor(line.color);
    out.rectangle(
      line.leftPt,
      bottomOf(page, line.topPt, line.heightPt),
      line.widthPt,
      line.heightPt,
    );
    out.fill();
  }
}

/**
 * One line of a paragraph. A run is written where layout put it along the line,
 * and a raised or lowered run at the baseline its own mark asks for.
 */
export function lineText(out: Content, options: TextOptions, placed: PlacedLine): void {
  for (const segment of placed.line.segments) {
    if (segment.kind !== "text") continue;
    shownRun(
      out,
      options,
      segment.mark,
      segment.text,
      placed.leftPt + segment.offsetPt,
      placed.baselinePt - segment.mark.raisePt,
    );
  }
}

// A list's number is drawn out of the text flow, at the position the level's
// hanging indent pulls the first line back to.
export function markerText(out: Content, options: TextOptions, marker: ParagraphMarker): void {
  shownRun(out, options, marker.mark, marker.text, marker.leftPt, marker.baselinePt);
}

/**
 * Glyphs the drawing named by number rather than by character.
 *
 * Written exactly as text is, since that is what they are: the face is the one the
 * page already embeds, addressed by glyph as everything else in this file is, so a
 * shape with no character costs the same two bytes as a letter.
 *
 * Each glyph is written at the place layout put it rather than let run on from the
 * one before, as every run here is. The spacing and the scale are stated rather
 * than left as the last run set them: a stretched delimiter takes neither, and a
 * page whose previous run asked for either would otherwise carry it into this.
 */
export function drawnGlyphs(
  out: Content,
  options: TextOptions,
  drawable: Extract<Drawable, { kind: "glyphs" }>,
): void {
  if (drawable.glyphs.length === 0) return;

  const face = options.fonts.faceFor(drawable.face);

  out.fillColor(drawable.color);
  out.beginText();
  out.font(face.resource, drawable.sizePt);
  out.characterSpacing(0);
  out.characterScale(1);

  for (const glyph of drawable.glyphs) {
    out.textPosition(glyph.leftPt, upFromTop(options.page, glyph.baselinePt));
    out.showGlyphs(face.glyphNamed(glyph, drawable.sizePt));
  }
  out.endText();
}

export function textOfBoxes(
  out: Content,
  options: TextOptions,
  drawable: Extract<Drawable, { kind: "text" }>,
): void {
  for (const box of drawable.boxes) {
    if (box.marker !== null) markerText(out, options, box.marker);
    for (const line of box.lines) lineText(out, options, line);
  }
  // **A run standing in no paragraph is drawn exactly as a list's number is**, which
  // is what a piece of a set equation is: a string at a place at a size, already
  // placed by arithmetic that is not a line's.
  for (const run of drawable.runs) markerText(out, options, run);
  underlines(out, options.page, drawable.underlines);
}
