import {
  aliasedSymbolText,
  type ParagraphBox,
  type ParagraphMark,
  type ParagraphMarker,
  type PlacedLine,
} from "@docx-pages/core";

import { bottomOf, upFromTop, type PdfPage } from "./coordinates.js";
import { faceOf, type PdfFonts, type PdfUnderline } from "./fonts.js";
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
  // Symbol faces the layout stood in for, by lowercased name, exactly as the
  // viewer takes them: a run written in one holds positions in that face's page
  // and is drawn as what those positions mean.
  readonly aliasSymbolFaces: ReadonlySet<string> | null;
};

// A run in a symbol face that was stood in for holds positions in that face's own
// page, and the stand-in would draw them as its own letters. Drawn as what the
// positions mean instead, which is how the layout measured them.
function shownText(
  mark: ParagraphMark,
  text: string,
  aliasFaces: ReadonlySet<string> | null,
): string {
  if (aliasFaces === null || mark.font.kind !== "named") return text;
  if (!aliasFaces.has(mark.font.name.trim().toLowerCase())) return text;
  return aliasedSymbolText(mark.font.name, text) ?? text;
}

// Black, which is what a run leaving its colour unstated is drawn in. Word states
// `auto` for text on a light ground and draws it black, and layout has already
// resolved anything else.
const DEFAULT_COLOR = "000000";

function shownRun(
  out: Content,
  options: TextOptions,
  mark: ParagraphMark,
  text: string,
  leftPt: number,
  baselinePt: number,
  widthPt: number,
): void {
  if (text === "") return;

  const face = options.fonts.faceFor(faceOf(mark));
  const glyphs = face.glyphsFor(shownText(mark, text, options.aliasSymbolFaces));

  out.fillColor(mark.color ?? DEFAULT_COLOR);
  out.beginText();
  out.font(face.resource, mark.fontSizePt);
  // Laid after every character of the run, the last one included, which is how
  // layout measured it and how a pdf's own character spacing behaves.
  out.characterSpacing(mark.characterSpacingPt);
  out.textPosition(leftPt, upFromTop(options.page, baselinePt));
  out.showGlyphs(glyphs);
  out.endText();

  if (mark.underline) underlined(out, options, face, mark, leftPt, baselinePt, widthPt);
}

/**
 * The line under an underlined run.
 *
 * A pdf has no such thing as an underline: the line is drawn, as Word draws it,
 * as a filled rectangle. **Where it goes is the face's own business** and not this
 * package's. Measured on 2026-08-07 off Word's own pdf of a reference document:
 * every underline there sat 0.1207 em below the baseline and was 0.0690 em thick,
 * the same at three places on the page, and those are the ratios the drawn face's
 * `post` table states rather than any constant Word carries.
 *
 * A face stating no `post` table gets no line, since nothing here could invent
 * where to put one and a line in the wrong place is worse than the run being
 * drawn without it. The README names it.
 */
function underlined(
  out: Content,
  options: TextOptions,
  face: { readonly underlineAt: (fontSizePt: number) => PdfUnderline | null },
  mark: ParagraphMark,
  leftPt: number,
  baselinePt: number,
  widthPt: number,
): void {
  const underline = face.underlineAt(mark.fontSizePt);
  if (underline === null || widthPt <= 0 || underline.thicknessPt <= 0) return;

  const topPt = baselinePt + underline.belowBaselinePt;
  out.fillColor(mark.color ?? DEFAULT_COLOR);
  out.rectangle(
    leftPt,
    bottomOf(options.page, topPt, underline.thicknessPt),
    widthPt,
    underline.thicknessPt,
  );
  out.fill();
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
      segment.widthPt,
    );
  }
}

// A list's number is drawn out of the text flow, at the position the level's
// hanging indent pulls the first line back to.
export function markerText(out: Content, options: TextOptions, marker: ParagraphMarker): void {
  shownRun(
    out,
    options,
    marker.mark,
    marker.text,
    marker.leftPt,
    marker.baselinePt,
    marker.widthPt,
  );
}

export function textOfBoxes(
  out: Content,
  options: TextOptions,
  boxes: readonly ParagraphBox[],
): void {
  for (const box of boxes) {
    if (box.marker !== null) markerText(out, options, box.marker);
    for (const line of box.lines) lineText(out, options, line);
  }
}
