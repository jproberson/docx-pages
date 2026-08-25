import type { BorderStyle } from "../docx/borders.js";
import type { DrawingFlip } from "../docx/drawing.js";
import type { ParagraphMark } from "../docx/styles.js";
import type { LaidOutDocument, LaidOutPage } from "./document.js";
import type { PlacedContent, PlacedFloat } from "./floats.js";
// The shape a glyph draws is the face's own answer, so it is declared with the
// reader that gets it out of the file rather than a second time here.
import type { GlyphOutline } from "./font-file.js";
import type { FaceRequest } from "./font-metrics.js";
import { mathPrimitivesOf, type MathPrimitive } from "./math.js";
import type { PlacedInline } from "./inlines.js";
import { paintOfCell, paintOfParagraph, type PaintedFill, type PaintedLine } from "./painting.js";
import type { ClipRect, ParagraphBox, ParagraphPaint, PlacedCell, PlacedLine } from "./stack.js";
import { aliasedSymbolText } from "./symbol-aliases.js";
import { turnedAbout } from "./turns.js";

// What a page draws and the order it draws it in, as one flat list. Layout says
// where everything sits; this says which of it is painted over which, which is
// the same question whatever the drawing is done with. Every backend walks this,
// so a rule about stacking is settled here once rather than answered twice.

/**
 * Where a drawn line lands down the page: a three-hundredth of an inch, which is
 * the device Word draws every page in.
 *
 * **Measured on 2026-08-14 over Word's own pdf of 132 authored documents and 180
 * corpus ones.** Word wraps each page in `0.24 0 0 0.24 x y cm` and counts
 * everything inside it in those units: 32276 of 32308 such transforms are exactly
 * that, and the page box itself is a whole number of them. Every baseline Word
 * writes is a whole unit (100.0% over 23813 letter placements and 8562 A4 ones),
 * and so is every type size, while a position **across** the page is not (28.2%
 * and 3.8%) and carries the exact sum of the face's own advances: `abcdef` in
 * Calibri moved 33.064464pt where the face's 5643 units come to 33.064453.
 *
 * **The height itself is not on the grid, so the layout must not be.** 31 runs of
 * fifteen or more solid lines in Word's own exports mix two gaps exactly one step
 * apart, which a height in whole steps cannot draw: 12pt Calibri lines come out
 * 14.64 and 14.88 apart about an exact 14.6484375, and a 20pt rule comes out 19.92
 * and 20.16 about an exact 20. So the layout keeps its exact arithmetic, as Word's
 * does, and only what is drawn lands here.
 *
 * **Applied once, and only here.** Neither backend rounds again: the pdf writer
 * subtracts a snapped distance from the page height to flip it, and the viewer
 * writes it down the page as it stands.
 */
const DEVICE_UNITS_TO_THE_INCH = 300;
const POINTS_TO_THE_INCH = 72;

// Written as the two divisions rather than as 0.24 so that a snapped number comes
// back as the number a reader of the page would write down: 198 units is 47.52,
// not 47.519999999999996.
export const onTheDeviceGrid = (downPt: number): number =>
  (Math.round((downPt * DEVICE_UNITS_TO_THE_INCH) / POINTS_TO_THE_INCH) * POINTS_TO_THE_INCH) /
  DEVICE_UNITS_TO_THE_INCH;

/**
 * What the machine could not do for itself, which a caller states and this decides
 * what to do about.
 *
 * A backend is handed the page as it is drawn; what a face was stood in for is a
 * fact about the machine the page was laid out on, and only the caller knows it.
 */
export type DrawingOptions = {
  // Symbol faces the layout stood in for, by lowercased name. A run written in one
  // holds positions in that face's own page, and the stand-in would draw them as
  // its own letters.
  readonly aliasSymbolFaces?: ReadonlySet<string> | null;
  // Where the drawn face puts the line under its own letters, which only a caller
  // holding the face can say. One that cannot gets no rectangle and has to draw
  // the line whatever way it can.
  readonly underlineFor?: (mark: ParagraphMark) => UnderlineMetrics | null;
};

// What a face states about the line under its letters, at the size a run is set
// at: how far below the baseline its top sits, and how thick it is.
export type UnderlineMetrics = {
  readonly belowBaselinePt: number;
  readonly thicknessPt: number;
};

/**
 * What a run of text actually shows.
 *
 * A run in a symbol face that was stood in for holds positions in that face's own
 * page, and the stand-in would draw them as its own letters. Drawn as what the
 * positions mean instead, which is how the layout measured them.
 *
 * **Decided here and not in a renderer.** Both backends worked this out for
 * themselves until 2026-08-14, from the same set and with the same words, which is
 * two chances to get it wrong and no way for a third backend to know it was a
 * question at all.
 */
function shownText(mark: ParagraphMark, text: string, options: DrawingOptions): string {
  const aliasFaces = options.aliasSymbolFaces;
  if (aliasFaces === undefined || aliasFaces === null || mark.font.kind !== "named") return text;
  if (!aliasFaces.has(mark.font.name.trim().toLowerCase())) return text;
  return aliasedSymbolText(mark.font.name, text) ?? text;
}

/**
 * A run's mark as it is drawn: the colour resolved, and the size on the same device
 * grid the baseline lands on.
 *
 * **Word sets every run at a whole number of device units.** Measured on 2026-08-14
 * over Word's own pdf of 138 authored documents: 10748 of 10748 drawn items, 100.00%.
 * Word writes the size into the text matrix rather than into `Tf`, as `50 0 0 50 x y
 * Tm` with `/TT2 1 Tf`, and those 50 units are the 12pt the document stated; an 11pt
 * run comes out at 46 units, which is 11.04, and a 20pt one at 83, which is 19.92.
 * The sizes Word drew across those documents are 12, 24, 7.92, 11.04, 13.92, 10.56,
 * 8.4, 19.92, 48, 12.96, and every one of them is `round(stated / 0.24) * 0.24`.
 *
 * **The advances do not follow the size, and that is why this is a drawing rule.**
 * Word writes the glyphs of a run in a `TJ` array and nudges each one back by a few
 * thousandths, so the run still advances by the exact sum of the face's integer
 * advances at the size the document stated: an 11pt run measured 36.94 by that sum
 * moved 36.96 and not the 37.07 the drawn size would give. So the layout goes on
 * measuring at the stated size, and only what is drawn lands here.
 *
 * **What that costs us.** The viewer holds a run to the width it was measured at, so
 * there it costs nothing. The pdf writer shows a run in one `Tj` and does not, so
 * inside a run its glyphs advance by the drawn size: a third of a percent of that
 * run's own width, which is 0.13pt at the end of a 37pt run and nothing at its
 * start, since every run is written at the place layout put it. Closing that is
 * `runWidthMadeUpBy` reaching the pdf writer, the way Word closes it with `TJ`.
 */
const drawnMark = (mark: ParagraphMark): ParagraphMark => ({
  ...mark,
  fontSizePt: onTheDeviceGrid(mark.fontSizePt),
  color: drawnColor(mark.color),
});

// A line as it is drawn. Both edges land on the grid and the height is what lies
// between them, so what is painted behind a line keeps the line's own foot instead
// of drifting a step off it.
function drawnLine(line: PlacedLine, options: DrawingOptions): PlacedLine {
  const topPt = onTheDeviceGrid(line.topPt);
  return {
    ...line,
    topPt,
    baselinePt: onTheDeviceGrid(line.baselinePt),
    fittingHeightPt: onTheDeviceGrid(line.topPt + line.fittingHeightPt) - topPt,
    line: {
      ...line.line,
      segments: line.line.segments.map((segment) =>
        segment.kind === "text"
          ? {
              ...segment,
              text: shownText(segment.mark, segment.text, options),
              mark: drawnMark(segment.mark),
            }
          : segment,
      ),
    },
  };
}

/**
 * A story's paragraphs as they are drawn.
 *
 * **One snapped copy answers for everything drawn about them**: the text, the fill
 * and border a paragraph asks for, and the highlight under a run are all worked
 * out from these very numbers further down, so none of them can part company with
 * the others. Only what is drawn is moved; what the layout used to decide where a
 * line went is left exactly as it was.
 */
const drawnBoxes = (
  boxes: readonly ParagraphBox[],
  options: DrawingOptions,
): readonly ParagraphBox[] =>
  boxes.map((box) => ({
    ...box,
    markTopPt: onTheDeviceGrid(box.markTopPt),
    contentBottomPt: onTheDeviceGrid(box.contentBottomPt),
    marker:
      box.marker === null
        ? null
        : {
            ...box.marker,
            baselinePt: onTheDeviceGrid(box.marker.baselinePt),
            text: shownText(box.marker.mark, box.marker.text, options),
            mark: drawnMark(box.marker.mark),
          },
    lines: box.lines.map((line) => drawnLine(line, options)),
  }));

/**
 * How a backend makes up the difference between the width a run was measured at
 * and the width the face it draws with comes to.
 *
 * The line was measured with the authored face's own widths. Holding each run to
 * that width keeps the break points Word chose even when the page draws the text
 * in a substitute, and starting each one where layout put it keeps a tab's gap and
 * an inline picture from carrying the rest of the line along with them.
 *
 * **A run the file scaled is stretched rather than spaced out.** Holding a run to
 * its measured width by the gaps between its glyphs is what keeps Word's break
 * points under a substituted face, but a run stating `w:w` is drawn wider or
 * narrower glyph by glyph, which is what the pdf writer's `Tz` does and what a
 * viewer asks for by name.
 *
 * **Decided here and not in a renderer.** A backend embedding the very face the
 * line was measured with needs none of this and the pdf writer does not ask; one
 * drawing in whatever face it can find needs all of it, and until 2026-08-14 the
 * only statement of it was inside that one backend, where no other could find it.
 */
export type WidthMadeUpBy = "spacing" | "glyphs";

export const runWidthMadeUpBy = (mark: ParagraphMark): WidthMadeUpBy =>
  mark.characterScale === 1 ? "spacing" : "glyphs";

/**
 * How far a metafile's own pen stands from the line it is told to draw.
 *
 * A metafile's pen hangs its line off the cell whose corner the line runs through,
 * where a stroke is centred on the line it follows, so the line moves half a unit
 * down and right to cover the same cells.
 *
 * **Decided here and not in a renderer**, which both of them did until 2026-08-14.
 * It belongs further back still, in the reader that turns a recording into shapes:
 * a shape reaching a backend already standing where it is drawn would need no
 * constant here at all. That reader is `metafile/picture.ts`.
 */
export const METAFILE_PEN_OFFSET = 0.5;

// A cell's floor is the ceiling of the cell under it, so both edges go on the grid
// and the height is the distance between them.
function drawnCell(cell: PlacedCell): PlacedCell {
  const topPt = onTheDeviceGrid(cell.topPt);
  return { ...cell, topPt, heightPt: onTheDeviceGrid(cell.topPt + cell.heightPt) - topPt };
}

// A paragraph's own fill and border, with the room its lines took: how far up and
// down they reach is the paragraph's, how far across is the text area's.
export type PaintedParagraph = {
  readonly paint: ParagraphPaint;
  readonly topPt: number;
  readonly bottomPt: number;
};

// What is painted behind one run of a line, which is the run's own advance across
// and the line's box down.
export type HighlightPaint = {
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly color: string;
};

/**
 * The line under an underlined run, as the rectangle it is drawn as.
 *
 * A pdf has no such thing as an underline and neither has Word: the line is
 * filled. **Where it goes is the drawn face's own business** and not a renderer's.
 * Measured on 2026-08-07 off Word's own pdf of a reference document: every
 * underline there sat 0.1207 em below the baseline and was 0.0690 em thick, the
 * same at three places on the page, and those are the ratios the drawn face's
 * `post` table states rather than any constant Word carries.
 *
 * A face stating no `post` table gets no rectangle, since nothing here could
 * invent where to put one and a line in the wrong place is worse than a run drawn
 * without it. The README names it. A backend holding no faces at all gets none
 * either, and draws the line whatever way it can.
 */
export type UnderlinePaint = HighlightPaint;

/**
 * How Word draws each pattern, measured at a width of a point and a half: a dashed
 * line runs four widths on and four off, where a dotted one runs one and one. A
 * double line is two bands, which the geometry has already made of it.
 *
 * **Decided here and not in a renderer.** Both backends carried this table until
 * 2026-08-14, which is two places for a measurement to drift from itself.
 */
const DASHES: Readonly<Record<BorderStyle, readonly number[] | null>> = {
  single: null,
  double: null,
  dashed: [4, 4],
  dotted: [1, 1],
};

// A band of a border as it is drawn: where the geometry puts it, and how the
// dashes fall along it at the width it is drawn at.
export type DrawnLine = Omit<PaintedLine, "color"> & {
  readonly dashes: readonly number[] | null;
  readonly color: string;
};

// One thing's paint: what it fills, and the bands it draws round itself. Kept
// apart from the next thing's so that what covers what is the order they are in.
export type DrawnPaint = {
  readonly fills: readonly PaintedFill[];
  readonly lines: readonly DrawnLine[];
};

const drawnPaint = (painted: {
  fills: readonly PaintedFill[];
  lines: readonly PaintedLine[];
}): DrawnPaint => ({
  fills: painted.fills,
  lines: painted.lines.map((line) => {
    const dashes = DASHES[line.style];
    return {
      ...line,
      color: drawnColor(line.color),
      dashes: dashes === null ? null : dashes.map((each) => each * line.widthPt),
    };
  }),
});

/**
 * One glyph of a face, named by its number in that face rather than by a
 * character.
 *
 * **The only way to ask for a shape that has no character.** The parenthesis Word
 * stretches round a fraction is the fourth taller variant of `(` in Cambria Math,
 * a glyph in no character map at all: measured on 2026-08-13 off Word's own pdf,
 * where it came back 21.60pt of continuous ink, and read out of the face's own
 * MATH table, which names it as glyph 3436 and states it reaches 4047 units.
 *
 * Everything else on a page is asked for in characters, and this is the one thing
 * that cannot be.
 */
export type DrawnGlyph = {
  readonly glyph: number;
  readonly leftPt: number;
  readonly baselinePt: number;
  // What the glyph draws, where whoever built the run could read it out of the
  // face. Absent where nothing did, and a backend with no other way to draw the
  // glyph then says so rather than drawing something else.
  readonly outline?: GlyphOutline;
  // What the glyph advances at the size it is drawn, which the face's own metrics
  // state and the layout measured with. It comes with the glyph because a glyph
  // with no character has no advance anything else here could look up: an advance
  // table answers by character.
  readonly advancePt: number;
  // The character the glyph stands for, where it stands for one, so that a reader
  // of the page can still select and search the text. **It is not what to draw
  // instead.** A stretched parenthesis drawn as a plain one is the right character
  // at the wrong height, which is a page that looks finished and is wrong; a
  // backend that cannot name a glyph draws nothing and says so.
  readonly standsFor: string | null;
};

// A stretch of glyphs of one face at one size, which is what the layout hands over
// for a shape it could not name in characters.
export type PlacedGlyphs = {
  readonly face: FaceRequest;
  readonly sizePt: number;
  // Six hex digits, as every other colour reaching a backend is.
  readonly color: string;
  // How far the glyphs reach above and below their baseline, which is the ink they
  // draw rather than the face's own ascent: it is what a backend that cannot draw
  // them shows the room of, and what a caller cutting the page to a box measures.
  readonly ascentPt: number;
  readonly descentPt: number;
  readonly glyphs: readonly DrawnGlyph[];
};

/**
 * A run of text drawn at a place of its own rather than along a line.
 *
 * **This is the shape a list's number already had**, which is why nothing below
 * learns a new way to draw text: a number pulled out of the flow and a fraction's
 * numerator are the one question, a string at a place at a size, and both
 * backends answer it in a function they already have.
 *
 * A run inside a line is placed by the line: it carries an offset along it and
 * takes the line's own baseline, raised by whatever its mark asks for. A run here
 * carries the whole answer, because the arithmetic that placed it is not the
 * line's.
 */
export type DrawnRun = {
  readonly text: string;
  readonly mark: ParagraphMark;
  // What the run was measured at, which a backend drawing it in another face holds
  // it to; see `runWidthMadeUpBy` for how the difference is made up.
  readonly widthPt: number;
  readonly leftPt: number;
  readonly baselinePt: number;
};

/**
 * A set equation as the drawing takes it: the pieces `math.ts` placed, and how far
 * the whole of it reaches either side of the line's own baseline.
 *
 * **The pieces are the seam between the arithmetic and the drawing**, and `math.ts`
 * declares them: it sets a fraction measured from the box's own baseline with up
 * positive, which is how the MATH table states every constant it works from, and
 * `mathPrimitivesOf` turns that into a list in the page's own coordinates, down
 * positive, in the order the pieces are painted. The axis turns over once, there,
 * and nothing here knows the other way up exists.
 *
 * The ink comes from the line rather than from the pieces, because it is the only
 * thing a glyph run needs that a piece does not yet carry: see `drawnMathGlyphs`.
 */
export type SetEquation = {
  readonly primitives: readonly MathPrimitive[];
  readonly ascentPt: number;
  readonly descentPt: number;
};

/**
 * The mark one piece of an equation is drawn with: the equation's own face, weight,
 * slant and colour, at the size the flattener set that piece at.
 *
 * **What a run states about spacing and scale is left off.** Both would move the
 * glyphs off the places `math.ts` measured, which it measures from the face's plain
 * advances, and the geometry standing round them is made of those same advances.
 * The glyph run states the same two for the same reason.
 *
 * A raise is left off because the flattener has already placed the piece: the
 * baseline reaching here is where the piece goes, raise and all. An underline and a
 * highlight are left off because the flattener states neither, and one backend
 * drawing what the other cannot is worse than neither drawing it.
 */
const markOfPiece = (mark: ParagraphMark, sizePt: number): ParagraphMark => ({
  ...drawnMark({ ...mark, fontSizePt: sizePt }),
  characterSpacingPt: 0,
  characterScale: 1,
  raisePt: 0,
  underline: false,
  highlight: null,
});

// The face a run asks to be drawn in, which is the three things a document states
// about it. **Decided here** because it is a fact about what is drawn, and both the
// pdf writer and the metafile player ask in this one shape.
export const faceAskedFor = (mark: ParagraphMark): FaceRequest => ({
  name: mark.font.kind === "named" ? mark.font.name : "",
  bold: mark.bold,
  italic: mark.italic,
});

/**
 * Where a set equation lands down the page, which is the rule for all three of the
 * pieces below.
 *
 * **Every piece goes on the grid on its own, and Word's own pdf is what says so.**
 * Measured on 2026-08-14 over Word's exports of the two authored equation probes:
 * all 301 text baselines on those pages are a whole device unit, the numerator and
 * the denominator of every fraction among them, and so are both edges of all 72
 * filled rectangles and every one of their heights. Across the page nothing is:
 * 45.2% of the lefts, which is the exact arithmetic the rest of the page keeps.
 *
 * So a fraction is **not** moved as one thing with its offsets kept exact. The bar
 * keeps its place against its halves because the three of them land on the one
 * grid, not because the distances between them survive. Snapping the origin alone
 * would leave every baseline inside the fraction off the grid Word has them on, and
 * snapping a bar's thickness rather than its two edges would let its foot drift off
 * the unit its head sits on: the bar is filled between two snapped edges, which is
 * what `drawnLine` and `drawnCell` already do with everything else that is filled.
 */
const drawnMathRun = (piece: Extract<MathPrimitive, { kind: "text" }>): DrawnRun => ({
  text: piece.text,
  mark: markOfPiece(piece.mark, piece.sizePt),
  // **A piece states no width, so none is held to one.** Every other run is held to
  // what it was measured at so that a stand-in face keeps Word's break points; a set
  // equation is drawn in the very face it was measured in or it is not set at all,
  // since a face stating no MATH table cannot set one. What is left is a browser
  // substituting under the viewer, where the halves would drift from their bar; the
  // piece carrying its own `box.widthPt` is what would close that.
  widthPt: 0,
  leftPt: piece.leftPt,
  baselinePt: onTheDeviceGrid(piece.baselinePt),
});

const drawnMathFill = (piece: Extract<MathPrimitive, { kind: "fill" }>): PaintedFill => {
  const topPt = onTheDeviceGrid(piece.topPt);
  return {
    // A bar takes its colour from the mark the piece carries, which for a fraction
    // is the `m:ctrlPr`'s own where the file wrote one.
    color: drawnColor(piece.mark.color),
    leftPt: piece.leftPt,
    topPt,
    widthPt: piece.widthPt,
    heightPt: onTheDeviceGrid(piece.topPt + piece.heightPt) - topPt,
  };
};

/**
 * A stretched delimiter, which is a rung of the face's own ladder and has no
 * character to be asked for by: the glyph run was built for exactly this.
 *
 * **Three things a glyph needs are not on the piece yet**, and each costs something
 * stated rather than hidden. Its advance is what the pdf writes as the glyph's own
 * width, since an advance table answers by character and can answer for none of
 * these; nought there costs a reader the width when it selects the shape, and never
 * moves it, because every glyph is drawn at a place of its own. What it stands for
 * is what lets a reader search the page for the bracket. Its outline is the whole of
 * what a browser can draw, so the viewer marks it undrawn and says which it was.
 *
 * All three are on `PlacedDelimiter` and `MathVariant` already, an argument away
 * from the flattener that placed the piece.
 */
const drawnMathGlyph = (piece: Extract<MathPrimitive, { kind: "glyph" }>): DrawnGlyph => ({
  glyph: piece.glyph,
  leftPt: piece.leftPt,
  baselinePt: onTheDeviceGrid(piece.baselinePt),
  advancePt: 0,
  standsFor: null,
});

// Whether a piece is drawn together with the one before it. Kinds part company
// because each is a drawable of its own, and two runs of glyphs part company at a
// change of size because a run of them is drawn at one size.
function joinsThePieceBefore(head: MathPrimitive, next: MathPrimitive): boolean {
  if (head.kind !== next.kind) return false;
  if (head.kind === "glyph" && next.kind === "glyph") return head.sizePt === next.sizePt;
  return true;
}

// **Consecutive pieces of one kind are drawn together, and the runs stay in the
// order the flattener gave them**, so what an equation paints over what is settled
// by that order rather than by gathering every fill of an equation into one layer.
function piecesInTheirOrder(
  primitives: readonly MathPrimitive[],
): readonly (readonly MathPrimitive[])[] {
  const groups: MathPrimitive[][] = [];
  for (const primitive of primitives) {
    const open = groups[groups.length - 1];
    const head = open?.[0];
    if (open === undefined || head === undefined || !joinsThePieceBefore(head, primitive)) {
      groups.push([primitive]);
      continue;
    }
    open.push(primitive);
  }
  return groups;
}

function drawableOfPieces(
  pieces: readonly MathPrimitive[],
  ink: { readonly ascentPt: number; readonly descentPt: number },
  key: string,
): readonly Drawable[] {
  const head = pieces[0];
  if (head === undefined) return [];

  switch (head.kind) {
    case "text": {
      const runs = pieces
        .filter((piece): piece is Extract<MathPrimitive, { kind: "text" }> => piece.kind === "text")
        .map(drawnMathRun);
      return [{ kind: "text", key, boxes: [], runs, underlines: [], clipTo: null, turnDegrees: 0 }];
    }
    case "fill": {
      const fills = pieces
        .filter((piece): piece is Extract<MathPrimitive, { kind: "fill" }> => piece.kind === "fill")
        .map(drawnMathFill);
      return [{ kind: "paint", key, painted: [{ fills, lines: [] }], highlights: [] }];
    }
    case "glyph": {
      const drawn = pieces.filter(
        (piece): piece is Extract<MathPrimitive, { kind: "glyph" }> => piece.kind === "glyph",
      );
      return [
        {
          kind: "glyphs",
          key,
          face: faceAskedFor(head.mark),
          sizePt: onTheDeviceGrid(head.sizePt),
          color: drawnColor(head.mark.color),
          // The equation's own reach, which is the nearest thing to the glyph's ink
          // that gets this far: what a backend that cannot draw the shape shows the
          // room of, and a delimiter is what makes an equation as tall as it is.
          ascentPt: ink.ascentPt,
          descentPt: ink.descentPt,
          glyphs: drawn.map(drawnMathGlyph),
        },
      ];
    }
  }
}

/**
 * A set equation as the drawables that draw it.
 *
 * **Nothing here is new to a backend.** A piece of text is the run a list's number
 * already is, a fraction's bar is a fill like any other, and a stretched delimiter
 * is the glyph run built for exactly this: a shape with no character to ask for it
 * by. What this settles is which of the three each piece is and where it lands, and
 * a backend that draws a page already draws all three.
 */
export const mathDrawables = (equation: SetEquation, key: string): readonly Drawable[] =>
  piecesInTheirOrder(equation.primitives).flatMap((pieces, at) =>
    drawableOfPieces(pieces, equation, `${key}-${String(at)}`),
  );

export type Drawable =
  | {
      readonly kind: "object";
      readonly key: string;
      readonly name: string;
      readonly content: PlacedContent;
      readonly leftPt: number;
      readonly topPt: number;
      readonly widthPt: number;
      readonly heightPt: number;
      // Which way round the shape was turned, which a connector inside a group
      // states and decides which corners its line runs between.
      readonly flip: DrawingFlip;
      // How far clockwise the shape is turned about the middle of the box above,
      // which is the box it stands in before it is turned.
      readonly turnDegrees: number;
      // What the line a drawing stands in lets through of it, and null for anything
      // the whole of which is drawn. `placeInlines` records why.
      readonly clipTo: ClipRect | null;
    }
  | {
      readonly kind: "text";
      readonly key: string;
      readonly boxes: readonly ParagraphBox[];
      // Text this layer draws that stands in no paragraph of it: the pieces of a
      // set equation, which are placed by their own arithmetic rather than by a
      // line. Drawn after the boxes and cut and turned with them.
      readonly runs: readonly DrawnRun[];
      // Drawn after the text of this same layer, which is where the run they
      // belong to puts them.
      readonly underlines: readonly UnderlinePaint[];
      // The rectangle the text is cut to, which a shape's own text has and the
      // text a story flowed down the page does not.
      readonly clipTo: Rect | null;
      // A shape's text is turned with the shape, about the middle of that same
      // rectangle. Text a story flowed down the page is never turned.
      readonly turnDegrees: number;
    }
  // Glyphs named by number, drawn where the text of the story is drawn: they are
  // text, and what stands over them and under them is what stands over and under
  // any other text.
  | ({ readonly kind: "glyphs"; readonly key: string } & PlacedGlyphs)
  // Everything drawn behind the text of a story: the cells of its tables and the
  // fills and borders its paragraphs ask for. One layer holds them all, since they
  // are drawn in the page's own coordinates and nothing stands between them.
  | {
      readonly kind: "paint";
      readonly key: string;
      // The cells of the story's tables first, then what each paragraph asks for,
      // which Word draws over the cell holding it. Each keeps its own fills and
      // bands together, since that is what settles which of them covers which.
      readonly painted: readonly DrawnPaint[];
      // Painted over both of those and under the text, which is where Word puts a
      // highlight: measured against a shaded paragraph holding a highlighted run,
      // whose fill Word drew first and the highlight over it.
      readonly highlights: readonly HighlightPaint[];
    };

export type Rect = {
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

// An object standing somewhere, before it is known whether it is one thing to
// draw or a group of them.
type Standing = {
  readonly name: string;
  readonly content: PlacedContent;
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly flip: DrawingFlip;
  readonly turnDegrees: number;
  readonly clipTo: ClipRect | null;
};

/**
 * What a painter walks for one object, which for a group is one item per shape
 * inside it and for a group inside a group is that again.
 *
 * **A group is flattened here and nowhere else.** Each child keeps the fraction of
 * its group's box it stands in, so multiplying that by the room the flow gave the
 * group is the whole of the arithmetic, and it is the same arithmetic at every
 * depth. A renderer therefore never learns that a group exists.
 */
function objectsOf(standing: Standing, key: string, options: DrawingOptions): readonly Drawable[] {
  const { content } = standing;
  if (content.kind === "group") {
    // A group turned as a whole carries its children round with it, so each one is
    // swung about the group's middle and turned that much further itself.
    const middle = {
      xPt: standing.leftPt + standing.widthPt / 2,
      yPt: standing.topPt + standing.heightPt / 2,
    };
    return content.children.flatMap((child, at) => {
      const widthPt = child.widthFraction * standing.widthPt;
      const heightPt = child.heightFraction * standing.heightPt;
      const stands = turnedAbout(
        {
          xPt: standing.leftPt + child.leftFraction * standing.widthPt + widthPt / 2,
          yPt: standing.topPt + child.topFraction * standing.heightPt + heightPt / 2,
        },
        middle,
        standing.turnDegrees,
      );
      return objectsOf(
        {
          name: standing.name,
          content: child.content,
          leftPt: stands.xPt - widthPt / 2,
          topPt: stands.yPt - heightPt / 2,
          widthPt,
          heightPt,
          flip: child.flip,
          turnDegrees: standing.turnDegrees + child.turnDegrees,
          clipTo: standing.clipTo,
        },
        `${key}-${String(at)}`,
        options,
      );
    });
  }

  return [{ kind: "object", key, ...standing }, ...textOf(standing, key, options)];
}

const fromFloat = (float: PlacedFloat, key: string, options: DrawingOptions): readonly Drawable[] =>
  objectsOf(
    {
      name: float.anchor.name,
      content: float.content,
      leftPt: float.leftPt,
      topPt: float.topPt,
      widthPt: float.widthPt,
      heightPt: float.heightPt,
      flip: float.flip,
      turnDegrees: float.turnDegrees,
      // A floating drawing stands beside the lines rather than in one, so no line
      // cuts it.
      clipTo: null,
    },
    key,
    options,
  );

const fromInline = (
  inline: PlacedInline,
  key: string,
  options: DrawingOptions,
): readonly Drawable[] =>
  objectsOf(
    {
      name: inline.drawing.name,
      content: inline.content,
      leftPt: inline.leftPt,
      topPt: inline.topPt,
      widthPt: inline.widthPt,
      heightPt: inline.heightPt,
      flip: inline.flip,
      turnDegrees: inline.turnDegrees,
      clipTo: inline.clipTo,
    },
    key,
    options,
  );

const hasText = (boxes: readonly ParagraphBox[]): boolean =>
  boxes.some((box) => box.lines.length > 0);

// A text box draws its own text straight after its frame, so the two keep the one
// place in the stack that Word gave the shape.
//
// Word cuts that text off at the frame. A box told to fit itself to its text was
// grown to hold all of it and loses nothing, but one that was not keeps the size
// it was stored at and shows only as much as fits: the rest is not moved
// anywhere, it is simply not drawn.
function textOf(standing: Standing, key: string, options: DrawingOptions): readonly Drawable[] {
  const { content } = standing;
  if (content.kind !== "text-box" || content.text === null) return [];

  const clipTo = {
    leftPt: standing.leftPt,
    topPt: standing.topPt,
    widthPt: standing.widthPt,
    heightPt: standing.heightPt,
  };
  // A turned shape's text is drawn under a turn of its own, so the page's own
  // grid is not the grid it lands on and it is left exactly where layout put it.
  const square = standing.turnDegrees === 0;
  const boxes = square ? drawnBoxes(content.text.boxes, options) : content.text.boxes;
  const cells = square ? content.text.cells.map(drawnCell) : content.text.cells;
  const { inlines } = content.text;
  const painted = paintLayer(cells, boxes, `${key}-paint`);
  const turnDegrees = standing.turnDegrees;
  return [
    ...painted,
    ...(hasText(boxes)
      ? [
          {
            kind: "text" as const,
            key: `${key}-text`,
            boxes,
            runs: [],
            underlines: underlinesIn(boxes, options),
            clipTo,
            turnDegrees,
          },
        ]
      : []),
    // A box's own text sets its equations as the flow does, and they are drawn where
    // that text is drawn. Nothing cuts them to the frame the text is cut to yet.
    ...equationLayers(boxes, `${key}-equation`),
    // **A drawing standing in the box's own text is drawn where that text put it.**
    // Fifteen corpus documents hold 76 of them and not one was drawn until
    // 2026-08-14: the text around each came out as usual, so no page said anything
    // was missing.
    ...inlines.flatMap((inline, at) => fromInline(inline, `${key}-inline-${String(at)}`, options)),
  ];
}

// Everything drawn behind a story's text, where there is anything at all.
function paintLayer(
  cells: readonly PlacedCell[],
  boxes: readonly ParagraphBox[],
  key: string,
): readonly Drawable[] {
  const paragraphs = paintedParagraphs(boxes);
  const highlights = highlightsIn(boxes);
  if (cells.length + paragraphs.length + highlights.length === 0) return [];

  const painted = [
    ...cells.map((cell) => drawnPaint(paintOfCell(cell))),
    ...paragraphs.map((each) =>
      drawnPaint(paintOfParagraph(each.paint, each.topPt, each.bottomPt)),
    ),
  ];
  return [{ kind: "paint", key, painted, highlights }];
}

function underlinesIn(
  boxes: readonly ParagraphBox[],
  options: DrawingOptions,
): readonly UnderlinePaint[] {
  const stated = options.underlineFor;
  if (stated === undefined) return [];

  const under = (
    mark: ParagraphMark,
    leftPt: number,
    baselinePt: number,
    widthPt: number,
  ): readonly UnderlinePaint[] => {
    if (!mark.underline || widthPt <= 0) return [];
    const line = stated(mark);
    if (line === null || line.thicknessPt <= 0) return [];
    return [
      {
        leftPt,
        topPt: baselinePt + line.belowBaselinePt,
        widthPt,
        heightPt: line.thicknessPt,
        color: drawnColor(mark.color),
      },
    ];
  };

  return boxes.flatMap((box) => [
    ...(box.marker === null
      ? []
      : under(box.marker.mark, box.marker.leftPt, box.marker.baselinePt, box.marker.widthPt)),
    ...box.lines.flatMap((placed) =>
      placed.line.segments.flatMap((segment) =>
        segment.kind !== "text"
          ? []
          : under(
              segment.mark,
              placed.leftPt + segment.offsetPt,
              placed.baselinePt - segment.mark.raisePt,
              segment.widthPt,
            ),
      ),
    ),
  ]);
}

/**
 * What anything leaving its colour unstated is drawn in.
 *
 * **Black.** Word states `auto` for text on a light ground and draws it black, and
 * layout has already resolved anything else; the colour Word draws an unstated
 * border in is black as well.
 *
 * **Decided here and not in a renderer**, which is where the two of them
 * disagreed until 2026-08-14: the pdf writer drew black and the viewer inherited
 * whatever colour the page around it was set in, so a page in a themed container
 * drew text Word would have drawn black. Written with its hash, which is the
 * spelling everything else reaching a backend uses.
 */
export const UNSTATED_COLOR = "#000000";

export const drawnColor = (stated: string | null): string => stated ?? UNSTATED_COLOR;

/**
 * **A highlight is the run's own advance across and the line's box down**, whatever
 * size the run itself is set at. Measured 2026-08-13 against Word's own pdf, six
 * cases of one highlighted word:
 *
 * | the line                          | Word painted |
 * | --------------------------------- | ------------ |
 * | 12pt throughout                   | 14.64pt tall, the 12pt line |
 * | 24pt throughout                   | 29.28pt      |
 * | a 12pt run on a line holding 24pt | 29.28pt, the whole line and not the run |
 * | a superscript run                 | 14.64pt, the line rather than the raised text |
 * | an exact rule of 24pt under 12pt  | 24.00pt, the whole of what the rule asked for |
 * | a line multiple of two            | 14.64pt, the text's own box and not the room |
 *
 * So the room a multiple opens below the text is not painted and the slot an exact
 * rule states is, which is exactly the height the line has to be given to stay on a
 * page: `fittingHeightPt` answers both without asking what the rule was.
 *
 * Nothing is painted for an empty paragraph whose mark states a highlight: Word's
 * pdf of one holds no fill at all.
 */
function highlightsIn(boxes: readonly ParagraphBox[]): readonly HighlightPaint[] {
  return boxes.flatMap((box) =>
    box.lines.flatMap((placed) => {
      const topPt = placed.topPt;
      const heightPt = placed.fittingHeightPt;
      return placed.line.segments.flatMap((segment) => {
        if (segment.kind !== "text" || segment.mark.highlight === null) return [];
        return [
          {
            leftPt: placed.leftPt + segment.offsetPt,
            topPt,
            widthPt: segment.widthPt,
            heightPt,
            color: segment.mark.highlight,
          },
        ];
      });
    }),
  );
}

// Word stacks the floats of one story by relativeHeight alone. It is not two
// layers either side of the text: these documents send a filled panel to the back
// of the stack and then draw a box marked behindDoc over it, so the height a
// shape was given is the whole of the order within a story.
//
// The stories themselves are layered, and the header and the footer lie under the
// body: what a panel anchored in the body covers on the first page includes the
// footer's own classification line, which is why Word shows that line on the
// second page and not on the first.
function stacked(
  floats: readonly PlacedFloat[],
  prefix: string,
  options: DrawingOptions,
): readonly Drawable[] {
  return floats
    .map((float, at) => ({ float, key: `${prefix}-${String(at)}` }))
    .sort((one, other) => one.float.anchor.relativeHeight - other.float.anchor.relativeHeight)
    .flatMap(({ float, key }) => fromFloat(float, key, options));
}

// The text a story flowed down the page, which nothing cuts off.
function flowedText(boxes: readonly ParagraphBox[], options: DrawingOptions): readonly Drawable[] {
  const uncut = boxes.filter((box) => box.clipTo === null);
  return hasText(uncut)
    ? [
        {
          kind: "text",
          key: "flowed-text",
          boxes: uncut,
          runs: [],
          underlines: underlinesIn(uncut, options),
          clipTo: null,
          turnDegrees: 0,
        },
      ]
    : [];
}

// A paragraph in a row told exactly how tall to be is drawn in a layer that is
// the row, so that what the row has no room for is not drawn at all. Each one
// keeps its own layer: cells of one row are cut off at different places along it.
function cutText(boxes: readonly ParagraphBox[], options: DrawingOptions): readonly Drawable[] {
  return boxes.flatMap((box, at) => {
    const clipTo = box.clipTo;
    if (clipTo === null || !hasText([box])) return [];
    return [
      {
        kind: "text" as const,
        key: `cut-text-${String(at)}`,
        boxes: [box],
        runs: [],
        underlines: underlinesIn([box], options),
        clipTo,
        turnDegrees: 0,
      },
    ];
  });
}

// What a paragraph draws behind itself reaches as far as its own lines do, and a
// paragraph with none is the room its mark stands in.
function paintedParagraphs(boxes: readonly ParagraphBox[]): readonly PaintedParagraph[] {
  return boxes.flatMap((box) =>
    box.paint === null
      ? []
      : [
          {
            paint: box.paint,
            topPt: box.lines[0]?.topPt ?? box.markTopPt,
            bottomPt: box.contentBottomPt,
          },
        ],
  );
}

/**
 * A page as this reads it: everything the layout states today, and the glyph runs
 * it will state.
 *
 * **The field is named here rather than on `LaidOutPage`** because the layout that
 * fills it is being built beside this, and a page carrying none draws none. The
 * day it moves onto the page proper, this line goes and nothing else changes.
 */
export type PageDrawing = LaidOutPage & {
  readonly glyphRuns?: readonly PlacedGlyphs[];
};

// A run holding no glyph at all draws nothing, so it is left out rather than
// handed to a backend to skip.
const glyphLayers = (page: PageDrawing): readonly Drawable[] =>
  (page.glyphRuns ?? [])
    .filter((run) => run.glyphs.length > 0)
    .map((run, at) => ({
      kind: "glyphs" as const,
      key: `glyphs-${String(at)}`,
      ...run,
      // The size lands on the grid as a run of text's does: a stretched delimiter is
      // text, and Word set every one of them at a whole number of units.
      sizePt: onTheDeviceGrid(run.sizePt),
      glyphs: run.glyphs.map((glyph) => ({
        ...glyph,
        baselinePt: onTheDeviceGrid(glyph.baselinePt),
      })),
    }));

/**
 * Every equation the story's lines hold, as the drawables that draw them.
 *
 * **An equation stands on a line like anything else on it**, so it is found where
 * the line put it rather than hung off the page: the segment carries the pieces
 * `setMath` placed about their own baseline, and the line says where that baseline
 * is. The boxes handed here are the drawn ones, so the origin is already on the
 * grid and each piece lands with its own line rather than a step off it.
 *
 * An equation that set nothing draws nothing, and is left out rather than handed to
 * a backend to skip.
 */
function equationLayers(boxes: readonly ParagraphBox[], prefix: string): readonly Drawable[] {
  const drawables: Drawable[] = [];

  boxes.forEach((box, boxAt) => {
    box.lines.forEach((placed, lineAt) => {
      placed.line.segments.forEach((segment, segmentAt) => {
        if (segment.kind !== "equation" || segment.pieces.length === 0) return;
        const primitives = mathPrimitivesOf(segment.pieces, {
          leftPt: placed.leftPt + segment.offsetPt,
          baselinePt: placed.baselinePt,
        });
        drawables.push(
          ...mathDrawables(
            { primitives, ascentPt: segment.ascentPt, descentPt: segment.descentPt },
            `${prefix}-${String(boxAt)}-${String(lineAt)}-${String(segmentAt)}`,
          ),
        );
      });
    });
  });
  return drawables;
}

export function drawablesOf(
  layout: LaidOutDocument,
  page: PageDrawing,
  options: DrawingOptions = {},
): readonly Drawable[] {
  const inlines = [...page.headerInlines, ...page.inlines, ...page.footerInlines].flatMap(
    (inline, at) => fromInline(inline, `inline-${String(at)}`, options),
  );

  const flowed = drawnBoxes([...page.header, ...page.body, ...page.footer], options);
  const text = [...flowedText(flowed, options), ...cutText(flowed, options)];
  const cells = [...page.headerCells, ...page.cells, ...page.footerCells].map(drawnCell);

  // **Where the text stands in the body's own stack is what `behindDoc` says**, and
  // the stack itself is still ordered by the height each float was given. Measured on
  // 2026-08-14 by a corpus document whose body anchors a grey band marked `behindDoc`
  // across the foot of its page: Word draws the footer's own text over that band, and
  // drawing every body float over the text buried it.
  //
  // **The cut is the highest float marked `behindDoc`, not the mark alone**: a
  // document already measured sends a panel to the back of the stack and draws a box
  // marked `behindDoc` over it, so a float standing under one that is behind the text
  // is behind the text as well, whatever it says of itself.
  const ordered = [...page.floats].sort(
    (one, other) => one.anchor.relativeHeight - other.anchor.relativeHeight,
  );
  const lastBehind = ordered.reduce(
    (found, float, at) => (float.anchor.behindDoc ? at : found),
    -1,
  );
  const behind = ordered.slice(0, lastBehind + 1);
  const inFront = ordered.slice(lastBehind + 1);

  return [
    ...stacked([...page.headerFloats, ...page.footerFloats], "story", options),
    ...stacked(behind, "behind", options),
    ...paintLayer(cells, flowed, "paint"),
    ...text,
    ...glyphLayers(page),
    // An equation is text, and stands where the story's text stands. Its own
    // pieces stack among themselves in the order the flattener set them in, a bar
    // over the half it crosses and not under it.
    ...equationLayers(flowed, "equation"),
    ...inlines,
    ...stacked(inFront, "float", options),
  ];
}
