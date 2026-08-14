import type { BorderStyle } from "../docx/borders.js";
import type { DrawingFlip } from "../docx/drawing.js";
import type { ParagraphMark } from "../docx/styles.js";
import type { LaidOutDocument, LaidOutPage } from "./document.js";
import type { PlacedContent, PlacedFloat } from "./floats.js";
import type { FaceRequest } from "./font-metrics.js";
import type { PlacedInline } from "./inlines.js";
import { paintOfCell, paintOfParagraph, type PaintedFill, type PaintedLine } from "./painting.js";
import type { ParagraphBox, ParagraphPaint, PlacedCell, PlacedLine } from "./stack.js";
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
              mark: { ...segment.mark, color: drawnColor(segment.mark.color) },
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
            mark: { ...box.marker.mark, color: drawnColor(box.marker.mark.color) },
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
// A point of an outline, in the face's own units, with x from the glyph's origin
// and y up from the baseline as the face states it.
export type GlyphPoint = readonly [number, number];

// What the outline does next. A TrueType face draws its curves with one control
// point and a PostScript one with two, and neither is worth turning into the other
// here: every backend that can draw a curve at all can draw both.
export type GlyphStep =
  | { readonly kind: "line"; readonly to: GlyphPoint }
  | { readonly kind: "quadratic"; readonly control: GlyphPoint; readonly to: GlyphPoint }
  | {
      readonly kind: "cubic";
      readonly first: GlyphPoint;
      readonly second: GlyphPoint;
      readonly to: GlyphPoint;
    };

// One closed contour of a glyph: where it starts and what it does from there. A
// letter with a counter, an `o` above all, is two of them, and what is inside what
// is settled by the winding rather than by their order.
export type GlyphContour = {
  readonly from: GlyphPoint;
  readonly steps: readonly GlyphStep[];
};

/**
 * The shape a glyph draws, as the face itself states it.
 *
 * **This is what a backend that cannot name a glyph draws instead**: a browser
 * addresses a face by character and by nothing else, so a stretched parenthesis
 * reaches it as the outline or not at all. A backend that embeds the face should
 * still name the glyph rather than draw this: the embedded one is hinted, it is
 * selectable, and it is the same shape by construction.
 *
 * In font units, so a caller scales by the size the run is set at. Nothing that
 * moves a glyph moves these: they are measured from the glyph's own origin, which
 * is what `leftPt` and `baselinePt` place.
 */
export type GlyphOutline = {
  readonly unitsPerEm: number;
  readonly contours: readonly GlyphContour[];
};

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
    }
  | {
      readonly kind: "text";
      readonly key: string;
      readonly boxes: readonly ParagraphBox[];
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
            underlines: underlinesIn(boxes, options),
            clipTo,
            turnDegrees,
          },
        ]
      : []),
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
      glyphs: run.glyphs.map((glyph) => ({
        ...glyph,
        baselinePt: onTheDeviceGrid(glyph.baselinePt),
      })),
    }));

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
    ...inlines,
    ...stacked(inFront, "float", options),
  ];
}
