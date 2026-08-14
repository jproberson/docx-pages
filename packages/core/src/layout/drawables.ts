import type { DrawingFlip } from "../docx/drawing.js";
import type { LaidOutDocument, LaidOutPage } from "./document.js";
import type { PlacedContent, PlacedFloat } from "./floats.js";
import type { FaceRequest } from "./font-metrics.js";
import type { PlacedInline } from "./inlines.js";
import type { ParagraphBox, ParagraphPaint, PlacedCell, PlacedLine } from "./stack.js";
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

// A line as it is drawn. Both edges land on the grid and the height is what lies
// between them, so what is painted behind a line keeps the line's own foot instead
// of drifting a step off it.
function drawnLine(line: PlacedLine): PlacedLine {
  const topPt = onTheDeviceGrid(line.topPt);
  return {
    ...line,
    topPt,
    baselinePt: onTheDeviceGrid(line.baselinePt),
    fittingHeightPt: onTheDeviceGrid(line.topPt + line.fittingHeightPt) - topPt,
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
const drawnBoxes = (boxes: readonly ParagraphBox[]): readonly ParagraphBox[] =>
  boxes.map((box) => ({
    ...box,
    markTopPt: onTheDeviceGrid(box.markTopPt),
    contentBottomPt: onTheDeviceGrid(box.contentBottomPt),
    marker:
      box.marker === null
        ? null
        : { ...box.marker, baselinePt: onTheDeviceGrid(box.marker.baselinePt) },
    lines: box.lines.map(drawnLine),
  }));

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
      readonly cells: readonly PlacedCell[];
      readonly paragraphs: readonly PaintedParagraph[];
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
function objectsOf(standing: Standing, key: string): readonly Drawable[] {
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
      );
    });
  }

  return [{ kind: "object", key, ...standing }, ...textOf(standing, key)];
}

const fromFloat = (float: PlacedFloat, key: string): readonly Drawable[] =>
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
  );

const fromInline = (inline: PlacedInline, key: string): readonly Drawable[] =>
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
function textOf(standing: Standing, key: string): readonly Drawable[] {
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
  const boxes = square ? drawnBoxes(content.text.boxes) : content.text.boxes;
  const cells = square ? content.text.cells.map(drawnCell) : content.text.cells;
  const { inlines } = content.text;
  const painted = paintLayer(cells, boxes, `${key}-paint`);
  const turnDegrees = standing.turnDegrees;
  return [
    ...painted,
    ...(hasText(boxes)
      ? [{ kind: "text" as const, key: `${key}-text`, boxes, clipTo, turnDegrees }]
      : []),
    // **A drawing standing in the box's own text is drawn where that text put it.**
    // Fifteen corpus documents hold 76 of them and not one was drawn until
    // 2026-08-14: the text around each came out as usual, so no page said anything
    // was missing.
    ...inlines.flatMap((inline, at) => fromInline(inline, `${key}-inline-${String(at)}`)),
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
  return [{ kind: "paint", key, cells, paragraphs, highlights }];
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
function stacked(floats: readonly PlacedFloat[], prefix: string): readonly Drawable[] {
  return floats
    .map((float, at) => ({ float, key: `${prefix}-${String(at)}` }))
    .sort((one, other) => one.float.anchor.relativeHeight - other.float.anchor.relativeHeight)
    .flatMap(({ float, key }) => fromFloat(float, key));
}

// The text a story flowed down the page, which nothing cuts off.
function flowedText(boxes: readonly ParagraphBox[]): readonly Drawable[] {
  const uncut = boxes.filter((box) => box.clipTo === null);
  return hasText(uncut)
    ? [{ kind: "text", key: "flowed-text", boxes: uncut, clipTo: null, turnDegrees: 0 }]
    : [];
}

// A paragraph in a row told exactly how tall to be is drawn in a layer that is
// the row, so that what the row has no room for is not drawn at all. Each one
// keeps its own layer: cells of one row are cut off at different places along it.
function cutText(boxes: readonly ParagraphBox[]): readonly Drawable[] {
  return boxes.flatMap((box, at) => {
    const clipTo = box.clipTo;
    if (clipTo === null || !hasText([box])) return [];
    return [
      {
        kind: "text" as const,
        key: `cut-text-${String(at)}`,
        boxes: [box],
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

export function drawablesOf(layout: LaidOutDocument, page: PageDrawing): readonly Drawable[] {
  const inlines = [...page.headerInlines, ...page.inlines, ...page.footerInlines].flatMap(
    (inline, at) => fromInline(inline, `inline-${String(at)}`),
  );

  const flowed = drawnBoxes([...page.header, ...page.body, ...page.footer]);
  const text = [...flowedText(flowed), ...cutText(flowed)];
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
    ...stacked([...page.headerFloats, ...page.footerFloats], "story"),
    ...stacked(behind, "behind"),
    ...paintLayer(cells, flowed, "paint"),
    ...text,
    ...glyphLayers(page),
    ...inlines,
    ...stacked(inFront, "float"),
  ];
}
