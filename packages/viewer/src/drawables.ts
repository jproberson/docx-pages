import type {
  LaidOutDocument,
  LaidOutPage,
  ParagraphBox,
  ParagraphPaint,
  PlacedCell,
  PlacedContent,
  PlacedFloat,
  PlacedInline,
} from "@docx-pages/core";

// A paragraph's own fill and border, with the room its lines took: how far up and
// down they reach is the paragraph's, how far across is the text area's.
export type PaintedParagraph = {
  readonly paint: ParagraphPaint;
  readonly topPt: number;
  readonly bottomPt: number;
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
    }
  | {
      readonly kind: "text";
      readonly key: string;
      readonly boxes: readonly ParagraphBox[];
      // The rectangle the text is cut to, which a shape's own text has and the
      // text a story flowed down the page does not.
      readonly clipTo: Rect | null;
    }
  // Everything drawn behind the text of a story: the cells of its tables and the
  // fills and borders its paragraphs ask for. One layer holds them all, since they
  // are drawn in the page's own coordinates and nothing stands between them.
  | {
      readonly kind: "paint";
      readonly key: string;
      readonly cells: readonly PlacedCell[];
      readonly paragraphs: readonly PaintedParagraph[];
    };

export type Rect = {
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

const fromFloat = (float: PlacedFloat, key: string): Drawable => ({
  kind: "object",
  key,
  name: float.anchor.name,
  content: float.content,
  leftPt: float.leftPt,
  topPt: float.topPt,
  widthPt: float.widthPt,
  heightPt: float.heightPt,
});

const fromInline = (inline: PlacedInline, key: string): Drawable => ({
  kind: "object",
  key,
  name: inline.drawing.name,
  content: inline.content,
  leftPt: inline.leftPt,
  topPt: inline.topPt,
  widthPt: inline.widthPt,
  heightPt: inline.heightPt,
});

const hasText = (boxes: readonly ParagraphBox[]): boolean =>
  boxes.some((box) => box.lines.length > 0);

// A text box draws its own text straight after its frame, so the two keep the one
// place in the stack that Word gave the shape.
//
// Word cuts that text off at the frame. A box told to fit itself to its text was
// grown to hold all of it and loses nothing, but one that was not keeps the size
// it was stored at and shows only as much as fits: the rest is not moved
// anywhere, it is simply not drawn.
function textOf(float: PlacedFloat, key: string): readonly Drawable[] {
  const { content } = float;
  if (content.kind !== "text-box" || content.text === null) return [];

  const clipTo = {
    leftPt: float.leftPt,
    topPt: float.topPt,
    widthPt: float.widthPt,
    heightPt: float.heightPt,
  };
  const { boxes, cells } = content.text;
  const painted = paintLayer(cells, boxes, `${key}-paint`);
  return [
    ...painted,
    ...(hasText(boxes) ? [{ kind: "text" as const, key: `${key}-text`, boxes, clipTo }] : []),
  ];
}

// Everything drawn behind a story's text, where there is anything at all.
function paintLayer(
  cells: readonly PlacedCell[],
  boxes: readonly ParagraphBox[],
  key: string,
): readonly Drawable[] {
  const paragraphs = paintedParagraphs(boxes);
  if (cells.length + paragraphs.length === 0) return [];
  return [{ kind: "paint", key, cells, paragraphs }];
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
    .flatMap(({ float, key }) => [fromFloat(float, key), ...textOf(float, key)]);
}

// The text a story flowed down the page, which nothing cuts off.
function flowedText(boxes: readonly ParagraphBox[]): readonly Drawable[] {
  const uncut = boxes.filter((box) => box.clipTo === null);
  return hasText(uncut) ? [{ kind: "text", key: "flowed-text", boxes: uncut, clipTo: null }] : [];
}

// A paragraph in a row told exactly how tall to be is drawn in a layer that is
// the row, so that what the row has no room for is not drawn at all. Each one
// keeps its own layer: cells of one row are cut off at different places along it.
function cutText(boxes: readonly ParagraphBox[]): readonly Drawable[] {
  return boxes.flatMap((box, at) => {
    const clipTo = box.clipTo;
    if (clipTo === null || !hasText([box])) return [];
    return [{ kind: "text" as const, key: `cut-text-${String(at)}`, boxes: [box], clipTo }];
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

export function drawablesOf(layout: LaidOutDocument, page: LaidOutPage): readonly Drawable[] {
  const inlines = [...layout.headerInlines, ...page.inlines, ...layout.footerInlines].map(
    (inline, at) => fromInline(inline, `inline-${String(at)}`),
  );

  const flowed = [...layout.header, ...page.body, ...layout.footer];
  const text = [...flowedText(flowed), ...cutText(flowed)];
  const cells = [...layout.headerCells, ...page.cells, ...layout.footerCells];

  return [
    ...stacked([...layout.headerFloats, ...layout.footerFloats], "story"),
    ...paintLayer(cells, flowed, "paint"),
    ...text,
    ...inlines,
    ...stacked(page.floats, "float"),
  ];
}
