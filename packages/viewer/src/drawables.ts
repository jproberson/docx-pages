import type {
  LaidOutDocument,
  LaidOutPage,
  ParagraphBox,
  PlacedContent,
  PlacedFloat,
  PlacedInline,
} from "@onepager/core";

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
  if (content.kind !== "text-box" || content.text === null || !hasText(content.text.boxes)) {
    return [];
  }
  const clipTo = {
    leftPt: float.leftPt,
    topPt: float.topPt,
    widthPt: float.widthPt,
    heightPt: float.heightPt,
  };
  return [{ kind: "text", key: `${key}-text`, boxes: content.text.boxes, clipTo }];
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

export function drawablesOf(layout: LaidOutDocument, page: LaidOutPage): readonly Drawable[] {
  const inlines = [...layout.headerInlines, ...page.inlines, ...layout.footerInlines].map(
    (inline, at) => fromInline(inline, `inline-${String(at)}`),
  );

  const flowed = [...layout.header, ...page.body, ...layout.footer];
  const text = [...flowedText(flowed), ...cutText(flowed)];

  return [
    ...stacked([...layout.headerFloats, ...layout.footerFloats], "story"),
    ...text,
    ...inlines,
    ...stacked(page.floats, "float"),
  ];
}
