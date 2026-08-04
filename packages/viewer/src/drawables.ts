import type {
  LaidOutDocument,
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
function textOf(float: PlacedFloat, key: string): readonly Drawable[] {
  const { content } = float;
  if (content.kind !== "text-box" || content.text === null || !hasText(content.text.boxes)) {
    return [];
  }
  return [{ kind: "text", key: `${key}-text`, boxes: content.text.boxes }];
}

// Word stacks floats by relativeHeight, with behindDoc ones under the text and the
// rest over it. Inline drawings live in the text itself, so they sit between.
export function drawablesOf(layout: LaidOutDocument): readonly Drawable[] {
  const floats = [...layout.headerFloats, ...layout.bodyFloats, ...layout.footerFloats].map(
    (float, at) => {
      const key = `float-${String(at)}`;
      return { float, drawables: [fromFloat(float, key), ...textOf(float, key)] };
    },
  );

  const byHeight = (one: (typeof floats)[number], other: (typeof floats)[number]): number =>
    one.float.anchor.relativeHeight - other.float.anchor.relativeHeight;

  const behind = floats.filter((each) => each.float.anchor.behindDoc).sort(byHeight);
  const above = floats.filter((each) => !each.float.anchor.behindDoc).sort(byHeight);
  const inlines = [...layout.headerInlines, ...layout.bodyInlines, ...layout.footerInlines].map(
    (inline, at) => fromInline(inline, `inline-${String(at)}`),
  );

  const flowed = [...layout.header, ...layout.body, ...layout.footer];
  const text: readonly Drawable[] = hasText(flowed)
    ? [{ kind: "text", key: "flowed-text", boxes: flowed }]
    : [];

  return [
    ...behind.flatMap((each) => each.drawables),
    ...text,
    ...inlines,
    ...above.flatMap((each) => each.drawables),
  ];
}
