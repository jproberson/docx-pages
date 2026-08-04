import type { LaidOutDocument, PlacedContent, PlacedFloat, PlacedInline } from "@onepager/core";

export type Drawable = {
  readonly key: string;
  readonly name: string;
  readonly content: PlacedContent;
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

const fromFloat = (float: PlacedFloat, key: string): Drawable => ({
  key,
  name: float.anchor.name,
  content: float.content,
  leftPt: float.leftPt,
  topPt: float.topPt,
  widthPt: float.widthPt,
  heightPt: float.heightPt,
});

const fromInline = (inline: PlacedInline, key: string): Drawable => ({
  key,
  name: inline.drawing.name,
  content: inline.content,
  leftPt: inline.leftPt,
  topPt: inline.topPt,
  widthPt: inline.widthPt,
  heightPt: inline.heightPt,
});

// Word stacks floats by relativeHeight, with behindDoc ones under the text and the
// rest over it. Inline drawings live in the text itself, so they sit between.
export function drawablesOf(layout: LaidOutDocument): readonly Drawable[] {
  const floats = [...layout.headerFloats, ...layout.bodyFloats].map(
    (float, at) => [float, fromFloat(float, `float-${String(at)}`)] as const,
  );
  const byHeight = (
    one: readonly [PlacedFloat, Drawable],
    other: readonly [PlacedFloat, Drawable],
  ): number => one[0].anchor.relativeHeight - other[0].anchor.relativeHeight;

  const behind = floats.filter(([float]) => float.anchor.behindDoc).sort(byHeight);
  const above = floats.filter(([float]) => !float.anchor.behindDoc).sort(byHeight);
  const inlines = [...layout.headerInlines, ...layout.bodyInlines].map((inline, at) =>
    fromInline(inline, `inline-${String(at)}`),
  );

  return [...behind.map(([, drawable]) => drawable), ...inlines, ...above.map(([, one]) => one)];
}
