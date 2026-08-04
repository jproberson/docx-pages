import type { CSSProperties, ReactElement } from "react";

import { twipsToPoints, type CropInsets, type LaidOutDocument } from "@onepager/core";

import { drawablesOf, type Drawable } from "./drawables.js";
import type { ImageResolver } from "./images.js";

export type FrameStyle = "hidden" | "outlined";

export type OnePagerPageProps = {
  readonly layout: LaidOutDocument;
  readonly imageUrl: ImageResolver;
  readonly scale?: number;
  readonly frames?: FrameStyle;
  readonly className?: string;
};

const pt = (value: number): string => `${String(value)}pt`;

const box = (drawable: Drawable): CSSProperties => ({
  position: "absolute",
  left: pt(drawable.leftPt),
  top: pt(drawable.topPt),
  width: pt(drawable.widthPt),
  height: pt(drawable.heightPt),
});

// srcRect hides a fraction of each edge, so the whole bitmap is larger than the
// placed rectangle by exactly that much and is shifted up and left behind it.
function croppedImage(drawable: Drawable, url: string, crop: CropInsets): ReactElement {
  const width = drawable.widthPt / Math.max(1 - crop.left - crop.right, Number.EPSILON);
  const height = drawable.heightPt / Math.max(1 - crop.top - crop.bottom, Number.EPSILON);

  return (
    <div key={drawable.key} style={{ ...box(drawable), overflow: "hidden" }} data-kind="picture">
      <img
        src={url}
        alt={drawable.name}
        style={{
          position: "absolute",
          left: pt(-crop.left * width),
          top: pt(-crop.top * height),
          width: pt(width),
          height: pt(height),
        }}
      />
    </div>
  );
}

function frame(drawable: Drawable, kind: string, frames: FrameStyle): ReactElement | null {
  if (frames === "hidden") return null;
  return (
    <svg
      key={drawable.key}
      style={box(drawable)}
      viewBox={`0 0 ${String(drawable.widthPt)} ${String(drawable.heightPt)}`}
      data-kind={kind}
    >
      <rect
        x={0}
        y={0}
        width={drawable.widthPt}
        height={drawable.heightPt}
        fill="none"
        stroke="currentColor"
        strokeDasharray="3 3"
        strokeWidth={0.5}
      />
    </svg>
  );
}

function render(
  drawable: Drawable,
  imageUrl: ImageResolver,
  frames: FrameStyle,
): ReactElement | null {
  const { content } = drawable;
  switch (content.kind) {
    case "picture": {
      const url = imageUrl(content.part);
      return url === undefined
        ? frame(drawable, "unresolved-picture", frames)
        : croppedImage(drawable, url, content.crop);
    }
    case "missing-picture":
      return frame(drawable, "missing-picture", frames);
    case "text-box":
      return frame(drawable, "text-box", frames);
    case "shape":
      return frame(drawable, "shape", frames);
    case "unknown":
      return frame(drawable, "unknown", frames);
  }
}

export function OnePagerPage(props: OnePagerPageProps): ReactElement {
  const { layout, imageUrl, scale = 1, frames = "hidden", className } = props;
  const widthPt = twipsToPoints(layout.page.widthTwips);
  const heightPt = twipsToPoints(layout.page.heightTwips);

  return (
    <div
      className={className}
      data-onepager-page=""
      style={{
        position: "relative",
        width: pt(widthPt),
        height: pt(heightPt),
        overflow: "hidden",
        transform: scale === 1 ? undefined : `scale(${String(scale)})`,
        transformOrigin: "top left",
      }}
    >
      {drawablesOf(layout).map((drawable) => render(drawable, imageUrl, frames))}
    </div>
  );
}
