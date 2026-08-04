import type { CSSProperties, ReactElement } from "react";

import {
  twipsToPoints,
  type CropInsets,
  type LaidOutDocument,
  type ParagraphBox,
  type ParagraphMark,
  type PlacedLine,
} from "@onepager/core";

import { drawablesOf, type Drawable } from "./drawables.js";
import type { ImageResolver } from "./images.js";

export type FrameStyle = "hidden" | "outlined";

export type OnePagerPageProps = {
  readonly layout: LaidOutDocument;
  readonly imageUrl: ImageResolver;
  readonly scale?: number;
  readonly frames?: FrameStyle;
  readonly fallbackFonts?: string;
  readonly className?: string;
};

// The document names the face it was authored in; whatever the page can actually
// load falls in behind it.
const DEFAULT_FALLBACK_FONTS = "sans-serif";

const pt = (value: number): string => `${String(value)}pt`;

type ObjectDrawable = Extract<Drawable, { kind: "object" }>;

const box = (drawable: ObjectDrawable): CSSProperties => ({
  position: "absolute",
  left: pt(drawable.leftPt),
  top: pt(drawable.topPt),
  width: pt(drawable.widthPt),
  height: pt(drawable.heightPt),
});

// srcRect hides a fraction of each edge, so the whole bitmap is larger than the
// placed rectangle by exactly that much and is shifted up and left behind it.
function croppedImage(drawable: ObjectDrawable, url: string, crop: CropInsets): ReactElement {
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

function frame(drawable: ObjectDrawable, kind: string, frames: FrameStyle): ReactElement | null {
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

const familyOf = (mark: ParagraphMark, fallback: string): string =>
  mark.font.kind === "named" ? `"${mark.font.name}", ${fallback}` : fallback;

// The line was measured with the authored face's own widths. Holding each run to
// that width keeps the break points Word chose even when the page draws the text
// in a substitute.
function lineText(placed: PlacedLine, key: string, fallback: string): ReactElement | null {
  const spans = placed.line.segments.flatMap((segment, at) =>
    segment.kind === "text"
      ? [
          <tspan
            key={at}
            xmlSpace="preserve"
            fontFamily={familyOf(segment.mark, fallback)}
            fontSize={segment.mark.fontSizePt}
            fontWeight={segment.mark.bold ? "bold" : undefined}
            fontStyle={segment.mark.italic ? "italic" : undefined}
            fill={segment.mark.color ?? undefined}
            textLength={segment.widthPt > 0 ? segment.widthPt : undefined}
            lengthAdjust="spacing"
          >
            {segment.text}
          </tspan>,
        ]
      : [],
  );

  if (spans.length === 0) return null;
  return (
    <text key={key} x={placed.leftPt} y={placed.baselinePt}>
      {spans}
    </text>
  );
}

function textLayer(
  drawable: Extract<Drawable, { kind: "text" }>,
  widthPt: number,
  heightPt: number,
  fallback: string,
): ReactElement {
  const lines = drawable.boxes.flatMap((paragraph: ParagraphBox) =>
    paragraph.lines.flatMap((placed, at) => {
      const element = lineText(placed, `${String(paragraph.index)}-${String(at)}`, fallback);
      return element === null ? [] : [element];
    }),
  );

  return (
    <svg
      key={drawable.key}
      data-kind="text"
      style={{ position: "absolute", left: 0, top: 0, overflow: "visible" }}
      width={widthPt}
      height={heightPt}
      viewBox={`0 0 ${String(widthPt)} ${String(heightPt)}`}
      fill="currentColor"
    >
      {lines}
    </svg>
  );
}

function renderObject(
  drawable: ObjectDrawable,
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
  const {
    layout,
    imageUrl,
    scale = 1,
    frames = "hidden",
    fallbackFonts = DEFAULT_FALLBACK_FONTS,
    className,
  } = props;
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
      {drawablesOf(layout).map((drawable) =>
        drawable.kind === "text"
          ? textLayer(drawable, widthPt, heightPt, fallbackFonts)
          : renderObject(drawable, imageUrl, frames),
      )}
    </div>
  );
}
