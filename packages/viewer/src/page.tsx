import type { CSSProperties, ReactElement } from "react";

import {
  twipsToPoints,
  type CropInsets,
  type LaidOutDocument,
  type LaidOutPage,
  type ParagraphBox,
  type ParagraphMark,
  type ParagraphMarker,
  type PlacedLine,
  type PlacedPaint,
} from "@onepager/core";

import { drawablesOf, type Drawable } from "./drawables.js";
import type { ImageResolver } from "./images.js";

export type FrameStyle = "hidden" | "outlined";

export type OnePagerDocumentProps = {
  readonly layout: LaidOutDocument;
  readonly imageUrl: ImageResolver;
  readonly scale?: number;
  readonly frames?: FrameStyle;
  readonly fallbackFonts?: string;
  readonly className?: string;
};

export type OnePagerPageProps = OnePagerDocumentProps & {
  readonly page: LaidOutPage;
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

// A shape's own paint. Word centres an outline on the edge it runs along, which
// is where a stroked path sits, so the geometry is drawn at the object's own
// bounds and the layer is grown by the stroke's width to leave room for the half
// of it that falls outside. That room is also what lets a line shape draw at all:
// the ones here are stored with no height.
function painted(
  drawable: ObjectDrawable,
  paint: PlacedPaint,
  key: string,
  kind: string,
): ReactElement | null {
  const { widthPt, heightPt } = drawable;
  const { outline } = paint;
  if (paint.fillColor === null && outline === null) return null;

  const room = outline === null ? 0 : outline.widthPt;
  const layerWidth = widthPt + room * 2;
  const layerHeight = heightPt + room * 2;
  if (layerWidth <= 0 || layerHeight <= 0) return null;

  const stroke = { stroke: outline?.color, strokeWidth: outline?.widthPt };
  return (
    <svg
      key={key}
      style={{
        position: "absolute",
        left: pt(drawable.leftPt - room),
        top: pt(drawable.topPt - room),
        width: pt(layerWidth),
        height: pt(layerHeight),
      }}
      viewBox={`${String(-room)} ${String(-room)} ${String(layerWidth)} ${String(layerHeight)}`}
      data-kind={kind}
    >
      {paint.geometry === "line" ? (
        <line x1={0} y1={0} x2={widthPt} y2={heightPt} {...stroke} />
      ) : (
        <rect
          x={0}
          y={0}
          width={widthPt}
          height={heightPt}
          fill={paint.fillColor ?? "none"}
          {...stroke}
        />
      )}
    </svg>
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
// in a substitute, and starting each one where layout put it keeps a tab's gap
// and an inline picture from carrying the rest of the line along with them.
function lineText(placed: PlacedLine, key: string, fallback: string): ReactElement | null {
  const spans = placed.line.segments.flatMap((segment, at) =>
    segment.kind === "text"
      ? [
          <tspan
            key={at}
            x={placed.leftPt + segment.offsetPt}
            y={placed.baselinePt - segment.mark.raisePt}
            xmlSpace="preserve"
            fontFamily={familyOf(segment.mark, fallback)}
            fontSize={segment.mark.fontSizePt}
            fontWeight={segment.mark.bold ? "bold" : undefined}
            fontStyle={segment.mark.italic ? "italic" : undefined}
            textDecoration={segment.mark.underline ? "underline" : undefined}
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

// A list's number is drawn out of the text flow, at the position the level's
// hanging indent pulls the first line back to.
function markerText(marker: ParagraphMarker, key: string, fallback: string): ReactElement | null {
  if (marker.text === "") return null;
  return (
    <text
      key={key}
      x={marker.leftPt}
      y={marker.baselinePt}
      xmlSpace="preserve"
      fontFamily={familyOf(marker.mark, fallback)}
      fontSize={marker.mark.fontSizePt}
      fontWeight={marker.mark.bold ? "bold" : undefined}
      fontStyle={marker.mark.italic ? "italic" : undefined}
      textDecoration={marker.mark.underline ? "underline" : undefined}
      fill={marker.mark.color ?? undefined}
      textLength={marker.widthPt > 0 ? marker.widthPt : undefined}
      lengthAdjust="spacing"
    >
      {marker.text}
    </text>
  );
}

function textLayer(
  drawable: Extract<Drawable, { kind: "text" }>,
  widthPt: number,
  heightPt: number,
  fallback: string,
): ReactElement {
  // The header, the body and the footer each number their paragraphs from zero, so
  // a key is where the box sits in the layer rather than the index it carries.
  const lines = drawable.boxes.flatMap((paragraph: ParagraphBox, box: number) => {
    const key = String(box);
    const marker =
      paragraph.marker === null ? null : markerText(paragraph.marker, `${key}-number`, fallback);
    return [
      ...(marker === null ? [] : [marker]),
      ...paragraph.lines.flatMap((placed, at) => {
        const element = lineText(placed, `${key}-${String(at)}`, fallback);
        return element === null ? [] : [element];
      }),
    ];
  });

  // The page around it is sized in points, so the layer has to be as well: sized
  // in the browser's own pixels instead, every glyph is drawn three quarters of
  // the size layout measured it at and lands three quarters of the way to where
  // layout put it.
  return (
    <svg
      key={drawable.key}
      data-kind="text"
      style={{ position: "absolute", left: 0, top: 0, overflow: "visible" }}
      width={pt(widthPt)}
      height={pt(heightPt)}
      viewBox={`0 0 ${String(widthPt)} ${String(heightPt)}`}
      fill="currentColor"
    >
      {lines}
    </svg>
  );
}

const kept = (element: ReactElement | null): readonly ReactElement[] =>
  element === null ? [] : [element];

// A picture is drawn over whatever fills its frame and under its own outline, so
// the two halves of its paint go either side of the bitmap.
function renderPicture(
  drawable: ObjectDrawable,
  paint: PlacedPaint,
  image: ReactElement,
): readonly ReactElement[] {
  return [
    ...kept(painted(drawable, { ...paint, outline: null }, `${drawable.key}-fill`, "picture-fill")),
    image,
    ...kept(
      painted(drawable, { ...paint, fillColor: null }, `${drawable.key}-line`, "picture-outline"),
    ),
  ];
}

function renderObject(
  drawable: ObjectDrawable,
  imageUrl: ImageResolver,
  frames: FrameStyle,
): readonly ReactElement[] {
  const { content } = drawable;
  const shown = (kind: string): readonly ReactElement[] => kept(frame(drawable, kind, frames));
  const paint = (kind: string, own: PlacedPaint): readonly ReactElement[] => [
    ...kept(painted(drawable, own, `${drawable.key}-paint`, kind)),
    ...shown(kind),
  ];

  switch (content.kind) {
    case "picture": {
      const url = imageUrl(content.part);
      return url === undefined
        ? shown("unresolved-picture")
        : renderPicture(drawable, content.paint, croppedImage(drawable, url, content.crop));
    }
    case "missing-picture":
      return shown("missing-picture");
    case "text-box":
      return paint("text-box", content.paint);
    case "shape":
      return paint("shape", content.paint);
    case "unknown":
      return shown("unknown");
  }
}

export function OnePagerPage(props: OnePagerPageProps): ReactElement {
  const {
    layout,
    page,
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
      data-onepager-page={page.index}
      style={{
        position: "relative",
        width: pt(widthPt),
        height: pt(heightPt),
        overflow: "hidden",
        transform: scale === 1 ? undefined : `scale(${String(scale)})`,
        transformOrigin: "top left",
      }}
    >
      {drawablesOf(layout, page).flatMap((drawable) =>
        drawable.kind === "text"
          ? [textLayer(drawable, widthPt, heightPt, fallbackFonts)]
          : renderObject(drawable, imageUrl, frames),
      )}
    </div>
  );
}

// Every page the body broke onto, in order, each drawn in the page's own
// coordinates.
export function OnePagerDocument(props: OnePagerDocumentProps): ReactElement {
  return (
    <>
      {props.layout.pages.map((page) => (
        <OnePagerPage key={page.index} {...props} page={page} />
      ))}
    </>
  );
}
