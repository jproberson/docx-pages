import type { CSSProperties, ReactElement } from "react";

import {
  aliasedSymbolText,
  drawablesOf,
  paintOfCell,
  paintOfParagraph,
  ROUNDED_CORNER_FRACTION,
  twipsToPoints,
  type BorderStyle,
  type CropInsets,
  type Drawable,
  type DrawingFlip,
  type PathCommand,
  type Painted,
  type PaintedFill,
  type PaintedLine,
  type LaidOutDocument,
  type LaidOutPage,
  type MetafilePicture,
  type MetafileRect,
  type MetafileShape,
  type ParagraphBox,
  type ParagraphMark,
  type ParagraphMarker,
  type PlacedLine,
  type PlacedPaint,
} from "@docx-pages/core";

import type { ImageResolver } from "./images.js";

export type FrameStyle = "hidden" | "outlined";

export type DocumentProps = {
  readonly layout: LaidOutDocument;
  readonly imageUrl: ImageResolver;
  readonly scale?: number;
  readonly frames?: FrameStyle;
  readonly fallbackFonts?: string;
  // Symbol faces the layout stood in for, by lowercased name: runs written in
  // one are drawn as what their positions mean rather than as the stand-in's
  // own letters. `DocxDocument` fills this from its substitution report; a
  // caller who supplied the real face leaves it out and the run draws as
  // written.
  readonly aliasSymbolFaces?: ReadonlySet<string>;
  readonly className?: string;
};

export type PageProps = DocumentProps & {
  readonly page: LaidOutPage;
};

// The document names the face it was authored in; whatever the page can actually
// load falls in behind it.
const DEFAULT_FALLBACK_FONTS = "sans-serif";

const pt = (value: number): string => `${String(value)}pt`;

type ObjectDrawable = Extract<Drawable, { kind: "object" }>;

// An object turned after it was drawn is turned about the middle of the box it
// stands in, which is where a css transform turns an element by default.
const turn = (turnDegrees: number): CSSProperties =>
  turnDegrees === 0 ? {} : { transform: `rotate(${String(turnDegrees)}deg)` };

const box = (drawable: ObjectDrawable): CSSProperties => ({
  position: "absolute",
  left: pt(drawable.leftPt),
  top: pt(drawable.topPt),
  width: pt(drawable.widthPt),
  height: pt(drawable.heightPt),
  ...turn(drawable.turnDegrees),
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

// A metafile records the drawing rather than a picture of it, so its shapes are
// drawn straight into the frame the document gave it. The frame decides the scale
// on its own, whatever the recording's own proportions were, and a source rectangle
// is a narrower window onto the same coordinates rather than a larger picture
// behind a smaller frame.
function metafileImage(
  drawable: ObjectDrawable,
  picture: MetafilePicture,
  crop: CropInsets,
  fallback: string,
): ReactElement {
  const { widthUnits, heightUnits } = picture;
  const window = [
    crop.left * widthUnits,
    crop.top * heightUnits,
    widthUnits * (1 - crop.left - crop.right),
    heightUnits * (1 - crop.top - crop.bottom),
  ];

  const clips = clipPathsOf(picture.shapes, drawable.key);
  return (
    <svg
      key={drawable.key}
      style={box(drawable)}
      viewBox={window.map(String).join(" ")}
      preserveAspectRatio="none"
      data-kind="picture"
    >
      {clips.size === 0 ? null : (
        <defs>
          {[...clips].map(([rect, id]) => (
            <clipPath key={id} id={id}>
              <rect
                x={rect.leftUnits}
                y={rect.topUnits}
                width={rect.widthUnits}
                height={rect.heightUnits}
              />
            </clipPath>
          ))}
        </defs>
      )}
      {picture.shapes.map((shape, at) => {
        const drawn = metafileShape(shape, String(at), fallback);
        const id = shape.clipTo === null ? undefined : clips.get(shape.clipTo);
        return id === undefined ? (
          drawn
        ) : (
          <g key={at} clipPath={`url(#${id})`}>
            {drawn}
          </g>
        );
      })}
    </svg>
  );
}

// One clip path per rectangle the shapes are cut to, named against the object so
// that two metafiles on a page cannot claim the same name.
function clipPathsOf(
  shapes: readonly MetafileShape[],
  key: string,
): ReadonlyMap<MetafileRect, string> {
  const byRect = new Map<MetafileRect, string>();
  const byShape = new Map<string, string>();
  for (const shape of shapes) {
    if (shape.clipTo === null) continue;
    const { leftUnits, topUnits, widthUnits, heightUnits } = shape.clipTo;
    const shapeOf = [leftUnits, topUnits, widthUnits, heightUnits].map(String).join(" ");
    const id = byShape.get(shapeOf) ?? `${key}-clip-${String(byShape.size)}`;
    byShape.set(shapeOf, id);
    byRect.set(shape.clipTo, id);
  }
  return byRect;
}

// A metafile's pen hangs its line off the cell whose corner the line runs through,
// where a stroke is centred on the line it follows, so the line moves half a unit
// down and right to cover the same cells.
const PEN_OFFSET = 0.5;

function metafileShape(shape: MetafileShape, key: string, fallback: string): ReactElement {
  switch (shape.kind) {
    case "fill":
      return (
        <rect
          key={key}
          x={shape.rect.leftUnits}
          y={shape.rect.topUnits}
          width={shape.rect.widthUnits}
          height={shape.rect.heightUnits}
          fill={shape.color}
        />
      );
    case "stroke":
      return (
        <line
          key={key}
          x1={shape.fromXUnits + PEN_OFFSET}
          y1={shape.fromYUnits + PEN_OFFSET}
          x2={shape.toXUnits + PEN_OFFSET}
          y2={shape.toYUnits + PEN_OFFSET}
          stroke={shape.color}
          strokeWidth={shape.widthUnits}
        />
      );
    case "text":
      return (
        <text
          key={key}
          x={shape.xUnits.map(String).join(" ")}
          y={shape.baselineUnits}
          xmlSpace="preserve"
          fontFamily={`"${shape.face.name}", ${fallback}`}
          fontSize={shape.emUnits}
          fontWeight={shape.face.bold ? "bold" : undefined}
          fontStyle={shape.face.italic ? "italic" : undefined}
          fill={shape.color}
        >
          {shape.text}
        </text>
      );
  }
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
  // A path this cannot play is drawn as nothing at all rather than as the box it
  // fits in, which would be a filled rectangle the size of whatever it rules.
  if (paint.geometry === "custom" && paint.path === null) return null;

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
        // The room an outline takes is the same on every side, so this layer keeps
        // the shape's own middle and turns about it.
        ...turn(drawable.turnDegrees),
      }}
      viewBox={`${String(-room)} ${String(-room)} ${String(layerWidth)} ${String(layerHeight)}`}
      data-kind={kind}
    >
      {geometry(paint, widthPt, heightPt, drawable.flip, stroke)}
    </svg>
  );
}

type Stroke = {
  readonly stroke: string | undefined;
  readonly strokeWidth: number | undefined;
};

/**
 * The preset the shape names, drawn in the box the object stands in.
 *
 * **A flip decides which corners a line runs between and nothing else here.** A
 * connector is stored as a box with a line across it, so the two corners it joins
 * are the ones the flips choose; a shape with a symmetry either way is drawn the
 * same however it was turned.
 */
function geometry(
  paint: PlacedPaint,
  widthPt: number,
  heightPt: number,
  flip: DrawingFlip,
  stroke: Stroke,
): ReactElement {
  const fill = paint.fillColor ?? "none";

  switch (paint.geometry) {
    case "line": {
      const [x1, x2] = flip.horizontal ? [widthPt, 0] : [0, widthPt];
      const [y1, y2] = flip.vertical ? [heightPt, 0] : [0, heightPt];
      return <line x1={x1} y1={y1} x2={x2} y2={y2} {...stroke} />;
    }
    case "ellipse":
      return (
        <ellipse
          cx={widthPt / 2}
          cy={heightPt / 2}
          rx={widthPt / 2}
          ry={heightPt / 2}
          fill={fill}
          {...stroke}
        />
      );
    case "rounded-rectangle": {
      const radius = Math.min(widthPt, heightPt) * ROUNDED_CORNER_FRACTION;
      return (
        <rect
          x={0}
          y={0}
          width={widthPt}
          height={heightPt}
          rx={radius}
          ry={radius}
          fill={fill}
          {...stroke}
        />
      );
    }
    case "triangle": {
      const apex = flip.vertical ? heightPt : 0;
      const base = flip.vertical ? 0 : heightPt;
      return (
        <polygon
          points={`${String(widthPt / 2)},${String(apex)} ${String(widthPt)},${String(base)} 0,${String(base)}`}
          fill={fill}
          {...stroke}
        />
      );
    }
    case "custom":
      return (
        <path d={pathData(paint.path ?? [], widthPt, heightPt, flip)} fill={fill} {...stroke} />
      );
    case "rectangle":
      return <rect x={0} y={0} width={widthPt} height={heightPt} fill={fill} {...stroke} />;
  }
}

// The outline a file drew point by point, in shares of its own box, written out as
// svg. **The shares are core's**, so that this and the pdf writer cannot disagree
// about where a point lands; all either does with them is put them in its own
// coordinates, and svg counts y down the page as the layout does.
function pathData(
  path: readonly PathCommand[],
  widthPt: number,
  heightPt: number,
  flip: DrawingFlip,
): string {
  const xOf = (share: number): string =>
    String(Math.round((flip.horizontal ? 1 - share : share) * widthPt * 1000) / 1000);
  const yOf = (share: number): string =>
    String(Math.round((flip.vertical ? 1 - share : share) * heightPt * 1000) / 1000);

  return path
    .map((command) => {
      if (command.kind === "move") return `M ${xOf(command.to.x)} ${yOf(command.to.y)}`;
      if (command.kind === "line") return `L ${xOf(command.to.x)} ${yOf(command.to.y)}`;
      if (command.kind === "close") return "Z";
      return (
        `C ${xOf(command.first.x)} ${yOf(command.first.y)}` +
        ` ${xOf(command.second.x)} ${yOf(command.second.y)}` +
        ` ${xOf(command.to.x)} ${yOf(command.to.y)}`
      );
    })
    .join(" ");
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

// A run in a symbol face that was stood in for holds positions in that face's
// page, and the stand-in would draw them as its own letters. Drawn as what the
// positions mean instead, which is how the layout measured them.
function shownText(
  mark: ParagraphMark,
  text: string,
  aliasFaces: ReadonlySet<string> | null,
): string {
  if (aliasFaces === null || mark.font.kind !== "named") return text;
  if (!aliasFaces.has(mark.font.name.trim().toLowerCase())) return text;
  return aliasedSymbolText(mark.font.name, text) ?? text;
}

// The line was measured with the authored face's own widths. Holding each run to
// that width keeps the break points Word chose even when the page draws the text
// in a substitute, and starting each one where layout put it keeps a tab's gap
// and an inline picture from carrying the rest of the line along with them.
function lineText(
  placed: PlacedLine,
  key: string,
  fallback: string,
  aliasFaces: ReadonlySet<string> | null,
): ReactElement | null {
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
            {shownText(segment.mark, segment.text, aliasFaces)}
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
function markerText(
  marker: ParagraphMarker,
  key: string,
  fallback: string,
  aliasFaces: ReadonlySet<string> | null,
): ReactElement | null {
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
      {shownText(marker.mark, marker.text, aliasFaces)}
    </text>
  );
}

function textLayer(
  drawable: Extract<Drawable, { kind: "text" }>,
  widthPt: number,
  heightPt: number,
  fallback: string,
  aliasFaces: ReadonlySet<string> | null,
): ReactElement {
  // The header, the body and the footer each number their paragraphs from zero, so
  // a key is where the box sits in the layer rather than the index it carries.
  const lines = drawable.boxes.flatMap((paragraph: ParagraphBox, box: number) => {
    const key = String(box);
    const marker =
      paragraph.marker === null
        ? null
        : markerText(paragraph.marker, `${key}-number`, fallback, aliasFaces);
    return [
      ...(marker === null ? [] : [marker]),
      ...paragraph.lines.flatMap((placed, at) => {
        const element = lineText(placed, `${key}-${String(at)}`, fallback, aliasFaces);
        return element === null ? [] : [element];
      }),
    ];
  });

  // The page around it is sized in points, so the layer has to be as well: sized
  // in the browser's own pixels instead, every glyph is drawn three quarters of
  // the size layout measured it at and lands three quarters of the way to where
  // layout put it.
  //
  // Text cut to a shape's frame is drawn in a layer that is the frame, which an
  // svg clips its contents to. The coordinates inside are the page's either way,
  // so a line is written at the point layout put it on the page.
  const { clipTo } = drawable;
  const window = clipTo ?? { leftPt: 0, topPt: 0, widthPt, heightPt };
  return (
    <svg
      key={drawable.key}
      data-kind="text"
      style={{
        position: "absolute",
        left: pt(window.leftPt),
        top: pt(window.topPt),
        ...(clipTo === null ? { overflow: "visible" } : {}),
        ...turn(drawable.turnDegrees),
      }}
      width={pt(window.widthPt)}
      height={pt(window.heightPt)}
      viewBox={`${String(window.leftPt)} ${String(window.topPt)} ${String(window.widthPt)} ${String(window.heightPt)}`}
      fill="currentColor"
    >
      {lines}
    </svg>
  );
}

// How Word draws each pattern, measured at a width of a point and a half: a
// dashed line runs four widths on and four off, where a dotted one runs one and
// one. A double line is two bands, which the geometry has already made of it.
const DASHES: Readonly<Record<BorderStyle, readonly number[] | null>> = {
  single: null,
  double: null,
  dashed: [4, 4],
  dotted: [1, 1],
};

function paintedLine(line: PaintedLine, key: string): ReactElement {
  const dashes = DASHES[line.style];
  return (
    <line
      key={key}
      x1={line.vertical ? line.atPt : line.fromPt}
      x2={line.vertical ? line.atPt : line.toPt}
      y1={line.vertical ? line.fromPt : line.atPt}
      y2={line.vertical ? line.toPt : line.atPt}
      stroke={line.color ?? "currentColor"}
      strokeWidth={line.widthPt}
      strokeDasharray={
        dashes === null ? undefined : dashes.map((each) => each * line.widthPt).join(" ")
      }
    />
  );
}

const paintedFill = (fill: PaintedFill, key: string): ReactElement => (
  <rect
    key={key}
    x={fill.leftPt}
    y={fill.topPt}
    width={fill.widthPt}
    height={fill.heightPt}
    fill={fill.color}
  />
);

const drawn = (painted: Painted, key: string): readonly ReactElement[] => [
  ...painted.fills.map((fill, at) => paintedFill(fill, `${key}-fill-${String(at)}`)),
  ...painted.lines.map((line, at) => paintedLine(line, `${key}-line-${String(at)}`)),
];

// Everything drawn behind a story's text: the cells of its tables first, then what
// each paragraph asks for, which Word draws over the cell holding it.
function paintLayer(
  drawable: Extract<Drawable, { kind: "paint" }>,
  widthPt: number,
  heightPt: number,
): ReactElement {
  return (
    <svg
      key={drawable.key}
      data-kind="paint"
      style={{ position: "absolute", left: 0, top: 0 }}
      width={pt(widthPt)}
      height={pt(heightPt)}
      viewBox={`0 0 ${String(widthPt)} ${String(heightPt)}`}
    >
      {drawable.cells.flatMap((cell, at) => drawn(paintOfCell(cell), `cell-${String(at)}`))}
      {drawable.paragraphs.flatMap((each, at) =>
        drawn(paintOfParagraph(each.paint, each.topPt, each.bottomPt), `paragraph-${String(at)}`),
      )}
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
  fallback: string,
): readonly ReactElement[] {
  const { content } = drawable;
  const shown = (kind: string): readonly ReactElement[] => kept(frame(drawable, kind, frames));
  const paint = (kind: string, own: PlacedPaint): readonly ReactElement[] => [
    ...kept(painted(drawable, own, `${drawable.key}-paint`, kind)),
    ...shown(kind),
  ];

  switch (content.kind) {
    case "picture": {
      const image = imageUrl(content.part);
      if (image === undefined) return shown("unresolved-picture");
      return renderPicture(
        drawable,
        content.paint,
        image.kind === "bitmap"
          ? croppedImage(drawable, image.url, content.crop)
          : metafileImage(drawable, image.picture, content.crop, fallback),
      );
    }
    case "missing-picture":
      return shown("missing-picture");
    case "text-box":
      return paint("text-box", content.paint);
    case "shape":
      return paint("shape", content.paint);
    // A group is flattened into one item per shape inside it before a renderer
    // sees anything, which is where the rule about what covers what belongs.
    case "group":
      return [];
    case "unknown":
      return shown("unknown");
  }
}

export function Page(props: PageProps): ReactElement {
  const {
    layout,
    page,
    imageUrl,
    scale = 1,
    frames = "hidden",
    fallbackFonts = DEFAULT_FALLBACK_FONTS,
    aliasSymbolFaces = null,
    className,
  } = props;
  // Each page is the size the section whose text opened it asked for, which a
  // document of more than one section can change partway down.
  const widthPt = twipsToPoints(page.geometry.widthTwips);
  const heightPt = twipsToPoints(page.geometry.heightTwips);

  return (
    <div
      className={className}
      data-docx-page={page.index}
      style={{
        position: "relative",
        width: pt(widthPt),
        height: pt(heightPt),
        overflow: "hidden",
        transform: scale === 1 ? undefined : `scale(${String(scale)})`,
        transformOrigin: "top left",
      }}
    >
      {drawablesOf(layout, page).flatMap((drawable) => {
        if (drawable.kind === "paint") return [paintLayer(drawable, widthPt, heightPt)];
        if (drawable.kind === "text") {
          return [textLayer(drawable, widthPt, heightPt, fallbackFonts, aliasSymbolFaces)];
        }
        return renderObject(drawable, imageUrl, frames, fallbackFonts);
      })}
    </div>
  );
}

// Every page the body broke onto, in order, each drawn in the page's own
// coordinates.
export function Document(props: DocumentProps): ReactElement {
  return (
    <>
      {props.layout.pages.map((page) => (
        <Page key={page.index} {...props} page={page} />
      ))}
    </>
  );
}
