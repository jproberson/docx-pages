import type { Paragraph } from "./blocks.js";
import {
  readDrawingContent,
  readDrawingFlip,
  type DrawingContent,
  type DrawingFlip,
} from "./drawing.js";
import { paragraphOwnDrawings } from "./paragraphs.js";
import { attribute, firstNamed, type XmlElement } from "./xml.js";

export const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

export type AnchorOrigin = "page" | "margin" | "column" | "paragraph" | "line" | "character";

export type AnchorPosition =
  | { readonly kind: "offset"; readonly from: AnchorOrigin; readonly offsetEmu: number }
  | { readonly kind: "align"; readonly from: AnchorOrigin; readonly align: string };

export type WrapMode = "none" | "square" | "tight" | "through" | "topAndBottom";

// How far text is kept off each edge of a wrapping object.
export type WrapDistances = {
  readonly topEmu: number;
  readonly rightEmu: number;
  readonly bottomEmu: number;
  readonly leftEmu: number;
};

// How much of an object's frame text is kept off, as fractions of it. A tight or
// through wrap carries a polygon instead of taking the whole frame, and Word
// keeps text off the rectangle around that polygon: measured by pulling one
// inside its frame and watching the text beside it move in by as much.
export type WrapArea = {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
};

export const WHOLE_FRAME: WrapArea = { left: 0, top: 0, right: 1, bottom: 1 };

export type FloatingAnchor = {
  readonly paragraphIndex: number;
  readonly name: string;
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly horizontal: AnchorPosition;
  readonly vertical: AnchorPosition;
  readonly content: DrawingContent;
  readonly wrap: WrapMode;
  readonly area: WrapArea;
  readonly distances: WrapDistances;
  readonly behindDoc: boolean;
  readonly relativeHeight: number;
};

const ORIGINS: readonly AnchorOrigin[] = [
  "page",
  "margin",
  "column",
  "paragraph",
  "line",
  "character",
];

const WRAPS: readonly (readonly [string, WrapMode])[] = [
  ["wrapNone", "none"],
  ["wrapSquare", "square"],
  ["wrapTight", "tight"],
  ["wrapThrough", "through"],
  ["wrapTopAndBottom", "topAndBottom"],
];

function toOrigin(raw: string | undefined, fallback: AnchorOrigin): AnchorOrigin {
  return ORIGINS.find((origin) => origin === raw) ?? fallback;
}

function readPosition(
  anchor: XmlElement,
  name: "positionH" | "positionV",
  fallback: AnchorOrigin,
): AnchorPosition {
  const node = firstNamed(anchor, WP_NS, name);
  const from = toOrigin(node === null ? undefined : attribute(node, "", "relativeFrom"), fallback);
  if (node === null) return { kind: "offset", from, offsetEmu: 0 };

  const align = firstNamed(node, WP_NS, "align");
  if (align !== null && align.text !== "") {
    return { kind: "align", from, align: align.text.trim() };
  }

  const offset = firstNamed(node, WP_NS, "posOffset");
  const value = offset === null ? Number.NaN : Number(offset.text.trim());
  return { kind: "offset", from, offsetEmu: Number.isFinite(value) ? value : 0 };
}

// The polygon is written in 21600ths of the frame it belongs to, whichever way
// round that frame ended up.
const POLYGON_UNITS = 21600;

const turned = (low: number, high: number, flipped: boolean): readonly [number, number] =>
  flipped ? [1 - high, 1 - low] : [low, high];

function readWrapArea(wrap: XmlElement, flip: DrawingFlip): WrapArea {
  const polygon = firstNamed(wrap, WP_NS, "wrapPolygon");
  const corners =
    polygon === null
      ? []
      : polygon.children.flatMap((point) => {
          const x = numberAttribute(point, "x", Number.NaN);
          const y = numberAttribute(point, "y", Number.NaN);
          return Number.isFinite(x) && Number.isFinite(y) ? [[x, y] as const] : [];
        });
  if (corners.length === 0) return WHOLE_FRAME;

  const xs = corners.map(([x]) => x / POLYGON_UNITS);
  const ys = corners.map(([, y]) => y / POLYGON_UNITS);
  const [left, right] = turned(Math.min(...xs), Math.max(...xs), flip.horizontal);
  const [top, bottom] = turned(Math.min(...ys), Math.max(...ys), flip.vertical);
  return { left, top, right, bottom };
}

type Wrapping = { readonly mode: WrapMode; readonly area: WrapArea };

function readWrapping(anchor: XmlElement, flip: DrawingFlip): Wrapping {
  for (const [name, mode] of WRAPS) {
    const wrap = firstNamed(anchor, WP_NS, name);
    if (wrap !== null) return { mode, area: readWrapArea(wrap, flip) };
  }
  return { mode: "none", area: WHOLE_FRAME };
}

const numberAttribute = (element: XmlElement, name: string, fallback: number): number => {
  const raw = attribute(element, "", name);
  const value = raw === undefined ? Number.NaN : Number(raw);
  return Number.isFinite(value) ? value : fallback;
};

export function readAnchors(paragraph: Paragraph): readonly FloatingAnchor[] {
  return paragraphOwnDrawings(paragraph, WP_NS, "anchor").map((anchor) => {
    const extent = firstNamed(anchor, WP_NS, "extent");
    const docPr = firstNamed(anchor, WP_NS, "docPr");
    const wrapping = readWrapping(anchor, readDrawingFlip(anchor));
    return {
      paragraphIndex: paragraph.index,
      name: docPr === null ? "" : (attribute(docPr, "", "name") ?? ""),
      widthEmu: extent === null ? 0 : numberAttribute(extent, "cx", 0),
      heightEmu: extent === null ? 0 : numberAttribute(extent, "cy", 0),
      content: readDrawingContent(anchor),
      horizontal: readPosition(anchor, "positionH", "column"),
      vertical: readPosition(anchor, "positionV", "paragraph"),
      wrap: wrapping.mode,
      area: wrapping.area,
      distances: {
        topEmu: numberAttribute(anchor, "distT", 0),
        rightEmu: numberAttribute(anchor, "distR", 0),
        bottomEmu: numberAttribute(anchor, "distB", 0),
        leftEmu: numberAttribute(anchor, "distL", 0),
      },
      behindDoc: attribute(anchor, "", "behindDoc") === "1",
      relativeHeight: numberAttribute(anchor, "relativeHeight", 0),
    };
  });
}
