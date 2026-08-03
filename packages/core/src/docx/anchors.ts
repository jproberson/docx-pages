import type { Paragraph } from "./blocks.js";
import { readDrawingContent, type DrawingContent } from "./drawing.js";
import { paragraphOwnDrawings } from "./paragraphs.js";
import { attribute, firstNamed, type XmlElement } from "./xml.js";

export const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

export type AnchorOrigin = "page" | "margin" | "column" | "paragraph" | "line" | "character";

export type AnchorPosition =
  | { readonly kind: "offset"; readonly from: AnchorOrigin; readonly offsetEmu: number }
  | { readonly kind: "align"; readonly from: AnchorOrigin; readonly align: string };

export type WrapMode = "none" | "square" | "tight" | "through" | "topAndBottom";

export type FloatingAnchor = {
  readonly paragraphIndex: number;
  readonly name: string;
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly horizontal: AnchorPosition;
  readonly vertical: AnchorPosition;
  readonly content: DrawingContent;
  readonly wrap: WrapMode;
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

function readWrap(anchor: XmlElement): WrapMode {
  for (const [name, mode] of WRAPS) {
    if (firstNamed(anchor, WP_NS, name) !== null) return mode;
  }
  return "none";
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
    return {
      paragraphIndex: paragraph.index,
      name: docPr === null ? "" : (attribute(docPr, "", "name") ?? ""),
      widthEmu: extent === null ? 0 : numberAttribute(extent, "cx", 0),
      heightEmu: extent === null ? 0 : numberAttribute(extent, "cy", 0),
      content: readDrawingContent(anchor),
      horizontal: readPosition(anchor, "positionH", "column"),
      vertical: readPosition(anchor, "positionV", "paragraph"),
      wrap: readWrap(anchor),
      behindDoc: attribute(anchor, "", "behindDoc") === "1",
      relativeHeight: numberAttribute(anchor, "relativeHeight", 0),
    };
  });
}
