import type { Paragraph } from "./blocks.js";
import {
  readDrawingContent,
  readDrawingFlip,
  readDrawingTurn,
  type DrawingContent,
  type DrawingFlip,
} from "./drawing.js";
import { paragraphOwnDrawings } from "./paragraphs.js";
import { W_NS } from "./section.js";
import { legacyAnchoredDrawingsIn } from "./vml.js";
import { attribute, firstNamed, statedNumber, type XmlElement } from "./xml.js";

export const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

export type AnchorOrigin = "page" | "margin" | "column" | "paragraph" | "line" | "character";

export type AnchorPosition =
  | { readonly kind: "offset"; readonly from: AnchorOrigin; readonly offsetEmu: number }
  | { readonly kind: "align"; readonly from: AnchorOrigin; readonly align: string };

export type WrapMode = "none" | "square" | "tight" | "through" | "topAndBottom";

// Which side of a wrapping object text is allowed on. `largest` leaves it to
// whichever side of the object has the most room in the column.
export type WrapSide = "bothSides" | "left" | "right" | "largest";

// How far text is kept off each edge of a wrapping object.
export type WrapDistances = {
  readonly topEmu: number;
  readonly rightEmu: number;
  readonly bottomEmu: number;
  readonly leftEmu: number;
};

// A corner of a wrap polygon, as fractions of the frame it belongs to.
export type WrapCorner = { readonly x: number; readonly y: number };

// How much of an object's frame text is kept off, as fractions of it. A tight or
// through wrap carries a polygon instead of taking the whole frame, and Word
// keeps text off the rectangle around that polygon: measured by pulling one
// inside its frame and watching the text beside it move in by as much. The
// corners are kept as well, since a line meeting only part of a polygon that is
// not a rectangle is held off that part alone.
export type WrapArea = {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly corners: readonly WrapCorner[];
};

export const WHOLE_FRAME: WrapArea = { left: 0, top: 0, right: 1, bottom: 1, corners: [] };

export type FloatingAnchor = {
  readonly paragraphIndex: number;
  readonly name: string;
  readonly widthEmu: number;
  readonly heightEmu: number;
  // A width the file states as a share of the text frame rather than as a length,
  // which only the old drawing form does and only the section knows the frame of.
  // See `statedWidthOf`: where one stands, it is drawn and the length is not.
  readonly frameWidthShare?: number;
  // How far round the object was turned after it was drawn, which is paint alone
  // here: what an object turned out of its own box does to the text wrapping
  // around it has not been asked of Word.
  readonly turnDegrees: number;
  // Which way round the object was flipped after it was drawn. Read already for
  // the wrap polygon, which is written for the object the right way round; this
  // carries it on to whatever draws the object itself.
  readonly flip: DrawingFlip;
  readonly horizontal: AnchorPosition;
  readonly vertical: AnchorPosition;
  readonly content: DrawingContent;
  readonly wrap: WrapMode;
  readonly side: WrapSide;
  readonly area: WrapArea;
  readonly distances: WrapDistances;
  // How far past its own extent the object is drawn, which Word writes down for it:
  // the half of an outline that falls outside the edge, and whatever a shadow or a
  // glow adds. **Text is kept off this as well as off the extent**, and the anchor's
  // distances are held off the whole of the two. An anchor built by hand rather than
  // read out of a document states none, and overhangs nothing.
  readonly effect?: WrapDistances;
  readonly behindDoc: boolean;
  readonly relativeHeight: number;
  /**
   * Whether a cell the object is anchored in is the frame it is placed against,
   * rather than the page.
   *
   * **Word measures every horizontal origin from the cell's own left**, page and
   * column alike. Read on 2026-08-18 off Word's own pdf over three documents:
   * `c8ca0c3c8292` and `2c1289b95c31` draw four page-relative pictures 340.80,
   * 340.76, 456.16 and 456.18 from where this project put them, and those two
   * offsets are the two cells' own lefts, whose grid widths differ by exactly the
   * 115.40 between them; `7eaa70746b70` draws three column-relative groups at its
   * two cells' lefts, all three within a quarter of a point of a constant that is
   * the table's own left rather than the rule.
   *
   * Every one of the 92 anchors the corpus states inside a cell states this on.
   */
  readonly inTheCell: boolean;
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

const over = (at: number, flipped: boolean): number => (flipped ? 1 - at : at);

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

  const turnedCorners = corners.map(([x, y]) => ({
    x: over(x / POLYGON_UNITS, flip.horizontal),
    y: over(y / POLYGON_UNITS, flip.vertical),
  }));
  const xs = turnedCorners.map((corner) => corner.x);
  const ys = turnedCorners.map((corner) => corner.y);
  return {
    left: Math.min(...xs),
    right: Math.max(...xs),
    top: Math.min(...ys),
    bottom: Math.max(...ys),
    corners: turnedCorners,
  };
}

type Wrapping = { readonly mode: WrapMode; readonly side: WrapSide; readonly area: WrapArea };

const SIDES: readonly WrapSide[] = ["bothSides", "left", "right", "largest"];

// An object wrapped top and bottom, or not at all, leaves nothing beside it to
// choose between, and one that states no side takes both.
function readSide(wrap: XmlElement): WrapSide {
  const stated = attribute(wrap, "", "wrapText");
  return SIDES.find((side) => side === stated) ?? "bothSides";
}

function readWrapping(anchor: XmlElement, flip: DrawingFlip): Wrapping {
  for (const [name, mode] of WRAPS) {
    const wrap = firstNamed(anchor, WP_NS, name);
    if (wrap !== null) return { mode, side: readSide(wrap), area: readWrapArea(wrap, flip) };
  }
  return { mode: "none", side: "bothSides", area: WHOLE_FRAME };
}

// Word's own default is on, so an attribute left out states the cell rather than
// the page. Every corpus anchor inside a cell writes it out all the same.
const statedOn = (raw: string | undefined): boolean => raw !== "0" && raw !== "false";

const numberAttribute = (element: XmlElement, name: string, fallback: number): number => {
  const value = statedNumber(attribute(element, "", name));
  return Number.isFinite(value) ? value : fallback;
};

export const NO_EFFECT: WrapDistances = { topEmu: 0, rightEmu: 0, bottomEmu: 0, leftEmu: 0 };

// An effect extent is written with a letter a side, and a drawing that overhangs
// nothing writes it out as four noughts rather than leaving it out.
const effectOf = (element: XmlElement | null): WrapDistances =>
  element === null
    ? NO_EFFECT
    : {
        topEmu: numberAttribute(element, "t", 0),
        rightEmu: numberAttribute(element, "r", 0),
        bottomEmu: numberAttribute(element, "b", 0),
        leftEmu: numberAttribute(element, "l", 0),
      };

// The containers the old form writes a floating drawing inside, which the walk asks
// about by name exactly as it asks for a `wp:anchor`.
const LEGACY_CONTAINERS = [
  { namespace: W_NS, name: "pict" },
  { namespace: W_NS, name: "object" },
];

/**
 * The drawings the old form hangs out of the flow, read into the same anchor as a
 * `wp:anchor` so that one set of placement rules answers for both.
 *
 * **A text box written this way reached no frame, no line and no page until this
 * was built**: 20 corpus documents hold 70 of them and every word in one was drawn
 * nowhere at all, while the report named the loss and nothing answered it. What
 * this reads and what it still passes over is `readLegacyDrawing`'s answer, and the
 * fidelity report asks the very same function, so a box drawn here stops being
 * named there in the same breath.
 */
function legacyAnchors(paragraph: Paragraph): readonly FloatingAnchor[] {
  return paragraphOwnDrawings(paragraph, LEGACY_CONTAINERS).flatMap((pict) =>
    legacyAnchoredDrawingsIn(pict).map((drawing): FloatingAnchor => ({
      paragraphIndex: paragraph.index,
      ...drawing,
      // No turn is read off a VML shape yet, and no positioned text box in the
      // corpus states one. The flip is the drawing's own: a box states none, and a
      // line states which of its box's two diagonals it runs along.
      turnDegrees: 0,
      // See `legacyAnchoredDrawingsIn`: what the old form says about wrapping is
      // not carried across, so nothing here moves a line of the flow.
      // What the old form says about a cell (`o:allowincell`) is not read either,
      // and no corpus document positions a VML shape inside one.
      inTheCell: false,
      wrap: "none",
      side: "bothSides",
      area: WHOLE_FRAME,
      effect: NO_EFFECT,
    })),
  );
}

export function readAnchors(paragraph: Paragraph): readonly FloatingAnchor[] {
  return [...drawingMlAnchors(paragraph), ...legacyAnchors(paragraph)];
}

function drawingMlAnchors(paragraph: Paragraph): readonly FloatingAnchor[] {
  return paragraphOwnDrawings(paragraph, [{ namespace: WP_NS, name: "anchor" }]).map((anchor) => {
    const extent = firstNamed(anchor, WP_NS, "extent");
    const docPr = firstNamed(anchor, WP_NS, "docPr");
    const flip = readDrawingFlip(anchor);
    const wrapping = readWrapping(anchor, flip);
    return {
      paragraphIndex: paragraph.index,
      name: docPr === null ? "" : (attribute(docPr, "", "name") ?? ""),
      widthEmu: extent === null ? 0 : numberAttribute(extent, "cx", 0),
      heightEmu: extent === null ? 0 : numberAttribute(extent, "cy", 0),
      turnDegrees: readDrawingTurn(anchor),
      flip,
      content: readDrawingContent(anchor),
      horizontal: readPosition(anchor, "positionH", "column"),
      vertical: readPosition(anchor, "positionV", "paragraph"),
      wrap: wrapping.mode,
      side: wrapping.side,
      area: wrapping.area,
      distances: {
        topEmu: numberAttribute(anchor, "distT", 0),
        rightEmu: numberAttribute(anchor, "distR", 0),
        bottomEmu: numberAttribute(anchor, "distB", 0),
        leftEmu: numberAttribute(anchor, "distL", 0),
      },
      effect: effectOf(firstNamed(anchor, WP_NS, "effectExtent")),
      behindDoc: attribute(anchor, "", "behindDoc") === "1",
      relativeHeight: numberAttribute(anchor, "relativeHeight", 0),
      inTheCell: statedOn(attribute(anchor, "", "layoutInCell")),
    };
  });
}
