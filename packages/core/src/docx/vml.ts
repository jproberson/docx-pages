import type { AnchorOrigin, AnchorPosition, WrapDistances } from "./anchors.js";
import { blocksIn, isDetachedContent } from "./blocks.js";
import {
  DEFAULT_TEXT_INSETS,
  type CropInsets,
  type DrawingContent,
  type GroupChild,
  type DrawingFlip,
  type ShapeGeometry,
  type ShapePaint,
  type TextBoxAnchor,
  type TextBoxBody,
  type TextBoxInsets,
} from "./drawing.js";
import { R_NS } from "./relationships.js";
import { W_NS } from "./section.js";
import type { ColorReference } from "./theme.js";
import { attribute, type XmlElement } from "./xml.js";

// The drawing form Word wrote before DrawingML and still writes for some of what
// it draws: a `v:shape` whose `v:imagedata` names the picture.
//
// **Two containers hold one, and they are the same picture.** A `w:pict` is a
// drawing; a `w:object` is something embedded, a spreadsheet or an equation, and
// what Word draws for it is a picture of it written exactly this way, with
// `o:ole` on the shape and an `o:OLEObject` beside it naming what it came from.
// The embedding is nothing this can open and the picture is all Word draws, so
// both are read the same and neither is treated as a gap.
//
// **Only the picture standing in the run's own line is read here.** A VML shape
// carrying `position:absolute` is out of flow like a `wp:anchor`, and one holding
// a `v:textbox` rather than a picture is a text box; neither is drawn yet, and
// `undrawnLegacyDrawings` is what the fidelity report names them by. The
// `mc:Fallback` twin of a DrawingML shape never reaches this at all, since the
// blocks drop it before anything looks inside.
//
// Measured over the 718: every inline VML picture in the corpus states its width
// and height in points, so a size stated any other way is left at nothing rather
// than guessed at. **The crop is another matter**: no inline picture in the corpus
// crops, and every picture bullet declared in one does, which is the whole of what
// makes a 334 by 103 logo a small green ball.

export const V_NS = "urn:schemas-microsoft-com:vml";

export type VmlPicture = {
  readonly widthEmu: number;
  readonly heightEmu: number;
  // What the crop lets through, as fractions of the source hidden on each edge. The
  // size above is the window the crop leaves, not the whole picture: measured on
  // 2026-08-12, a shape saying 36 by 24pt and hiding four fifths across and half
  // down was drawn by Word 180 by 48pt with 36 by 24 of it showing.
  readonly crop: CropInsets;
  readonly relationshipId: string;
};

const EMU_PER_POINT = 12700;

// A VML shape's `style` is css, so it is read by declaration rather than by
// position, and a declaration this does not know is passed over.
function declarationsOf(shape: XmlElement): ReadonlyMap<string, string> {
  const stated = new Map<string, string>();
  const style = attribute(shape, "", "style");
  if (style === undefined) return stated;

  for (const each of style.split(";")) {
    const at = each.indexOf(":");
    if (at < 0) continue;
    stated.set(each.slice(0, at).trim().toLowerCase(), each.slice(at + 1).trim());
  }
  return stated;
}

const declaration = (shape: XmlElement, name: string): string | null =>
  declarationsOf(shape).get(name) ?? null;

const pointsOf = (shape: XmlElement, name: string): number => {
  const raw = declaration(shape, name);
  if (raw === null || !raw.toLowerCase().endsWith("pt")) return 0;
  const value = Number(raw.slice(0, -2));
  return Number.isFinite(value) ? value : 0;
};

const isPositioned = (shape: XmlElement): boolean =>
  declaration(shape, "position")?.toLowerCase() === "absolute";

// A VML crop is written as a fraction in sixteenths of a thousandth, `48837f`, or
// as a percentage, or as a plain fraction. Word writes the first.
const FRACTION_UNITS = 65536;

function cropEdge(imagedata: XmlElement, name: string): number {
  const raw = attribute(imagedata, "", name);
  if (raw === undefined) return 0;
  const scale = raw.endsWith("f") ? FRACTION_UNITS : raw.endsWith("%") ? 100 : 1;
  const value = Number(scale === 1 ? raw : raw.slice(0, -1));
  return Number.isFinite(value) ? value / scale : 0;
}

const cropOf = (imagedata: XmlElement): CropInsets => ({
  left: cropEdge(imagedata, "cropleft"),
  top: cropEdge(imagedata, "croptop"),
  right: cropEdge(imagedata, "cropright"),
  bottom: cropEdge(imagedata, "cropbottom"),
});

// The shapes VML draws a picture with. A `v:group` is a drawing of its own and is
// not one of them: reading its first picture would draw that one at the group's
// size and the rest nowhere at all.
const PICTURE_SHAPES = new Set(["shape", "rect", "roundrect", "oval"]);

function imageOf(node: XmlElement): XmlElement | null {
  for (const child of node.children) {
    if (isDetachedContent(child)) continue;
    if (child.namespace === V_NS && child.name === "imagedata") {
      const id = attribute(child, R_NS, "id");
      return id === undefined || id === "" ? null : child;
    }
    const found = imageOf(child);
    if (found !== null) return found;
  }
  return null;
}

// The elements a legacy picture is written inside, which the readers ask about by
// name rather than looking for a VML shape anywhere: a shape under anything else
// belongs to that thing.
export const holdsALegacyPicture = (namespace: string | null, name: string): boolean =>
  namespace === W_NS && (name === "pict" || name === "object");

// The picture a shape puts on the line, or null where it puts none there: one of
// the shapes that draw no picture at all, one out of the flow, or one naming no
// picture.
function pictureInLine(shape: XmlElement): XmlElement | null {
  if (shape.namespace !== V_NS || !PICTURE_SHAPES.has(shape.name)) return null;
  if (isPositioned(shape)) return null;
  return imageOf(shape);
}

/**
 * The picture a `w:pict` or a `w:object` puts on the line, or null where it puts
 * none there: an empty one, a shape that is not a picture, or one positioned out
 * of the flow.
 */
export function inlinePictureOf(pict: XmlElement): VmlPicture | null {
  for (const child of pict.children) {
    if (isDetachedContent(child)) continue;

    const imagedata = pictureInLine(child);
    if (imagedata === null) continue;

    return {
      widthEmu: pointsOf(child, "width") * EMU_PER_POINT,
      heightEmu: pointsOf(child, "height") * EMU_PER_POINT,
      crop: cropOf(imagedata),
      relationshipId: attribute(imagedata, R_NS, "id") ?? "",
    };
  }
  return null;
}

// A drawing written in the old form that nothing here draws. Whether it holds text
// is the whole of what is asked about it, since text drawn nowhere is a different
// fault from a line drawn nowhere.
export type UndrawnLegacyDrawing = {
  readonly holdsText: boolean;
};

/**
 * A drawing in the old form that stands out of the flow, read into the same anchor
 * a `wp:anchor` produces.
 *
 * **The point of answering in the other dialect's words** is that where a floating
 * object lands, what it is drawn over and how the text inside it is laid out are all
 * measured already. A second set of rules for the old form would be a second set to
 * get wrong; this way the reader is a translation and the layout is untouched.
 */
export type LegacyAnchoredDrawing = {
  readonly name: string;
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly flip: DrawingFlip;
  // See `statedWidthOf`: a width the file states as a share of the text frame rather
  // than as a length, which the layout resolves and the reading only carries.
  readonly frameWidthShare?: number;
  readonly horizontal: AnchorPosition;
  readonly vertical: AnchorPosition;
  readonly content: DrawingContent;
  readonly behindDoc: boolean;
  readonly relativeHeight: number;
  readonly distances: WrapDistances;
};

// What one drawing in a container comes to: what the layout is given of it, and
// what is left over for the report to name. A drawing is one or the other, and a
// group can be both, holding a text box this reads beside a picture it does not.
type LegacyReading = {
  readonly drawing: LegacyAnchoredDrawing | null;
  readonly undrawn: UndrawnLegacyDrawing | null;
};

// The VML elements that are drawings. Everything else a shape holds describes it:
// `v:shapetype` is the definition it names in its `type`, and `v:stroke`, `v:path`,
// `v:fill`, `v:textbox` and `v:imagedata` are its own properties.
const DRAWINGS = new Set([
  "shape",
  "rect",
  "roundrect",
  "oval",
  "line",
  "polyline",
  "curve",
  "arc",
  "group",
  "image",
]);

const drawingsIn = (node: XmlElement): readonly XmlElement[] =>
  node.children.filter(
    (child) => !isDetachedContent(child) && child.namespace === V_NS && DRAWINGS.has(child.name),
  );

// Whether a shape, or anything grouped inside it, holds a text box. What the text
// box says is passed over here as it is everywhere else: what is asked is whether
// the shape has text at all.
function holdsATextBox(shape: XmlElement): boolean {
  if (shape.namespace === V_NS && shape.name === "textbox") return true;
  return shape.children.some((child) => !isDetachedContent(child) && holdsATextBox(child));
}

// A length in a VML style is css and carries its own unit. Measured over the 718:
// of the 70 positioned text boxes in the corpus, 62 state their width in points and
// six in inches, and the bare numbers are all nought, which is the one bare number
// css allows.
const POINTS_IN: ReadonlyMap<string, number> = new Map([
  ["pt", 1],
  ["in", 72],
  ["pc", 12],
  ["cm", 72 / 2.54],
  ["mm", 7.2 / 2.54],
  ["px", 0.75],
]);

const LENGTH = /^(-?(?:\d+\.?\d*|\.\d+))([a-z]*)$/;

function lengthPt(raw: string | undefined): number | null {
  if (raw === undefined) return null;
  const stated = LENGTH.exec(raw.trim().toLowerCase());
  if (stated === null) return null;

  const value = Number(stated[1]);
  if (!Number.isFinite(value)) return null;

  const unit = stated[2] ?? "";
  if (unit === "") return value === 0 ? 0 : null;

  const points = POINTS_IN.get(unit);
  return points === undefined ? null : value * points;
}

const emuOf = (points: number): number => Math.round(points * EMU_PER_POINT);

type PositionNames = {
  readonly offset: string;
  readonly keyword: string;
  readonly relative: string;
};

const HORIZONTAL: PositionNames = {
  offset: "margin-left",
  keyword: "mso-position-horizontal",
  relative: "mso-position-horizontal-relative",
};

const VERTICAL: PositionNames = {
  offset: "margin-top",
  keyword: "mso-position-vertical",
  relative: "mso-position-vertical-relative",
};

// What each dialect calls the thing an offset is measured from. VML's `text` is the
// column across and the paragraph down, which is what a `wp:anchor` naming neither
// falls back to as well.
const HORIZONTAL_ORIGINS: ReadonlyMap<string, AnchorOrigin> = new Map([
  ["page", "page"],
  ["margin", "margin"],
  ["text", "column"],
  ["char", "character"],
]);

const VERTICAL_ORIGINS: ReadonlyMap<string, AnchorOrigin> = new Map([
  ["page", "page"],
  ["margin", "margin"],
  ["text", "paragraph"],
  ["paragraph", "paragraph"],
  ["line", "line"],
]);

/**
 * Where a shape stands on one axis, or null where the file states it in a way this
 * cannot read.
 *
 * **`left` and `top` are not the position.** Word writes `left:0;text-align:left`
 * beside a `margin-left` of 433.85pt on the same shape, and 19 of the corpus's 70
 * boxes state the pair; reading the first would stack them at the edge of the page.
 *
 * **An alignment beats the offset on its own axis.** The nine that state
 * `mso-position-horizontal:right` state a `margin-left` of 823.5pt, 1153pt or
 * 658.8pt beside it, every one of them off the paper.
 */
function positionOn(
  style: ReadonlyMap<string, string>,
  names: PositionNames,
  origins: ReadonlyMap<string, AnchorOrigin>,
  fallback: AnchorOrigin,
): AnchorPosition | null {
  const stated = style.get(names.relative);
  const from = stated === undefined ? fallback : origins.get(stated.toLowerCase());
  if (from === undefined) return null;

  const keyword = style.get(names.keyword)?.toLowerCase();
  if (keyword !== undefined && keyword !== "absolute") {
    return { kind: "align", from, align: keyword };
  }

  const offsetPt = lengthPt(style.get(names.offset));
  return offsetPt === null ? null : { kind: "offset", from, offsetEmu: emuOf(offsetPt) };
}

/**
 * How wide a shape is where its style states that twice, as a length and as a share
 * of something else.
 *
 * **The share wins, and it is in tenths of a percent.** Measured against Word on
 * 2026-08-19: a box stating `width:150pt` beside
 * `mso-width-percent:400;mso-width-relative:margin`, standing in a 540pt text frame,
 * broke its own eighteen words at 439.7, 428.5 and 411.2 from a left of 236.4. The
 * longest of those lines is 203.3pt and the word after it did not fit, which is a box
 * 216pt wide and not the 150 the same style states. 216 is 40.0% of the frame, so the
 * 400 is tenths of a percent rather than a percentage, and the `margin` it is a share
 * of is the text frame rather than the page.
 *
 * **Only `margin` is measured**, and every one of the nine corpus shapes stating a
 * share states it, so a share of anything else is refused rather than guessed at.
 *
 * **The height's own share is not read.** `mso-height-percent:200` stands beside the
 * width's on all nine, and what Word makes of it has never been asked; it costs
 * nothing to leave, since all nine fit their shape to their text and so draw no
 * stated height either.
 */
type StatedWidth =
  | { readonly kind: "in-points" }
  | { readonly kind: "share-of-the-frame"; readonly share: number }
  | { readonly kind: "unreadable" };

const TENTHS_OF_A_PERCENT = 1000;

function statedWidthOf(style: ReadonlyMap<string, string>): StatedWidth {
  const stated = style.get("mso-width-percent");
  if (stated === undefined) return { kind: "in-points" };

  const tenths = Number(stated);
  if (!Number.isFinite(tenths)) return { kind: "unreadable" };
  // Word writes the declaration out as a nought for a shape stating no share at all,
  // which is what nearly every shape in the corpus does.
  if (tenths === 0) return { kind: "in-points" };

  return style.get("mso-width-relative")?.toLowerCase() === "margin"
    ? { kind: "share-of-the-frame", share: tenths / TENTHS_OF_A_PERCENT }
    : { kind: "unreadable" };
}

// Word's own distances round a floating object, which VML states in points and a
// `wp:anchor` in EMU: an eighth of an inch at the sides and nothing above or below.
const DEFAULT_WRAP_DISTANCES: WrapDistances = {
  leftEmu: 114300,
  rightEmu: 114300,
  topEmu: 0,
  bottomEmu: 0,
};

function distancesOf(style: ReadonlyMap<string, string>): WrapDistances {
  const edge = (name: string, fallback: number): number => {
    const points = lengthPt(style.get(`mso-wrap-distance-${name}`));
    return points === null ? fallback : emuOf(points);
  };
  return {
    leftEmu: edge("left", DEFAULT_WRAP_DISTANCES.leftEmu),
    rightEmu: edge("right", DEFAULT_WRAP_DISTANCES.rightEmu),
    topEmu: edge("top", DEFAULT_WRAP_DISTANCES.topEmu),
    bottomEmu: edge("bottom", DEFAULT_WRAP_DISTANCES.bottomEmu),
  };
}

// Where the drawing stands and how big it is, with everything about it that does not
// depend on what it holds.
type LegacyFrame = {
  readonly widthEmu: number;
  readonly heightEmu: number;
  // Which way round the shape is drawn in the box it was given. A box states none
  // and a line states which of its box's two diagonals it runs along.
  readonly flip: DrawingFlip;
  // See `statedWidthOf`: what the frame is a share of is the section's, which the
  // reading does not know, so a width stated as one is carried across and the layout
  // is what turns it into points.
  readonly frameWidthShare?: number;
  readonly horizontal: AnchorPosition;
  readonly vertical: AnchorPosition;
  readonly behindDoc: boolean;
  readonly relativeHeight: number;
  readonly distances: WrapDistances;
};

// The geometry a shape's own element name states. `v:shape` is deliberately absent:
// what one of those draws is the `v:shapetype` it names, and reading it as a
// rectangle would paint a box over whatever the type actually draws.
const NAMED_GEOMETRIES: ReadonlyMap<string, ShapeGeometry> = new Map([
  ["rect", "rectangle"],
  ["roundrect", "rounded-rectangle"],
  ["oval", "ellipse"],
  ["line", "line"],
]);

// Whether the style turns the shape out of its own box, which VML writes in
// fractions of a degree: `rotation:1752415fd` and `rotation:-1945699fd` both stand
// in the corpus. Nothing here reads one, so what states one is not drawn.
function isTurned(style: ReadonlyMap<string, string>): boolean {
  const stated = style.get("rotation");
  if (stated === undefined) return false;
  const turn = Number(stated.toLowerCase().replace(/fd$/, ""));
  return !Number.isFinite(turn) || turn !== 0;
}

const UNFLIPPED: DrawingFlip = { horizontal: false, vertical: false };

/**
 * How a shape is turned over inside the box it was given, which VML writes as one
 * declaration naming the axes: `flip:y`, `flip:x`, `flip:xy`.
 */
function flipIn(style: ReadonlyMap<string, string>): DrawingFlip {
  const stated = style.get("flip")?.toLowerCase() ?? "";
  return { horizontal: stated.includes("x"), vertical: stated.includes("y") };
}

// One end of a line, `from="134.05pt,128.05pt"`, in the space a `margin-left` is
// measured in.
function endOf(stated: string | undefined): { x: number; y: number } | null {
  if (stated === undefined) return null;
  const [across, down] = stated.split(",");
  const x = lengthPt(across);
  const y = lengthPt(down);
  return x === null || y === null ? null : { x, y };
}

/**
 * Where a `v:line` stands, which it states in a way no other shape does.
 *
 * **A line states no offset and no size at all**: `from` and `to` are its two ends,
 * in the same space a `margin-left` is measured in, and the box it is drawn in is
 * the one they span. All nine in the corpus are written exactly so, with a
 * `position:absolute` and nothing else about where they are, which is why `frameOf`
 * turned every one of them down.
 *
 * **Which of the box's two diagonals it runs along is the two ends' own order**,
 * turned over again by whatever the style's `flip` says: a line whose `to` is above
 * its `from` runs up, and eight of the nine say `flip:y` beside coordinates that
 * already run down. What that composition comes to has not been asked of Word, and
 * on these nine it cannot be seen: every one of them spans 0.6pt or less from end to
 * end vertically over 300pt or more across.
 */
function lineFrameOf(shape: XmlElement, style: ReadonlyMap<string, string>): LegacyFrame | null {
  const from = endOf(attribute(shape, "", "from"));
  const to = endOf(attribute(shape, "", "to"));
  if (from === null || to === null) return null;

  const horizontal = endAt(style, HORIZONTAL, HORIZONTAL_ORIGINS, "column", Math.min(from.x, to.x));
  const vertical = endAt(style, VERTICAL, VERTICAL_ORIGINS, "paragraph", Math.min(from.y, to.y));
  if (horizontal === null || vertical === null) return null;

  const stated = flipIn(style);
  const depth = depthOf(style);
  return {
    widthEmu: emuOf(Math.abs(to.x - from.x)),
    heightEmu: emuOf(Math.abs(to.y - from.y)),
    flip: {
      horizontal: to.x < from.x !== stated.horizontal,
      vertical: to.y < from.y !== stated.vertical,
    },
    horizontal,
    vertical,
    behindDoc: depth < 0,
    relativeHeight: Math.abs(depth),
    distances: distancesOf(style),
  };
}

// Where a line stands on one axis. **`positionOn` cannot answer for a line**: it
// reads the offset the style states and a line states none, so what stands here is
// the same reading with the line's own nearer end in place of that offset. The
// origin and the alignment are read exactly as they are for a box.
function endAt(
  style: ReadonlyMap<string, string>,
  names: PositionNames,
  origins: ReadonlyMap<string, AnchorOrigin>,
  fallback: AnchorOrigin,
  atPt: number,
): AnchorPosition | null {
  const stated = style.get(names.relative);
  const from = stated === undefined ? fallback : origins.get(stated.toLowerCase());
  if (from === undefined) return null;

  const keyword = style.get(names.keyword)?.toLowerCase();
  if (keyword !== undefined && keyword !== "absolute")
    return { kind: "align", from, align: keyword };
  return { kind: "offset", from, offsetEmu: emuOf(atPt) };
}

// Word stacks its drawings by the size of the z-index and draws the ones below
// nought behind the text, which is where 60 of the corpus's 70 boxes are.
function depthOf(style: ReadonlyMap<string, string>): number {
  const stated = Number(style.get("z-index") ?? "");
  return Number.isFinite(stated) ? stated : 0;
}

function frameOf(style: ReadonlyMap<string, string>): LegacyFrame | null {
  const width = statedWidthOf(style);
  if (width.kind === "unreadable") return null;

  const widthPt = lengthPt(style.get("width"));
  const heightPt = lengthPt(style.get("height"));
  if (widthPt === null || heightPt === null || widthPt <= 0 || heightPt <= 0) return null;

  const horizontal = positionOn(style, HORIZONTAL, HORIZONTAL_ORIGINS, "column");
  const vertical = positionOn(style, VERTICAL, VERTICAL_ORIGINS, "paragraph");
  if (horizontal === null || vertical === null) return null;

  const depth = depthOf(style);

  return {
    widthEmu: emuOf(widthPt),
    heightEmu: emuOf(heightPt),
    flip: UNFLIPPED,
    ...(width.kind === "share-of-the-frame" ? { frameWidthShare: width.share } : {}),
    horizontal,
    vertical,
    behindDoc: depth < 0,
    relativeHeight: Math.abs(depth),
    distances: distancesOf(style),
  };
}

// Where the text sits in the box it is given, which VML states as one word and
// DrawingML as another. Every value the format has is named, though the corpus
// states only the top and the bottom.
const TEXT_ANCHORS: ReadonlyMap<string, TextBoxAnchor> = new Map([
  ["top", "top"],
  ["middle", "center"],
  ["bottom", "bottom"],
  ["top-center", "top"],
  ["middle-center", "center"],
  ["bottom-center", "bottom"],
  ["top-baseline", "top"],
  ["bottom-baseline", "bottom"],
]);

const ownChild = (node: XmlElement, namespace: string, name: string): XmlElement | undefined =>
  node.children.find((child) => child.namespace === namespace && child.name === name);

// What the text is held off the walls of its box by, `left,top,right,bottom` in one
// attribute, each of them a css length. Word's own default is a tenth of an inch at
// the sides and a twentieth above and below, which is DrawingML's default as well.
function insetsOf(textbox: XmlElement): TextBoxInsets {
  const stated = attribute(textbox, "", "inset");
  if (stated === undefined) return DEFAULT_TEXT_INSETS;

  const edges = stated.split(",");
  const edge = (at: number, fallback: number): number => {
    const points = lengthPt(edges[at]);
    return points === null ? fallback : emuOf(points);
  };
  return {
    leftEmu: edge(0, DEFAULT_TEXT_INSETS.leftEmu),
    topEmu: edge(1, DEFAULT_TEXT_INSETS.topEmu),
    rightEmu: edge(2, DEFAULT_TEXT_INSETS.rightEmu),
    bottomEmu: edge(3, DEFAULT_TEXT_INSETS.bottomEmu),
  };
}

function textBoxBodyOf(shape: XmlElement, style: ReadonlyMap<string, string>): TextBoxBody | null {
  const textbox = ownChild(shape, V_NS, "textbox");
  if (textbox === undefined) return null;

  // The content is what `isDetachedContent` holds off every other walk, so it is
  // reached by name here rather than by one of the helpers that skip it.
  const content = ownChild(textbox, W_NS, "txbxContent");
  if (content === undefined) return null;

  return {
    blocks: blocksIn(content),
    insets: insetsOf(textbox),
    anchor: TEXT_ANCHORS.get(style.get("v-text-anchor")?.toLowerCase() ?? "") ?? "top",
    // A box wraps its text inside itself unless it says otherwise, which is Word's
    // own default for one.
    wraps: style.get("mso-wrap-style")?.toLowerCase() !== "none",
    // "Resize shape to fit text", which the box states on itself rather than on the
    // shape around it.
    fitsText: declarationsOf(textbox).get("mso-fit-shape-to-text")?.toLowerCase() === "t",
  };
}

const HEX_COLOR = /^#?([0-9a-f]{6})$/i;

// A VML colour is a hex triple, sometimes with the palette entry Word remembers
// beside it (`#f2f2f2 [3052]`) and sometimes a name this keeps no table of.
function colorOf(raw: string): ColorReference | null {
  const stated = HEX_COLOR.exec(raw.trim().split(/\s+/)[0] ?? "");
  const hex = stated?.[1];
  return hex === undefined ? null : { base: { kind: "literal", hex }, ...UNSHADED };
}

const UNSHADED = { luminanceScale: 1, luminanceOffset: 0 };

// **A colour the file states and this cannot read is left unpainted.** Putting the
// default down instead would be a colour nobody asked for, over text.
const paintedIn = (stated: string | undefined, fallback: string): ColorReference | null =>
  stated === undefined ? colorOf(fallback) : colorOf(stated);

// VML writes a switch off as `f` or `false`; the attribute missing means on.
function turnedOff(shape: XmlElement, name: string): boolean {
  const stated = attribute(shape, "", name)?.toLowerCase();
  return stated === "f" || stated === "false" || stated === "0";
}

// **A VML shape is filled white and stroked black unless it says otherwise**, and
// every one of the corpus's 70 text boxes says otherwise about the stroke: 41 keep
// a fill, 32 of them a grey they name and nine the default.
const DEFAULT_FILL = "ffffff";
const DEFAULT_STROKE = "000000";
const DEFAULT_STROKE_WIDTH_PT = 0.75;

function paintOf(shape: XmlElement, geometry: ShapeGeometry = "rectangle"): ShapePaint {
  const strokeWidthPt = lengthPt(attribute(shape, "", "strokeweight"));
  const strokeColor = paintedIn(attribute(shape, "", "strokecolor"), DEFAULT_STROKE);

  return {
    fill: turnedOff(shape, "filled")
      ? null
      : paintedIn(attribute(shape, "", "fillcolor"), DEFAULT_FILL),
    outline:
      turnedOff(shape, "stroked") || strokeColor === null
        ? null
        : {
            color: strokeColor,
            widthPt: strokeWidthPt ?? DEFAULT_STROKE_WIDTH_PT,
            widthStated: strokeWidthPt !== null,
          },
    geometry,
    path: null,
  };
}

// The space a group states its children in, which is `a:chOff` and `a:chExt` under
// another name. Word writes either half empty for nought: `coordorigin=",-1244"`.
type GroupSpace = { readonly x: number; readonly y: number };

const NO_OFFSET: GroupSpace = { x: 0, y: 0 };

function spaceOf(raw: string | undefined): GroupSpace | null {
  if (raw === undefined) return null;
  const [across, down] = raw.split(",");
  const value = (stated: string | undefined): number =>
    stated === undefined || stated.trim() === "" ? 0 : Number(stated);

  const x = value(across);
  const y = value(down);
  return Number.isFinite(x) && Number.isFinite(y) ? { x, y } : null;
}

// Where one child stands in the box its group was given. **A child states itself in
// the group's own units rather than in points**, and one that states no `left` or no
// `top` stands at nought rather than at the origin.
type GroupPlace = {
  readonly leftFraction: number;
  readonly topFraction: number;
  readonly widthFraction: number;
  readonly heightFraction: number;
};

function placeInGroup(
  style: ReadonlyMap<string, string>,
  size: GroupSpace,
  origin: GroupSpace,
): GroupPlace | null {
  const coordinate = (name: string, fallback: number): number => {
    const stated = style.get(name);
    if (stated === undefined) return fallback;
    const value = Number(stated);
    return Number.isFinite(value) ? value : Number.NaN;
  };

  const left = coordinate("left", 0);
  const top = coordinate("top", 0);
  const width = coordinate("width", Number.NaN);
  const height = coordinate("height", Number.NaN);
  if ([left, top, width, height].some((each) => !Number.isFinite(each))) return null;

  return {
    leftFraction: (left - origin.x) / size.x,
    topFraction: (top - origin.y) / size.y,
    widthFraction: width / size.x,
    heightFraction: height / size.y,
  };
}

const UNTURNED = { flip: { horizontal: false, vertical: false }, turnDegrees: 0 };

// What a group draws of itself: the children this can read, and whether that was
// all of them. **A group half read is still drawn**, since a text box drawn in the
// right place is worth having whether or not the rule beside it could be; what was
// left out is what the report goes on naming.
type GroupReading = {
  readonly children: readonly GroupChild[];
  readonly whole: boolean;
  readonly leavesText: boolean;
};

const NOTHING_READ: GroupReading = { children: [], whole: false, leavesText: false };

function groupReadingOf(group: XmlElement): GroupReading {
  const size = spaceOf(attribute(group, "", "coordsize"));
  if (size === null || size.x === 0 || size.y === 0) {
    return { ...NOTHING_READ, leavesText: holdsATextBox(group) };
  }
  const origin = spaceOf(attribute(group, "", "coordorigin")) ?? NO_OFFSET;

  const children: GroupChild[] = [];
  let whole = true;
  let leavesText = false;

  for (const child of drawingsIn(group)) {
    const style = declarationsOf(child);
    const place = placeInGroup(style, size, origin);
    const content = place === null ? null : contentInGroupOf(child, style);
    if (place === null || content === null) {
      whole = false;
      leavesText = leavesText || holdsATextBox(child);
      continue;
    }
    children.push({ ...place, ...UNTURNED, content });
  }

  return { children, whole, leavesText };
}

// What one child of a group is drawn as: a text box, or a group of them again. A
// group inside a group is a space inside a space, and each level keeps the fractions
// of its own parent, so nothing has to be multiplied out here.
function contentInGroupOf(
  child: XmlElement,
  style: ReadonlyMap<string, string>,
): DrawingContent | null {
  if (child.name === "group") {
    const inside = groupReadingOf(child);
    return inside.children.length === 0 ? null : { kind: "group", children: inside.children };
  }

  const body = textBoxBodyOf(child, style);
  if (body !== null) return { kind: "text-box", body, paint: paintOf(child) };

  // The same reading `readLegacyDrawing` gives a shape standing on its own: its
  // paint, and the geometry its element name states. A group's own children are
  // where the corpus keeps most of its ovals and rounded rectangles.
  //
  // **A shape the style turns is left undrawn rather than drawn straight.** VML
  // states the turn in fractions of a degree (`rotation:1752415fd`) and nothing here
  // reads one, and a stake drawn upright where the file leans it over is a shape in
  // the wrong place rather than a shape half right.
  const geometry = NAMED_GEOMETRIES.get(child.name);
  if (geometry === undefined || imageOf(child) !== null || isTurned(style)) return null;
  return { kind: "shape", paint: paintOf(child, geometry) };
}

/**
 * What one drawing in the old form comes to: what the layout is given of it, and
 * what the fidelity report is left to name.
 *
 * **The two are answered together on purpose.** A box drawn in the right place must
 * stop being named, and one this cannot read must go on being named, and a reader
 * and a report that decide that separately drift apart the first time either
 * changes.
 *
 * What is read is the text box: out of the flow, sized in a unit this knows, and
 * positioned by an offset or an alignment it can name. What is not is everything
 * else in the form, and a box whose size the file states as a share of something
 * else.
 */
function readLegacyDrawing(shape: XmlElement): LegacyReading {
  const style = declarationsOf(shape);
  const undrawnWhole: LegacyReading = {
    drawing: null,
    undrawn: { holdsText: holdsATextBox(shape) },
  };

  // A drawing standing in the line is not one of these: the picture is drawn by
  // `inlinePictureOf` and anything else in the line is drawn nowhere at all.
  if (style.get("position")?.toLowerCase() !== "absolute") return undrawnWhole;

  // **A line is read before anything else asks about a size**, since it states
  // none: where it stands and how big it is are the same two coordinates.
  if (shape.name === "line") {
    const drawn = lineFrameOf(shape, style);
    return drawn === null
      ? undrawnWhole
      : {
          drawing: {
            ...drawn,
            name: attribute(shape, "", "alt") ?? "",
            content: { kind: "shape", paint: paintOf(shape, "line") },
          },
          undrawn: null,
        };
  }

  const frame = frameOf(style);
  if (frame === null) return undrawnWhole;
  const name = attribute(shape, "", "alt") ?? "";

  if (shape.name === "group") {
    const inside = groupReadingOf(shape);
    if (inside.children.length === 0) return undrawnWhole;
    return {
      drawing: { ...frame, name, content: { kind: "group", children: inside.children } },
      undrawn: inside.whole ? null : { holdsText: inside.leavesText },
    };
  }

  const body = textBoxBodyOf(shape, style);
  if (body !== null) {
    return {
      drawing: { ...frame, name, content: { kind: "text-box", body, paint: paintOf(shape) } },
      undrawn: null,
    };
  }

  // **A shape holding neither text nor a picture is its paint and its geometry**,
  // where its own element name states one. A `v:shape` is not one of those: what it
  // draws is the `v:shapetype` it names, which is a reading of its own, and a shape
  // holding a picture is a picture drawn nowhere yet rather than a box to fill in.
  const geometry = NAMED_GEOMETRIES.get(shape.name);
  if (geometry === undefined || imageOf(shape) !== null) return undrawnWhole;
  return {
    drawing: { ...frame, name, content: { kind: "shape", paint: paintOf(shape, geometry) } },
    undrawn: null,
  };
}

// Every drawing a `w:pict` or a `w:object` holds, read once for both the layout and
// the report. The picture standing in the line is left out of both: that one is
// `inlinePictureOf`'s, and the reader draws the first of them and nothing after it.
function readingsIn(pict: XmlElement): readonly LegacyReading[] {
  const readings: LegacyReading[] = [];
  let drawnInLine = false;

  for (const child of drawingsIn(pict)) {
    if (!drawnInLine && pictureInLine(child) !== null) {
      drawnInLine = true;
      continue;
    }
    readings.push(readLegacyDrawing(child));
  }

  return readings;
}

/**
 * What a `w:pict` or a `w:object` holds that nothing here draws.
 *
 * **One entry a drawing rather than one a shape**: a `v:group` answers once for
 * everything inside it, which is both how Word draws it and how the corpus was
 * counted. The `mc:Fallback` twin of a DrawingML shape never reaches here, since
 * the blocks drop it before anything looks inside.
 */
export const undrawnLegacyDrawings = (pict: XmlElement): readonly UndrawnLegacyDrawing[] =>
  readingsIn(pict).flatMap((each) => (each.undrawn === null ? [] : [each.undrawn]));

/**
 * The drawings a `w:pict` or a `w:object` hangs out of the flow, in the words a
 * `wp:anchor` answers in.
 *
 * **Nothing they state about wrapping is carried across**, and every one of them is
 * taken as wrapping nothing at all. `w10:wrap` stands on 68 of the corpus's 70
 * boxes and 65 of those state no type on it, so what Word makes of one is a default
 * nobody has asked it about; 60 of the 70 sit behind the text, where Word wraps
 * nothing. Taking them all as `none` leaves every line of the flow exactly where it
 * was broken before any of this was read, and puts the text of the box on the page,
 * which is the half of it that is not a guess.
 */
export const legacyAnchoredDrawingsIn = (pict: XmlElement): readonly LegacyAnchoredDrawing[] =>
  readingsIn(pict).flatMap((each) => (each.drawing === null ? [] : [each.drawing]));
