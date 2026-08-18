import type { AnchorOrigin, AnchorPosition, WrapDistances } from "./anchors.js";
import { blocksIn, isDetachedContent } from "./blocks.js";
import {
  DEFAULT_TEXT_INSETS,
  type CropInsets,
  type DrawingContent,
  type GroupChild,
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

// **A size stated as a share of something else is refused rather than guessed at.**
// 26 items in eight corpus documents state `mso-width-percent:400` or
// `mso-height-percent:200` beside a width and a height in points on the same shape,
// and which of the two Word draws has not been asked of it.
const SHARES = ["mso-width-percent", "mso-height-percent"];

const sizedByShare = (style: ReadonlyMap<string, string>): boolean =>
  SHARES.some((name) => {
    const stated = style.get(name);
    if (stated === undefined) return false;
    const share = Number(stated);
    return !Number.isFinite(share) || share !== 0;
  });

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
  readonly horizontal: AnchorPosition;
  readonly vertical: AnchorPosition;
  readonly behindDoc: boolean;
  readonly relativeHeight: number;
  readonly distances: WrapDistances;
};

function frameOf(style: ReadonlyMap<string, string>): LegacyFrame | null {
  if (sizedByShare(style)) return null;

  const widthPt = lengthPt(style.get("width"));
  const heightPt = lengthPt(style.get("height"));
  if (widthPt === null || heightPt === null || widthPt <= 0 || heightPt <= 0) return null;

  const horizontal = positionOn(style, HORIZONTAL, HORIZONTAL_ORIGINS, "column");
  const vertical = positionOn(style, VERTICAL, VERTICAL_ORIGINS, "paragraph");
  if (horizontal === null || vertical === null) return null;

  // Word stacks its drawings by the size of the z-index and draws the ones below
  // nought behind the text, which is where 60 of the corpus's 70 boxes are.
  const stated = Number(style.get("z-index") ?? "");
  const depth = Number.isFinite(stated) ? stated : 0;

  return {
    widthEmu: emuOf(widthPt),
    heightEmu: emuOf(heightPt),
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

function paintOf(shape: XmlElement): ShapePaint {
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
    geometry: "rectangle",
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
  return body === null ? null : { kind: "text-box", body, paint: paintOf(child) };
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
  if (body === null) return undrawnWhole;
  return {
    drawing: { ...frame, name, content: { kind: "text-box", body, paint: paintOf(shape) } },
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
