import { blocksIn, isDetachedContent, type Block } from "./blocks.js";
import { R_NS } from "./relationships.js";
import { W_NS } from "./section.js";
import { A_NS } from "./styles.js";
import { readColorReference, type ColorReference } from "./theme.js";
import { attribute, firstNamed, type XmlElement } from "./xml.js";

export const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
export const WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

// Fractions of the source bitmap hidden on each edge, which is what srcRect
// records. A cropped picture is drawn larger than its extent by exactly this much.
export type CropInsets = {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
};

export const NO_CROP: CropInsets = { left: 0, top: 0, right: 0, bottom: 0 };

// Word's own defaults for a shape's text inset, in EMU: a tenth of an inch at the
// sides and a twentieth above and below.
export const DEFAULT_TEXT_INSETS: TextBoxInsets = {
  leftEmu: 91440,
  topEmu: 45720,
  rightEmu: 91440,
  bottomEmu: 45720,
};

export type TextBoxInsets = {
  readonly leftEmu: number;
  readonly topEmu: number;
  readonly rightEmu: number;
  readonly bottomEmu: number;
};

export type TextBoxAnchor = "top" | "center" | "bottom";

export type TextBoxBody = {
  readonly blocks: readonly Block[];
  readonly insets: TextBoxInsets;
  readonly anchor: TextBoxAnchor;
  readonly wraps: boolean;
  // A box that fits itself to its text keeps no stored height worth wrapping
  // around: Word measures the text again and wraps against that.
  readonly fitsText: boolean;
};

/**
 * What a preset geometry is to draw inside the box the object stands in.
 *
 * The reference documents named two of these and the corpus names more. Measured
 * on 2026-08-10 over the 40 corpus documents that hang a group of shapes in their
 * flow, whose 98 groups hold 2085 shapes between them: 244 ellipses, 51
 * rectangles, 18 straight connectors, 15 round rectangles, 5 triangles and one
 * arc. Everything this does not name is drawn as a rectangle, which is what the
 * whole of it used to be.
 */
export type ShapeGeometry =
  | "rectangle"
  | "line"
  | "ellipse"
  | "rounded-rectangle"
  | "triangle"
  // A path the file draws point by point, which nothing here plays. **Its
  // bounding box is not a fallback**: one corpus document rules a whole page with
  // a custom path, and drawing the box it fits in painted the page black under
  // everything else on it. Drawn as nothing until the path itself is read.
  | "custom";

// A connector is a line from one corner of its box to the other, which is what a
// `line` already was.
const GEOMETRIES = new Map<string, ShapeGeometry>([
  ["line", "line"],
  ["straightConnector1", "line"],
  ["ellipse", "ellipse"],
  ["roundRect", "rounded-rectangle"],
  ["triangle", "triangle"],
  ["rect", "rectangle"],
]);

// An outline takes its width along the shape's edge whether or not it paints
// anything there, which is what a box fitting itself to its text has to leave
// room for. A line with nothing to draw with carries no colour.
export type ShapeOutline = {
  readonly color: ColorReference | null;
  readonly widthPt: number;
  // Whether the file states that width or Word's own hairline stands in for it.
  // A box fitting itself to its text grows by a width the file states and by
  // nothing at all for one it does not, though Word draws both.
  readonly widthStated: boolean;
};

// How a shape is painted, before the theme has said what its colours are.
export type ShapePaint = {
  readonly fill: ColorReference | null;
  readonly outline: ShapeOutline | null;
  readonly geometry: ShapeGeometry;
};

export const NO_PAINT: ShapePaint = { fill: null, outline: null, geometry: "rectangle" };

/**
 * One shape inside a group, and where it stands in the box the group was given.
 *
 * **Kept as fractions of that box rather than as points**, because a group states
 * its children in a coordinate space of its own (`a:chOff` and `a:chExt`) that has
 * nothing to do with the room the flow gave it. Resolving the two into a fraction
 * here is what leaves everything below this free of the group's own arithmetic,
 * and it is what lets a group inside a group be the same thing again.
 */
export type GroupChild = {
  readonly leftFraction: number;
  readonly topFraction: number;
  readonly widthFraction: number;
  readonly heightFraction: number;
  readonly flip: DrawingFlip;
  // How far round the child was turned inside its group, which is its own turn and
  // not the group's: a group flattens onto its children before anything is drawn,
  // so a group's own turn is added to each of theirs there.
  readonly turnDegrees: number;
  readonly content: DrawingContent;
};

export type DrawingContent =
  | {
      readonly kind: "picture";
      readonly relationshipId: string;
      readonly crop: CropInsets;
      readonly paint: ShapePaint;
    }
  | { readonly kind: "text-box"; readonly body: TextBoxBody; readonly paint: ShapePaint }
  | { readonly kind: "shape"; readonly paint: ShapePaint }
  | { readonly kind: "group"; readonly children: readonly GroupChild[] }
  | { readonly kind: "unknown" };

// A text box's own drawings belong to the paragraphs inside it, not to the frame.
function findOwn(root: XmlElement, namespace: string, name: string): XmlElement | null {
  const visit = (node: XmlElement): XmlElement | null => {
    for (const child of node.children) {
      if (isDetachedContent(child)) continue;
      if (child.namespace === namespace && child.name === name) return child;
      const found = visit(child);
      if (found !== null) return found;
    }
    return null;
  };
  return visit(root);
}

const PERCENT_UNITS = 100000;

function cropEdge(srcRect: XmlElement | null, name: string): number {
  if (srcRect === null) return 0;
  const raw = attribute(srcRect, "", name);
  const value = raw === undefined ? Number.NaN : Number(raw);
  return Number.isFinite(value) ? value / PERCENT_UNITS : 0;
}

// Word leaves the width off an outline it draws at its own default, which it
// measures a hairline at: 9525 EMU, three quarters of a point.
const DEFAULT_OUTLINE_WIDTH_PT = 0.75;
const EMU_PER_POINT = 12700;

function readOutline(shapeProperties: XmlElement): ShapeOutline | null {
  const line = firstNamed(shapeProperties, A_NS, "ln");
  if (line === null) return null;

  const fill = firstNamed(line, A_NS, "solidFill");
  const color = fill === null ? null : readColorReference(fill);
  const raw = attribute(line, "", "w");
  const width = raw === undefined ? Number.NaN : Number(raw);
  const widthStated = Number.isFinite(width);
  return {
    color,
    widthPt: widthStated ? width / EMU_PER_POINT : DEFAULT_OUTLINE_WIDTH_PT,
    widthStated,
  };
}

// A shape with no fill element at all is as unpainted as one that says noFill:
// nothing in these documents leaves its fill to the theme's own format scheme.
function readPaint(shapeProperties: XmlElement | null): ShapePaint {
  if (shapeProperties === null) return NO_PAINT;

  const fill = firstNamed(shapeProperties, A_NS, "solidFill");
  const geometry = firstNamed(shapeProperties, A_NS, "prstGeom");
  const preset = geometry === null ? undefined : attribute(geometry, "", "prst");
  const custom = geometry === null && firstNamed(shapeProperties, A_NS, "custGeom") !== null;
  return {
    fill: fill === null ? null : readColorReference(fill),
    outline: readOutline(shapeProperties),
    geometry: custom
      ? "custom"
      : ((preset === undefined ? undefined : GEOMETRIES.get(preset)) ?? "rectangle"),
  };
}

const WPG_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

type Box = {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
};

const emu = (node: XmlElement | null, name: string): number => {
  const raw = node === null ? undefined : attribute(node, "", name);
  const value = raw === undefined ? Number.NaN : Number(raw);
  return Number.isFinite(value) ? value : 0;
};

function boxIn(transform: XmlElement | null, offset: string, extent: string): Box | null {
  if (transform === null) return null;
  const off = firstNamed(transform, A_NS, offset);
  const ext = firstNamed(transform, A_NS, extent);
  if (ext === null) return null;
  return { x: emu(off, "x"), y: emu(off, "y"), cx: emu(ext, "cx"), cy: emu(ext, "cy") };
}

const flipIn = (transform: XmlElement | null): DrawingFlip => ({
  horizontal: transform !== null && attribute(transform, "", "flipH") === "1",
  vertical: transform !== null && attribute(transform, "", "flipV") === "1",
});

// A turn is stated in sixtieths of a degree, clockwise, about the middle of the
// box the shape stands in.
const SIXTIETHS_OF_A_DEGREE = 60000;

const turnIn = (transform: XmlElement | null): number => {
  const raw = transform === null ? undefined : attribute(transform, "", "rot");
  const value = raw === undefined ? Number.NaN : Number(raw);
  return Number.isFinite(value) ? value / SIXTIETHS_OF_A_DEGREE : 0;
};

// Where a group's own properties live, which is one name under a group in the
// flow and another under a group inside one.
const groupProperties = (group: XmlElement): XmlElement | null =>
  firstNamed(group, WPG_NS, "grpSpPr") ?? firstNamed(group, A_NS, "grpSpPr");

// The properties of a child, whichever of the three kinds it is.
const shapeProperties = (child: XmlElement): XmlElement | null =>
  firstNamed(child, WPS_NS, "spPr") ??
  firstNamed(child, PIC_NS, "spPr") ??
  firstNamed(child, A_NS, "spPr") ??
  groupProperties(child);

/**
 * A group of shapes, each placed as a fraction of the box the group was given.
 *
 * **A group states its children in a space of its own.** `a:chOff` and `a:chExt`
 * say what that space is, and `a:ext` says how big the group is drawn; a child at
 * `a:off` in the first is drawn that far through the second. Where a group states
 * no child space the two are the same and every child is already a fraction of the
 * group, which is what Word writes for a group nothing has been scaled inside.
 */
function readGroup(group: XmlElement): DrawingContent {
  const transform =
    groupProperties(group) === null
      ? null
      : firstNamed(groupProperties(group) ?? group, A_NS, "xfrm");
  const own = boxIn(transform, "off", "ext");
  const space = boxIn(transform, "chOff", "chExt") ?? own;

  const acrossEmu = space?.cx === undefined || space.cx === 0 ? (own?.cx ?? 0) : space.cx;
  const downEmu = space?.cy === undefined || space.cy === 0 ? (own?.cy ?? 0) : space.cy;
  if (acrossEmu === 0 || downEmu === 0) return { kind: "unknown" };

  const children: GroupChild[] = [];
  for (const child of group.children) {
    const content = childContent(child);
    if (content === null) continue;

    const childTransform = firstNamed(shapeProperties(child) ?? child, A_NS, "xfrm");
    const at = boxIn(childTransform, "off", "ext");
    if (at === null) continue;

    children.push({
      leftFraction: (at.x - (space?.x ?? 0)) / acrossEmu,
      topFraction: (at.y - (space?.y ?? 0)) / downEmu,
      widthFraction: at.cx / acrossEmu,
      heightFraction: at.cy / downEmu,
      flip: flipIn(childTransform),
      turnDegrees: turnIn(childTransform),
      content,
    });
  }

  return children.length === 0 ? { kind: "unknown" } : { kind: "group", children };
}

// What one child of a group is, or nothing where it is a property of the group
// rather than something drawn.
function childContent(child: XmlElement): DrawingContent | null {
  if (child.namespace === PIC_NS && child.name === "pic") return readPicture(child);
  if (child.namespace === WPG_NS && (child.name === "grpSp" || child.name === "wgp")) {
    return readGroup(child);
  }
  if (child.namespace === WPS_NS && child.name === "wsp") return readShape(child);
  return null;
}

function readShape(shape: XmlElement): DrawingContent {
  const txbx = findOwn(shape, WPS_NS, "txbx");
  const paint = readPaint(firstNamed(shape, WPS_NS, "spPr"));
  return txbx === null
    ? { kind: "shape", paint }
    : { kind: "text-box", body: readTextBoxBody(shape, txbx), paint };
}

function readPicture(picture: XmlElement): DrawingContent {
  const blip = findOwn(picture, A_NS, "blip");
  const relationshipId = blip === null ? undefined : attribute(blip, R_NS, "embed");
  if (relationshipId === undefined || relationshipId === "") return { kind: "unknown" };

  const srcRect = findOwn(picture, A_NS, "srcRect");
  return {
    kind: "picture",
    relationshipId,
    crop: {
      left: cropEdge(srcRect, "l"),
      top: cropEdge(srcRect, "t"),
      right: cropEdge(srcRect, "r"),
      bottom: cropEdge(srcRect, "b"),
    },
    paint: readPaint(firstNamed(picture, PIC_NS, "spPr")),
  };
}

function insetEmu(bodyPr: XmlElement | null, name: string, fallback: number): number {
  const raw = bodyPr === null ? undefined : attribute(bodyPr, "", name);
  const value = raw === undefined ? Number.NaN : Number(raw);
  return Number.isFinite(value) ? value : fallback;
}

function anchorOf(bodyPr: XmlElement | null): TextBoxAnchor {
  const value = bodyPr === null ? undefined : attribute(bodyPr, "", "anchor");
  if (value === "ctr") return "center";
  if (value === "b") return "bottom";
  return "top";
}

function readTextBoxBody(shape: XmlElement, txbx: XmlElement): TextBoxBody {
  const content = firstNamed(txbx, W_NS, "txbxContent");
  const bodyPr = findOwn(shape, WPS_NS, "bodyPr") ?? findOwn(shape, A_NS, "bodyPr");

  return {
    blocks: content === null ? [] : blocksIn(content),
    insets: {
      leftEmu: insetEmu(bodyPr, "lIns", DEFAULT_TEXT_INSETS.leftEmu),
      topEmu: insetEmu(bodyPr, "tIns", DEFAULT_TEXT_INSETS.topEmu),
      rightEmu: insetEmu(bodyPr, "rIns", DEFAULT_TEXT_INSETS.rightEmu),
      bottomEmu: insetEmu(bodyPr, "bIns", DEFAULT_TEXT_INSETS.bottomEmu),
    },
    anchor: anchorOf(bodyPr),
    wraps: (bodyPr === null ? undefined : attribute(bodyPr, "", "wrap")) !== "none",
    fitsText: bodyPr !== null && firstNamed(bodyPr, A_NS, "spAutoFit") !== null,
  };
}

// Which way round an object was turned after it was drawn. A wrap polygon is
// written for the object the right way round, so a flip turns the polygon over
// with it.
export type DrawingFlip = {
  readonly horizontal: boolean;
  readonly vertical: boolean;
};

export const readDrawingFlip = (drawing: XmlElement): DrawingFlip =>
  flipIn(findOwn(drawing, A_NS, "xfrm"));

/**
 * How far round the object itself was turned after it was drawn, clockwise.
 *
 * The extent the flow was given is the object the right way up, so a turn is
 * something the layout has to answer for rather than a matter of paint alone: a
 * quarter turn leaves a picture as wide as it was tall. What that costs the line
 * it stands on is `roomForTurn`, and where it is then painted is `boundsOfTurn`.
 */
export const readDrawingTurn = (drawing: XmlElement): number =>
  turnIn(findOwn(drawing, A_NS, "xfrm"));

export function readDrawingContent(drawing: XmlElement): DrawingContent {
  // **The group is asked for first, and that is the whole of why this order
  // matters.** A group holds pictures and shapes inside it, so asking for a
  // picture first hands back whichever one happens to sit in the group and draws
  // it at the group's own place and size, with everything else in the group drawn
  // nowhere at all. One corpus document drew a photograph's cropped middle where
  // Word draws a diagram of 323 shapes, and nothing in the report said a word.
  const group = findOwn(drawing, WPG_NS, "wgp");
  if (group !== null) return readGroup(group);

  const picture = findOwn(drawing, PIC_NS, "pic");
  if (picture !== null) return readPicture(picture);

  const shape = findOwn(drawing, WPS_NS, "wsp");
  if (shape !== null) return readShape(shape);

  return { kind: "unknown" };
}
