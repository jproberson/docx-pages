import { blocksIn, isDetachedContent, type Block } from "./blocks.js";
import { R_NS } from "./relationships.js";
import { W_NS } from "./section.js";
import { A_NS } from "./styles.js";
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

export type DrawingContent =
  | { readonly kind: "picture"; readonly relationshipId: string; readonly crop: CropInsets }
  | { readonly kind: "text-box"; readonly body: TextBoxBody }
  | { readonly kind: "shape" }
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

export function readDrawingContent(drawing: XmlElement): DrawingContent {
  const picture = findOwn(drawing, PIC_NS, "pic");
  if (picture !== null) return readPicture(picture);

  const shape = findOwn(drawing, WPS_NS, "wsp");
  if (shape !== null) {
    const txbx = findOwn(shape, WPS_NS, "txbx");
    return txbx === null
      ? { kind: "shape" }
      : { kind: "text-box", body: readTextBoxBody(shape, txbx) };
  }

  return { kind: "unknown" };
}
