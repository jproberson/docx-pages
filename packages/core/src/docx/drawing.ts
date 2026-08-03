import { isDetachedContent } from "./blocks.js";
import { R_NS } from "./relationships.js";
import { A_NS } from "./styles.js";
import { attribute, type XmlElement } from "./xml.js";

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

export type DrawingContent =
  | { readonly kind: "picture"; readonly relationshipId: string; readonly crop: CropInsets }
  | { readonly kind: "text-box" }
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

export function readDrawingContent(drawing: XmlElement): DrawingContent {
  const picture = findOwn(drawing, PIC_NS, "pic");
  if (picture !== null) return readPicture(picture);

  const shape = findOwn(drawing, WPS_NS, "wsp");
  if (shape !== null) {
    return findOwn(shape, WPS_NS, "txbx") === null ? { kind: "shape" } : { kind: "text-box" };
  }

  return { kind: "unknown" };
}
