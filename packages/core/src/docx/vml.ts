import { isDetachedContent } from "./blocks.js";
import { R_NS } from "./relationships.js";
import { attribute, type XmlElement } from "./xml.js";

// The drawing form Word wrote before DrawingML and still writes for some of what
// it draws: a `w:pict` holding a `v:shape` whose `v:imagedata` names the picture.
//
// **Only the picture standing in the run's own line is read here.** A VML shape
// carrying `position:absolute` is out of flow like a `wp:anchor`, and one holding
// a `v:textbox` rather than a picture is a text box; neither is answered for yet.
// The `mc:Fallback` twin of a DrawingML shape never reaches this at all, since the
// blocks drop it before anything looks inside.
//
// Measured over the 718: every inline VML picture in the corpus states its width
// and height in points and none of them crops, so a size stated any other way is
// left at nothing rather than guessed at.

export const V_NS = "urn:schemas-microsoft-com:vml";

export type VmlPicture = {
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly relationshipId: string;
};

const EMU_PER_POINT = 12700;

// A VML shape's `style` is css, so it is read by declaration rather than by
// position, and a declaration this does not know is passed over.
function declaration(shape: XmlElement, name: string): string | null {
  const style = attribute(shape, "", "style");
  if (style === undefined) return null;

  for (const each of style.split(";")) {
    const at = each.indexOf(":");
    if (at < 0) continue;
    if (each.slice(0, at).trim().toLowerCase() !== name) continue;
    return each.slice(at + 1).trim();
  }
  return null;
}

const pointsOf = (shape: XmlElement, name: string): number => {
  const raw = declaration(shape, name);
  if (raw === null || !raw.toLowerCase().endsWith("pt")) return 0;
  const value = Number(raw.slice(0, -2));
  return Number.isFinite(value) ? value : 0;
};

const isPositioned = (shape: XmlElement): boolean =>
  declaration(shape, "position")?.toLowerCase() === "absolute";

// The shapes VML draws a picture with. A `v:group` is a drawing of its own and is
// not one of them: reading its first picture would draw that one at the group's
// size and the rest nowhere at all.
const PICTURE_SHAPES = new Set(["shape", "rect", "roundrect", "oval"]);

function imageOf(node: XmlElement): string | null {
  for (const child of node.children) {
    if (isDetachedContent(child)) continue;
    if (child.namespace === V_NS && child.name === "imagedata") {
      const id = attribute(child, R_NS, "id");
      return id === undefined || id === "" ? null : id;
    }
    const found = imageOf(child);
    if (found !== null) return found;
  }
  return null;
}

/**
 * The picture a `w:pict` puts on the line, or null where it puts none there:
 * an empty `w:pict`, a shape that is not a picture, or one positioned out of the
 * flow.
 */
export function inlinePictureOf(pict: XmlElement): VmlPicture | null {
  for (const child of pict.children) {
    if (isDetachedContent(child)) continue;
    if (child.namespace !== V_NS || !PICTURE_SHAPES.has(child.name)) continue;
    if (isPositioned(child)) continue;

    const relationshipId = imageOf(child);
    if (relationshipId === null) continue;

    return {
      widthEmu: pointsOf(child, "width") * EMU_PER_POINT,
      heightEmu: pointsOf(child, "height") * EMU_PER_POINT,
      relationshipId,
    };
  }
  return null;
}
