import { isDetachedContent } from "./blocks.js";
import { type CropInsets } from "./drawing.js";
import { R_NS } from "./relationships.js";
import { W_NS } from "./section.js";
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
// a `v:textbox` rather than a picture is a text box; neither is answered for yet.
// The `mc:Fallback` twin of a DrawingML shape never reaches this at all, since the
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

/**
 * The picture a `w:pict` or a `w:object` puts on the line, or null where it puts
 * none there: an empty one, a shape that is not a picture, or one positioned out
 * of the flow.
 */
export function inlinePictureOf(pict: XmlElement): VmlPicture | null {
  for (const child of pict.children) {
    if (isDetachedContent(child)) continue;
    if (child.namespace !== V_NS || !PICTURE_SHAPES.has(child.name)) continue;
    if (isPositioned(child)) continue;

    const imagedata = imageOf(child);
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
