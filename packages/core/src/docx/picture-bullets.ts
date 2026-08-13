import { type Block } from "./blocks.js";
import { NUMBERING_PART } from "./numbering.js";
import { partXml, type DocxPackage } from "./package.js";
import { drawablePicture } from "./pictures.js";
import { readRelationships, R_NS } from "./relationships.js";
import { W_NS } from "./section.js";
import { inlinePictureOf, V_NS } from "./vml.js";
import { clark, firstNamed, type XmlElement } from "./xml.js";

// What Word draws for a `w:numPicBullet`, which is not a bullet.
//
// **Measured on 2026-08-12, over eight probes and every corpus document declaring
// one.** A numbering part's picture bullet is drawn exactly once, at the head of the
// first paragraph of the body, and it is drawn there whether or not any level points
// at it. Where a level does point at it with `w:lvlPicBulletId`, the paragraphs
// wearing that level are given no mark at all: this Word draws a picture bullet
// nowhere as a bullet, so nothing here treats it as one.
//
// The probes, each written by hand and exported by Word:
//
// | the document                                  | what Word drew                          |
// | --------------------------------------------- | --------------------------------------- |
// | one bullet, no level naming it                | the picture at the first paragraph       |
// | a level naming it, worn by all three          | the picture at the first paragraph alone |
// | the first paragraph centred                   | centred with the line, the text after it |
// | the first paragraph empty                     | the picture on its own line             |
// | the first block a table                       | inside its first cell, at the cell's text |
// | a decimal list beside it                      | the numbers untouched, the picture once  |
// | the bullet cropped                            | the stated size is what the crop lets through |
// | two bullets declared                          | a document Word will not open at all    |
//
// So the picture stands where an inline picture at the head of the paragraph would:
// its bottom on the first line's baseline, the line's ascent its own height, the
// text following at its right edge, and the whole line aligned as the paragraph
// asks. That is what this reads it as, which is why the rest of the layout needs to
// hear nothing about picture bullets: breaking, height, alignment and drawing are
// answered by the rules an inline picture already has.
//
// The seven corpus documents this was worth the most to declare a bullet 250pt wide,
// which eats the first line and wraps the title under it; the other thirty declare
// one a few points across and it costs them a line's height.

export type PictureBullet = {
  // The name the standing-in drawing gives the picture. A relationship id is a name
  // in one part, and this one is a name in the numbering part while the drawing is
  // read with the body's own: `rId1` is the picture in the first and the styles in
  // the second. So the drawing is given a name of neither, which no package can
  // have written, since a relationship id is an XML `ID` and holds no slash.
  readonly relationshipId: string;
  readonly part: string;
  // The run that stands at the head of the paragraph wearing the bullet.
  readonly run: XmlElement;
};

const element = (
  namespace: string,
  name: string,
  children: readonly XmlElement[],
  attributes: ReadonlyMap<string, string> = new Map(),
): XmlElement => ({
  namespace,
  name,
  attributes,
  children,
  text: "",
  preservesSpace: false,
});

// The picture the shape names, under the name the standing-in drawing uses instead.
function renamed(node: XmlElement, relationshipId: string): XmlElement {
  if (node.namespace === V_NS && node.name === "imagedata") {
    const attributes = new Map(node.attributes);
    attributes.set(clark(R_NS, "id"), relationshipId);
    return { ...node, attributes };
  }
  return { ...node, children: node.children.map((child) => renamed(child, relationshipId)) };
}

/**
 * The picture a document's numbering part declares as a bullet, or null where it
 * declares none, names no picture, or names a part this cannot draw.
 *
 * Only the first is read. A document declaring two is one Word refuses to open, so
 * there is no second for anything to agree with.
 */
export function pictureBulletOf(pkg: DocxPackage): PictureBullet | null {
  if (!pkg.parts.has(NUMBERING_PART)) return null;

  const declared = firstNamed(partXml(pkg, NUMBERING_PART), W_NS, "numPicBullet");
  const pict = declared === null ? null : firstNamed(declared, W_NS, "pict");
  const picture = pict === null ? null : inlinePictureOf(pict);
  if (pict === null || picture === null) return null;

  const part = readRelationships(pkg, NUMBERING_PART).get(picture.relationshipId)?.part;
  if (part === undefined || !pkg.parts.has(part) || !drawablePicture(part, pkg.parts.get(part)))
    return null;

  const relationshipId = `${NUMBERING_PART}#${picture.relationshipId}`;
  return {
    relationshipId,
    part,
    run: element(W_NS, "r", [renamed(pict, relationshipId)]),
  };
}

// A paragraph writes its properties first and Word draws the bullet before the
// paragraph's own first run, so the run stands after `w:pPr` and ahead of everything
// else.
function wearing(paragraph: XmlElement, run: XmlElement): XmlElement {
  const properties = paragraph.children.filter(
    (child) => child.namespace === W_NS && child.name === "pPr",
  );
  return {
    ...paragraph,
    children: [
      ...properties,
      run,
      ...paragraph.children.filter((child) => !properties.includes(child)),
    ],
  };
}

/**
 * The same blocks with the bullet standing at the head of the first paragraph, which
 * is the first in the order the body writes them and is inside a table's first cell
 * where the body opens with a table.
 */
export function wearingPictureBullet(
  blocks: readonly Block[],
  bullet: PictureBullet,
): readonly Block[] {
  let worn = false;

  const dressed = (of: readonly Block[]): readonly Block[] =>
    of.map((block) => {
      if (worn) return block;
      if (block.kind === "paragraph") {
        worn = true;
        return {
          ...block,
          paragraph: { ...block.paragraph, element: wearing(block.paragraph.element, bullet.run) },
        };
      }
      return {
        ...block,
        rows: block.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({ ...cell, blocks: dressed(cell.blocks) })),
        })),
      };
    });

  return dressed(blocks);
}
