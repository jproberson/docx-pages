import type { Paragraph } from "./blocks.js";
import {
  readDrawingContent,
  readDrawingFlip,
  readDrawingTurn,
  NO_PAINT,
  type DrawingContent,
  type DrawingFlip,
} from "./drawing.js";
import { paragraphOwnDrawings } from "./paragraphs.js";
import { W_NS } from "./section.js";
import { inlinePictureOf } from "./vml.js";
import { attribute, firstNamed, type XmlElement } from "./xml.js";

export const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

export type InlineDrawing = {
  readonly paragraphIndex: number;
  readonly name: string;
  readonly widthEmu: number;
  readonly heightEmu: number;
  // How far round the drawing was turned after it was drawn, clockwise. The extent
  // above is the drawing the right way up whatever this says.
  readonly turnDegrees: number;
  // Which way round the drawing was flipped after it was drawn. Paint alone: a
  // flipped box covers the room the unflipped one did, and only what is drawn
  // inside it is turned over.
  readonly flip: DrawingFlip;
  readonly content: DrawingContent;
};

const numberAttribute = (element: Parameters<typeof attribute>[0], name: string): number => {
  const raw = attribute(element, "", name);
  const value = raw === undefined ? Number.NaN : Number(raw);
  return Number.isFinite(value) ? value : 0;
};

// The two forms of drawing that stand on a line, **in the order the paragraph
// writes them**, since a drawing is matched to its place on the line by counting
// rather than by name: reading all of one form and then all of the other would
// hand a paragraph mixing them each other's pictures.
export function readInlines(paragraph: Paragraph): readonly InlineDrawing[] {
  return paragraphOwnDrawings(paragraph, [
    { namespace: WP_NS, name: "inline" },
    { namespace: W_NS, name: "pict" },
  ]).flatMap((element) =>
    element.namespace === WP_NS
      ? [drawnInline(paragraph, element)]
      : legacyPicture(paragraph, element),
  );
}

const drawnInline = (paragraph: Paragraph, inline: XmlElement): InlineDrawing => {
  const extent = firstNamed(inline, WP_NS, "extent");
  const docPr = firstNamed(inline, WP_NS, "docPr");
  return {
    paragraphIndex: paragraph.index,
    name: docPr === null ? "" : (attribute(docPr, "", "name") ?? ""),
    widthEmu: extent === null ? 0 : numberAttribute(extent, "cx"),
    heightEmu: extent === null ? 0 : numberAttribute(extent, "cy"),
    turnDegrees: readDrawingTurn(inline),
    flip: readDrawingFlip(inline),
    content: readDrawingContent(inline),
  };
};

// A `w:pict` that draws nothing on the line puts nothing on it, which is why this
// hands back a list rather than a drawing: `runs.ts` passes over the same ones,
// and the two have to agree about which they are.
function legacyPicture(paragraph: Paragraph, pict: XmlElement): readonly InlineDrawing[] {
  const picture = inlinePictureOf(pict);
  if (picture === null) return [];

  return [
    {
      paragraphIndex: paragraph.index,
      name: "",
      widthEmu: picture.widthEmu,
      heightEmu: picture.heightEmu,
      turnDegrees: 0,
      flip: { horizontal: false, vertical: false },
      content: {
        kind: "picture",
        relationshipId: picture.relationshipId,
        crop: { left: 0, top: 0, right: 0, bottom: 0 },
        paint: NO_PAINT,
      },
    },
  ];
}
