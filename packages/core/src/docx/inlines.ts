import type { Paragraph } from "./blocks.js";
import {
  readDrawingContent,
  readDrawingFlip,
  readDrawingTurn,
  type DrawingContent,
  type DrawingFlip,
} from "./drawing.js";
import { paragraphOwnDrawings } from "./paragraphs.js";
import { attribute, firstNamed } from "./xml.js";

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

export function readInlines(paragraph: Paragraph): readonly InlineDrawing[] {
  return paragraphOwnDrawings(paragraph, WP_NS, "inline").map((inline) => {
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
  });
}
