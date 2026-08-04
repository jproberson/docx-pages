import type { Paragraph } from "./blocks.js";
import { readDrawingContent, type DrawingContent } from "./drawing.js";
import { paragraphOwnDrawings } from "./paragraphs.js";
import { attribute, firstNamed } from "./xml.js";

export const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

export type InlineDrawing = {
  readonly paragraphIndex: number;
  readonly name: string;
  readonly widthEmu: number;
  readonly heightEmu: number;
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
      content: readDrawingContent(inline),
    };
  });
}
