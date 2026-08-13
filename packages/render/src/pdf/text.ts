import { DocxPagesError } from "@docx-pages/core";

export type TextPlacement = {
  readonly kind: "text";
  readonly text: string;
  readonly fontName: string;
  readonly pageIndex: number;
  readonly leftPt: number;
  readonly baselinePt: number;
  readonly widthPt: number;
  readonly fontSizePt: number;
};

type TextItemShape = {
  readonly str: string;
  readonly transform: readonly number[];
  readonly width: number;
  readonly fontName: string;
};

function parseTextItem(value: unknown): TextItemShape | null {
  if (typeof value !== "object" || value === null) return null;
  const candidate: Record<string, unknown> = { ...value };

  const str = candidate["str"];
  const width = candidate["width"];
  const fontName = candidate["fontName"];
  const rawTransform = candidate["transform"];
  if (typeof str !== "string" || typeof width !== "number") return null;
  if (typeof fontName !== "string" || !Array.isArray(rawTransform)) return null;

  const transform: number[] = [];
  for (const entry of rawTransform) {
    if (typeof entry !== "number") return null;
    transform.push(entry);
  }
  if (transform.length !== 6) return null;

  return { str, width, fontName, transform };
}

// Everything a pdf says about the text it draws: where each item was drawn, and how many
// pages the file holds.
//
// **The page count cannot be read off the items and has to come from the file.** A page
// drawing nothing but a picture holds no text item at all, so counting pages by the
// highest item was short by one on 19 of the 715 corpus documents, every one of them a
// last page with no text on it. Read that way `deformed.ts` accused 21 clean documents of
// making the wrong number of pages where the rasteriser, which counts the pdf's own
// pages, says 5.
export type DrawnText = {
  readonly placements: readonly TextPlacement[];
  readonly pages: number;
};

export const readTextPlacements = async (bytes: Uint8Array): Promise<readonly TextPlacement[]> =>
  (await readDrawnText(bytes)).placements;

export async function readDrawnText(bytes: Uint8Array): Promise<DrawnText> {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");

  let document;
  try {
    document = await pdfjs.getDocument({ data: bytes, useSystemFonts: false }).promise;
  } catch (error: unknown) {
    throw new DocxPagesError({
      code: "pdf-unreadable",
      message: "the bytes are not a readable pdf",
      at: "render/pdf/text.readTextPlacements",
      context: { byteLength: bytes.byteLength },
      cause: error,
    });
  }

  const placements: TextPlacement[] = [];

  for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
    const page = await document.getPage(pageNumber);
    const [, , , pageHeightPt = 0] = page.view;
    const content = await page.getTextContent();

    for (const rawItem of content.items) {
      const item = parseTextItem(rawItem);
      if (item === null || item.str === "") continue;
      const [scaleX = 0, , , scaleY = 0, offsetX = 0, offsetY = 0] = item.transform;
      placements.push({
        kind: "text",
        text: item.str,
        fontName: item.fontName,
        pageIndex: pageNumber - 1,
        leftPt: offsetX,
        baselinePt: pageHeightPt - offsetY,
        widthPt: item.width,
        fontSizePt: Math.hypot(scaleX, scaleY) === 0 ? 0 : Math.abs(scaleY),
      });
    }
  }

  return { placements, pages: document.numPages };
}
