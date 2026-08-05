import { DocxPagesError } from "@docx-pages/core";

export type PlacedRect = {
  readonly pageIndex: number;
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

export type ImagePlacement = {
  readonly kind: "image";
  readonly id: string;
  readonly rect: PlacedRect;
};

type Matrix = readonly [number, number, number, number, number, number];

const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

const multiply = (m: Matrix, n: Matrix): Matrix => [
  m[0] * n[0] + m[1] * n[2],
  m[0] * n[1] + m[1] * n[3],
  m[2] * n[0] + m[3] * n[2],
  m[2] * n[1] + m[3] * n[3],
  m[4] * n[0] + m[5] * n[2] + n[4],
  m[4] * n[1] + m[5] * n[3] + n[5],
];

const apply = (m: Matrix, x: number, y: number): readonly [number, number] => [
  m[0] * x + m[2] * y + m[4],
  m[1] * x + m[3] * y + m[5],
];

function unitSquareBounds(m: Matrix, pageHeightPt: number, pageIndex: number): PlacedRect {
  const corners = [apply(m, 0, 0), apply(m, 1, 0), apply(m, 0, 1), apply(m, 1, 1)];
  const xs = corners.map(([x]) => x);
  const ys = corners.map(([, y]) => y);
  const left = Math.min(...xs);
  const right = Math.max(...xs);
  const bottom = Math.min(...ys);
  const top = Math.max(...ys);
  return {
    pageIndex,
    leftPt: left,
    topPt: pageHeightPt - top,
    widthPt: right - left,
    heightPt: top - bottom,
  };
}

const isMatrix = (value: unknown): value is Matrix =>
  Array.isArray(value) && value.length === 6 && value.every((n) => typeof n === "number");

export async function readImagePlacements(bytes: Uint8Array): Promise<readonly ImagePlacement[]> {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");

  let document;
  try {
    document = await pdfjs.getDocument({ data: bytes, useSystemFonts: false }).promise;
  } catch (error: unknown) {
    throw new DocxPagesError({
      code: "pdf-unreadable",
      message: "the bytes are not a readable pdf",
      at: "render/pdf/placements.readImagePlacements",
      context: { byteLength: bytes.byteLength },
      cause: error,
    });
  }

  const placements: ImagePlacement[] = [];

  for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
    const page = await document.getPage(pageNumber);
    const [, , , pageHeightPt] = page.view;
    const operators = await page.getOperatorList();

    let current: Matrix = IDENTITY;
    const stack: Matrix[] = [];

    operators.fnArray.forEach((fn, index) => {
      const args: unknown = operators.argsArray[index];
      if (fn === pdfjs.OPS.save) {
        stack.push(current);
        return;
      }
      if (fn === pdfjs.OPS.restore) {
        current = stack.pop() ?? IDENTITY;
        return;
      }
      if (fn === pdfjs.OPS.transform && isMatrix(args)) {
        current = multiply(args, current);
        return;
      }
      if (fn === pdfjs.OPS.paintImageXObject && Array.isArray(args)) {
        const id: unknown = args[0];
        placements.push({
          kind: "image",
          id: typeof id === "string" ? id : `image-${String(placements.length)}`,
          rect: unitSquareBounds(current, pageHeightPt ?? 0, pageNumber - 1),
        });
      }
    });
  }

  return placements;
}
