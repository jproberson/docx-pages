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
  // Where the image itself was put, which is the whole of it however little shows.
  readonly rect: PlacedRect;
  // **The part of it the clip lets through, which is the part that is ink.** Word
  // writes a cropped picture by scaling the whole image up and clipping to the part
  // it wants, so `rect` above answers with a box many times the size of anything
  // drawn: one corpus letterhead is placed 3028.9 by 386.8pt at a left of -2336.9
  // and drawn 251.1 by 76.9 at 55.8, and a reference document's logo is placed
  // 254.6 wide and drawn 180. Null where the clip leaves nothing at all.
  //
  // **Read this and not `rect` when the question is what the page holds**, and
  // `rect` when the question is what the document asked for.
  readonly inkRect: PlacedRect | null;
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

// A box in the pdf's own space, whose y runs up from the foot of the page.
type Box = {
  readonly minX: number;
  readonly minY: number;
  readonly maxX: number;
  readonly maxY: number;
};

const boundsOf = (corners: readonly (readonly [number, number])[]): Box => ({
  minX: Math.min(...corners.map(([x]) => x)),
  minY: Math.min(...corners.map(([, y]) => y)),
  maxX: Math.max(...corners.map(([x]) => x)),
  maxY: Math.max(...corners.map(([, y]) => y)),
});

// An image is drawn by putting the unit square where the matrix says.
const unitSquareBounds = (m: Matrix): Box =>
  boundsOf([apply(m, 0, 0), apply(m, 1, 0), apply(m, 0, 1), apply(m, 1, 1)]);

const rectangleBounds = (m: Matrix, box: Box): Box =>
  boundsOf([
    apply(m, box.minX, box.minY),
    apply(m, box.maxX, box.minY),
    apply(m, box.minX, box.maxY),
    apply(m, box.maxX, box.maxY),
  ]);

// What the two have in common, or null where they have nothing: an image clipped
// away entirely draws no ink and is not a placement at all.
function meeting(one: Box, other: Box | null): Box | null {
  if (other === null) return one;
  const box = {
    minX: Math.max(one.minX, other.minX),
    minY: Math.max(one.minY, other.minY),
    maxX: Math.min(one.maxX, other.maxX),
    maxY: Math.min(one.maxY, other.maxY),
  };
  return box.maxX > box.minX && box.maxY > box.minY ? box : null;
}

const placedAt = (box: Box, pageHeightPt: number, pageIndex: number): PlacedRect => ({
  pageIndex,
  leftPt: box.minX,
  topPt: pageHeightPt - box.maxY,
  widthPt: box.maxX - box.minX,
  heightPt: box.maxY - box.minY,
});

const isMatrix = (value: unknown): value is Matrix =>
  Array.isArray(value) && value.length === 6 && value.every((n) => typeof n === "number");

// A path's own extent, which `constructPath` hands over already worked out as
// `[minX, minY, maxX, maxY]` beside the segments themselves.
function extentOf(args: unknown): Box | null {
  if (!Array.isArray(args)) return null;
  const stated: unknown = args[2];
  const numbers: readonly number[] =
    stated instanceof Float32Array || stated instanceof Float64Array
      ? Array.from(stated)
      : Array.isArray(stated)
        ? stated.filter((each): each is number => typeof each === "number")
        : [];

  const [minX, minY, maxX, maxY] = numbers;
  if (minX === undefined || minY === undefined || maxX === undefined || maxY === undefined) {
    return null;
  }
  return { minX, minY, maxX, maxY };
}

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
    // **What a clip leaves is what the page actually holds.** Word writes a cropped
    // picture by scaling the whole image up and clipping to the part it wants, so
    // the matrix alone answers with a box many times the size of the ink: one corpus
    // letterhead reports 3028.9 by 386.8pt at a left of -2336.9, and is drawn 251.1
    // by 76.9 at 55.8. Read without the clip, an image comparison would call every
    // cropped picture misplaced.
    let clip: Box | null = null;
    let clipping = false;
    const stack: { readonly matrix: Matrix; readonly clip: Box | null }[] = [];

    operators.fnArray.forEach((fn, index) => {
      const args: unknown = operators.argsArray[index];
      if (fn === pdfjs.OPS.save) {
        stack.push({ matrix: current, clip });
        return;
      }
      if (fn === pdfjs.OPS.restore) {
        const held = stack.pop();
        current = held?.matrix ?? IDENTITY;
        clip = held?.clip ?? null;
        return;
      }
      if (fn === pdfjs.OPS.transform && isMatrix(args)) {
        current = multiply(args, current);
        return;
      }
      // The operator list names the clip before the path that draws it, so the
      // rectangle is the next one built and not the one already behind us.
      if (fn === pdfjs.OPS.clip || fn === pdfjs.OPS.eoClip) {
        clipping = true;
        return;
      }
      if (fn === pdfjs.OPS.constructPath) {
        if (!clipping) return;
        clipping = false;
        const extent = extentOf(args);
        if (extent !== null) clip = meeting(rectangleBounds(current, extent), clip);
        return;
      }
      if (fn === pdfjs.OPS.paintImageXObject && Array.isArray(args)) {
        const placed = unitSquareBounds(current);
        const drawn = meeting(placed, clip);
        const id: unknown = args[0];
        placements.push({
          kind: "image",
          id: typeof id === "string" ? id : `image-${String(placements.length)}`,
          rect: placedAt(placed, pageHeightPt ?? 0, pageNumber - 1),
          inkRect: drawn === null ? null : placedAt(drawn, pageHeightPt ?? 0, pageNumber - 1),
        });
      }
    });
  }

  return placements;
}
