import { OnePagerError } from "@onepager/core";

// A filled path, reported by the rectangle it covers on the page. Only the bounds
// are read: what a diagram draws is blocks and rules, and where a block lies is
// the whole of what a rendering of it has to agree with.
export type FillPlacement = {
  readonly kind: "fill";
  readonly pageIndex: number;
  readonly color: string;
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
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

const isMatrix = (value: unknown): value is Matrix =>
  Array.isArray(value) && value.length === 6 && value.every((n) => typeof n === "number");

// A path arrives as one run of numbers per subpath, each opening with the code for
// what to do next and carrying that many coordinates behind it.
const MOVE_TO = 0;
const LINE_TO = 1;
const CURVE_TO = 2;
const RECTANGLE = 3;

const STEPS: ReadonlyMap<number, number> = new Map([
  [MOVE_TO, 3],
  [LINE_TO, 3],
  [CURVE_TO, 7],
  [RECTANGLE, 5],
]);

// A subpath's coordinates arrive either as a plain list or as a typed one.
function numbersOf(value: unknown): readonly number[] {
  if (value instanceof Float32Array || value instanceof Float64Array) return Array.from(value);
  if (!Array.isArray(value)) return [];

  const numbers: number[] = [];
  for (const entry of value) {
    if (typeof entry !== "number") return [];
    numbers.push(entry);
  }
  return numbers;
}

// Curves are read by the points they are steered from, which stand outside the ink
// they draw; a diagram of blocks and rules has none of them.
function cornersOf(subpaths: unknown, current: Matrix): readonly (readonly [number, number])[] {
  if (!Array.isArray(subpaths)) return [];

  const corners: (readonly [number, number])[] = [];
  for (const subpath of subpaths) {
    const numbers = numbersOf(subpath);

    let at = 0;
    while (at < numbers.length) {
      const step = numbers[at] ?? -1;
      const stride = STEPS.get(step) ?? 1;
      const x = numbers[at + 1] ?? 0;
      const y = numbers[at + 2] ?? 0;
      if (step === MOVE_TO || step === LINE_TO) corners.push(apply(current, x, y));
      if (step === CURVE_TO)
        corners.push(apply(current, numbers[at + 5] ?? 0, numbers[at + 6] ?? 0));
      if (step === RECTANGLE) {
        corners.push(apply(current, x, y));
        corners.push(apply(current, x + (numbers[at + 3] ?? 0), y + (numbers[at + 4] ?? 0)));
      }
      at += stride;
    }
  }
  return corners;
}

export async function readFillPlacements(bytes: Uint8Array): Promise<readonly FillPlacement[]> {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");

  let document;
  try {
    document = await pdfjs.getDocument({ data: bytes, useSystemFonts: false }).promise;
  } catch (error: unknown) {
    throw new OnePagerError({
      code: "pdf-unreadable",
      message: "the bytes are not a readable pdf",
      at: "render/pdf/fills.readFillPlacements",
      context: { byteLength: bytes.byteLength },
      cause: error,
    });
  }

  const placements: FillPlacement[] = [];

  for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
    const page = await document.getPage(pageNumber);
    const [, , , pageHeightPt = 0] = page.view;
    const operators = await page.getOperatorList();

    let current: Matrix = IDENTITY;
    let color = "#000000";
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
      if (fn === pdfjs.OPS.setFillRGBColor && Array.isArray(args)) {
        const stated: unknown = args[0];
        if (typeof stated === "string") color = stated;
        return;
      }
      if (fn !== pdfjs.OPS.constructPath || !Array.isArray(args)) return;

      // A path states how it is to be painted, and one that only sets the clip or
      // is stroked rather than filled is not a block of colour.
      const paint: unknown = args[0];
      if (paint !== pdfjs.OPS.fill && paint !== pdfjs.OPS.eoFill) return;

      const corners = cornersOf(args[1], current);
      if (corners.length === 0) return;
      const xs = corners.map(([x]) => x);
      const ys = corners.map(([, y]) => y);
      const left = Math.min(...xs);
      const top = Math.max(...ys);
      placements.push({
        kind: "fill",
        pageIndex: pageNumber - 1,
        color,
        leftPt: left,
        topPt: pageHeightPt - top,
        widthPt: Math.max(...xs) - left,
        heightPt: top - Math.min(...ys),
      });
    });
  }

  return placements;
}
