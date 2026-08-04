import type { Block, CellVerticalAlign, Paragraph, TableRow } from "../docx/blocks.js";
import { paragraphOwnDrawings } from "../docx/paragraphs.js";
import {
  resolveParagraphMark,
  resolveRunMarks,
  type ParagraphMark,
  type StyleTable,
} from "../docx/styles.js";
import { attribute } from "../docx/xml.js";
import { lineHeightPt, type FaceRequest, type MetricsLookup } from "./font-metrics.js";
import { EMU_PER_POINT } from "./units.js";

export const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

export type MetricsResolver = (request: FaceRequest) => MetricsLookup;

export type LayoutBlocker =
  | { readonly kind: "unresolved-font"; readonly part: string; readonly paragraphIndex: number }
  | {
      readonly kind: "unknown-font-metrics";
      readonly part: string;
      readonly paragraphIndex: number;
      readonly fontName: string;
    };

export type ParagraphBox = {
  readonly index: number;
  readonly topPt: number;
  readonly heightPt: number;
};

export type StackMeasurement =
  | {
      readonly kind: "measured";
      readonly boxes: readonly ParagraphBox[];
      readonly heightPt: number;
    }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

export type MeasureStackInput = {
  readonly blocks: readonly Block[];
  readonly styles: StyleTable;
  readonly metricsFor: MetricsResolver;
  readonly part: string;
  readonly originPt: number;
};

type Context = Omit<MeasureStackInput, "blocks" | "originPt">;

type MarkHeight =
  { readonly kind: "height"; readonly value: number } | { readonly kind: "blocked" };

function inlineHeightPt(paragraph: Paragraph): number {
  let tallest = 0;
  for (const inline of paragraphOwnDrawings(paragraph, WP_NS, "inline")) {
    for (const extent of inline.children) {
      if (extent.namespace !== WP_NS || extent.name !== "extent") continue;
      const cy = Number(attribute(extent, "", "cy") ?? "0");
      if (Number.isFinite(cy)) tallest = Math.max(tallest, cy / EMU_PER_POINT);
    }
  }
  return tallest;
}

export const measureStack = (input: MeasureStackInput): StackMeasurement =>
  measureBlocks(input.blocks, input, input.originPt);

function measureBlocks(
  blocks: readonly Block[],
  context: Context,
  originPt: number,
): StackMeasurement {
  const boxes: ParagraphBox[] = [];
  let top = originPt;

  for (const block of blocks) {
    if (block.kind === "paragraph") {
      const height = paragraphHeight(block.paragraph, context);
      if (height.kind === "blocked") {
        return {
          kind: "blocked",
          blocker: blockerFor(height.mark, context.part, block.paragraph.index),
        };
      }
      boxes.push({ index: block.paragraph.index, topPt: top, heightPt: height.value });
      top += height.value;
      continue;
    }

    for (const row of block.rows) {
      const measured = measureRow(row, context, top);
      if (measured.kind === "blocked") return measured;
      boxes.push(...measured.boxes);
      top += measured.heightPt;
    }
  }

  return { kind: "measured", boxes, heightPt: top - originPt };
}

type MeasuredCell = {
  readonly align: CellVerticalAlign;
  readonly boxes: readonly ParagraphBox[];
  readonly heightPt: number;
};

// A row is as tall as its tallest cell, and every cell starts at the row's top;
// cells sit beside each other, so their heights never add up.
function measureRow(row: TableRow, context: Context, topPt: number): StackMeasurement {
  const measured: MeasuredCell[] = [];
  let heightPt = 0;

  for (const cell of row.cells) {
    const of = measureBlocks(cell.blocks, context, 0);
    if (of.kind === "blocked") return of;
    measured.push({ align: cell.verticalAlign, boxes: of.boxes, heightPt: of.heightPt });
    heightPt = Math.max(heightPt, of.heightPt);
  }

  const boxes: ParagraphBox[] = [];
  for (const cell of measured) {
    const offset = topPt + seatingOffset(cell.align, heightPt, cell.heightPt);
    for (const box of cell.boxes) boxes.push({ ...box, topPt: box.topPt + offset });
  }

  return { kind: "measured", boxes, heightPt };
}

function seatingOffset(
  align: CellVerticalAlign,
  rowHeightPt: number,
  cellHeightPt: number,
): number {
  const slack = rowHeightPt - cellHeightPt;
  if (align === "center") return slack / 2;
  if (align === "bottom") return slack;
  return 0;
}

type ParagraphHeight =
  | { readonly kind: "height"; readonly value: number }
  | { readonly kind: "blocked"; readonly mark: ParagraphMark };

function paragraphHeight(paragraph: Paragraph, context: Context): ParagraphHeight {
  const marks: readonly ParagraphMark[] = [
    resolveParagraphMark(paragraph, context.styles),
    ...resolveRunMarks(paragraph, context.styles),
  ];

  let tallest = inlineHeightPt(paragraph);
  for (const mark of marks) {
    const height = heightOf(mark, context.metricsFor);
    if (height.kind === "blocked") return { kind: "blocked", mark };
    tallest = Math.max(tallest, height.value);
  }
  return { kind: "height", value: tallest };
}

function heightOf(mark: ParagraphMark, metricsFor: MetricsResolver): MarkHeight {
  if (mark.font.kind === "unresolved") return { kind: "blocked" };
  const lookup = metricsFor(faceRequestFor(mark));
  if (lookup.kind === "missing") return { kind: "blocked" };
  return { kind: "height", value: lineHeightPt(lookup.metrics, mark.fontSizePt) };
}

export const faceRequestFor = (mark: ParagraphMark): FaceRequest => ({
  name: mark.font.kind === "named" ? mark.font.name : "",
  bold: mark.bold,
  italic: mark.italic,
});

function blockerFor(mark: ParagraphMark, part: string, paragraphIndex: number): LayoutBlocker {
  if (mark.font.kind === "unresolved") return { kind: "unresolved-font", part, paragraphIndex };
  return {
    kind: "unknown-font-metrics",
    part,
    paragraphIndex,
    fontName: mark.font.name,
  };
}
