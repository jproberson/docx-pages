import { paragraphOwnDrawings, type Paragraph } from "../docx/paragraphs.js";
import {
  resolveParagraphMark,
  resolveRunMarks,
  type ParagraphMark,
  type StyleTable,
} from "../docx/styles.js";
import { attribute } from "../docx/xml.js";
import { lineHeightPt, type MetricsLookup } from "./font-metrics.js";
import { EMU_PER_POINT } from "./units.js";

export const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

export type MetricsResolver = (fontName: string) => MetricsLookup;

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
  readonly paragraphs: readonly Paragraph[];
  readonly styles: StyleTable;
  readonly metricsFor: MetricsResolver;
  readonly part: string;
  readonly originPt: number;
};

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

export function measureStack(input: MeasureStackInput): StackMeasurement {
  const boxes: ParagraphBox[] = [];
  let top = input.originPt;

  for (const paragraph of input.paragraphs) {
    const marks: readonly ParagraphMark[] = [
      resolveParagraphMark(paragraph, input.styles),
      ...resolveRunMarks(paragraph, input.styles),
    ];

    let tallest = inlineHeightPt(paragraph);
    for (const mark of marks) {
      const height = heightOf(mark, input.metricsFor);
      if (height.kind === "blocked") {
        return {
          kind: "blocked",
          blocker: blockerFor(mark, input.part, paragraph.index),
        };
      }
      tallest = Math.max(tallest, height.value);
    }

    boxes.push({ index: paragraph.index, topPt: top, heightPt: tallest });
    top += tallest;
  }

  return { kind: "measured", boxes, heightPt: top - input.originPt };
}

function heightOf(mark: ParagraphMark, metricsFor: MetricsResolver): MarkHeight {
  if (mark.font.kind === "unresolved") return { kind: "blocked" };
  const lookup = metricsFor(mark.font.name);
  if (lookup.kind === "missing") return { kind: "blocked" };
  return { kind: "height", value: lineHeightPt(lookup.metrics, mark.fontSizePt) };
}

function blockerFor(mark: ParagraphMark, part: string, paragraphIndex: number): LayoutBlocker {
  if (mark.font.kind === "unresolved") return { kind: "unresolved-font", part, paragraphIndex };
  return {
    kind: "unknown-font-metrics",
    part,
    paragraphIndex,
    fontName: mark.font.name,
  };
}
