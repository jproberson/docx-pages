import type {
  Block,
  CellVerticalAlign,
  Paragraph,
  TableCell,
  TableInsets,
  TableRow,
} from "../docx/blocks.js";
import { numberParagraphs, type ParagraphNumber } from "../docx/list-numbers.js";
import type { NumberSuffix } from "../docx/numbering.js";
import { readRuns } from "../docx/runs.js";
import { W_NS } from "../docx/section.js";
import {
  resolveNumberMark,
  resolveParagraphFrame,
  resolveParagraphMark,
  resolveRunMarks,
  type ParagraphFrame,
  type ParagraphMark,
  type StyleTable,
} from "../docx/styles.js";
import { attribute, firstNamed } from "../docx/xml.js";
import { lineHeightPt } from "./font-metrics.js";
import {
  breakLines,
  faceRequestFor,
  justifyLine,
  measureText,
  type MeasureFailure,
  type MetricsResolver,
  type TextLine,
} from "./lines.js";
import { nextTabStopPt, tabStopsPt } from "./tab-stops.js";
import { twipsToPoints } from "./units.js";
import { fitLine, type WrapBand } from "./wrapping.js";

export const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

export type { MetricsResolver };

export type LayoutBlocker =
  | { readonly kind: "unresolved-font"; readonly part: string; readonly paragraphIndex: number }
  | {
      readonly kind: "unknown-font-metrics";
      readonly part: string;
      readonly paragraphIndex: number;
      readonly fontName: string;
    }
  | {
      readonly kind: "unmeasurable-text";
      readonly part: string;
      readonly paragraphIndex: number;
      readonly failure: MeasureFailure;
    }
  | {
      readonly kind: "unsupported-number-format";
      readonly part: string;
      readonly paragraphIndex: number;
      readonly numId: string;
      readonly ilvl: number;
    };

// The number a list puts in front of its paragraph, drawn out of the text flow at
// the hanging position the level indents to.
export type ParagraphMarker = {
  readonly text: string;
  readonly mark: ParagraphMark;
  readonly widthPt: number;
  readonly leftPt: number;
  readonly baselinePt: number;
};

export type PlacedLine = {
  readonly line: TextLine;
  readonly leftPt: number;
  readonly topPt: number;
  // What the line takes out of the stack, which is its own height plus whatever
  // the paragraph's line rule adds to it.
  readonly heightPt: number;
  readonly baselinePt: number;
};

export type ParagraphBox = {
  readonly index: number;
  readonly topPt: number;
  readonly heightPt: number;
  readonly lines: readonly PlacedLine[];
  readonly marker: ParagraphMarker | null;
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
  readonly leftPt: number;
  readonly widthPt: number;
  // A shape can refuse to wrap, in which case its text runs on past the frame
  // rather than breaking inside it.
  readonly wraps?: boolean;
  // The objects a paragraph anchors, in page coordinates, asked for as the stack
  // reaches that paragraph: every line from there on has to sit clear of them.
  readonly bandsFor?: BandResolver;
};

export type BandResolver = (paragraph: Paragraph, topPt: number) => readonly WrapBand[];

// The objects met so far, which grows as the stack walks forward.
type Region = { readonly bands: WrapBand[] };

type Context = Omit<MeasureStackInput, "blocks" | "originPt" | "leftPt" | "widthPt"> & {
  readonly numbers: ReadonlyMap<number, ParagraphNumber>;
  readonly region: Region;
};

type Frame = {
  readonly leftPt: number;
  readonly widthPt: number;
};

// Numbering is counted over the whole part before anything is measured, since a
// paragraph's number depends on every numbered paragraph ahead of it.
export function measureStack(input: MeasureStackInput): StackMeasurement {
  const numbered = numberParagraphs(input.blocks, input.styles);
  if (numbered.kind === "unsupported") {
    return {
      kind: "blocked",
      blocker: {
        kind: "unsupported-number-format",
        part: input.part,
        paragraphIndex: numbered.paragraphIndex,
        numId: numbered.numId,
        ilvl: numbered.ilvl,
      },
    };
  }

  const context: Context = { ...input, numbers: numbered.numbers, region: { bands: [] } };
  return measureBlocks(input.blocks, context, input.originPt, {
    leftPt: input.leftPt,
    widthPt: input.widthPt,
  });
}

function measureBlocks(
  blocks: readonly Block[],
  context: Context,
  originPt: number,
  frame: Frame,
): StackMeasurement {
  const boxes: ParagraphBox[] = [];
  let top = originPt;

  for (const block of blocks) {
    if (block.kind === "paragraph") {
      const measured = measureParagraph(block.paragraph, context, top, frame);
      if (measured.kind === "blocked") return measured;
      boxes.push(measured.box);
      top += measured.box.heightPt;
      continue;
    }

    for (const row of block.rows) {
      const measured = measureRow(row, context, top, frame, block.insets);
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
// cells sit beside each other, so their heights never add up. A cell holds its
// own content off its edges by the table's margins, which is what a paragraph
// inside one is indented from.
function measureRow(
  row: TableRow,
  context: Context,
  topPt: number,
  frame: Frame,
  insets: TableInsets,
): StackMeasurement {
  const measured: MeasuredCell[] = [];
  const leftMarginPt = twipsToPoints(insets.leftTwips);
  const rightMarginPt = twipsToPoints(insets.rightTwips);
  let heightPt = 0;
  let leftPt = frame.leftPt + twipsToPoints(insets.indentTwips);

  // A cell is measured from its own origin and only then moved down to the row, so
  // the page coordinates a wrapping object stands in cannot reach inside one.
  const inCell: Context = { ...context, region: { bands: [] }, bandsFor: () => [] };

  for (const cell of row.cells) {
    const widthPt = cellWidthPt(cell) ?? frame.widthPt;
    const cellFrame = {
      leftPt: leftPt + leftMarginPt,
      widthPt: Math.max(0, widthPt - leftMarginPt - rightMarginPt),
    };
    const of = measureBlocks(cell.blocks, inCell, 0, cellFrame);
    if (of.kind === "blocked") return of;
    measured.push({ align: cell.verticalAlign, boxes: of.boxes, heightPt: of.heightPt });
    heightPt = Math.max(heightPt, of.heightPt);
    leftPt += widthPt;
  }

  const boxes: ParagraphBox[] = [];
  for (const cell of measured) {
    const offset = topPt + seatingOffset(cell.align, heightPt, cell.heightPt);
    for (const box of cell.boxes) boxes.push(shiftBox(box, offset));
  }

  return { kind: "measured", boxes, heightPt };
}

function cellWidthPt(cell: TableCell): number | null {
  const properties = firstNamed(cell.element, W_NS, "tcPr");
  const width = properties === null ? null : firstNamed(properties, W_NS, "tcW");
  if (width === null || attribute(width, W_NS, "type") !== "dxa") return null;
  const twips = Number(attribute(width, W_NS, "w") ?? Number.NaN);
  return Number.isFinite(twips) ? twipsToPoints(twips) : null;
}

export const shiftBox = (box: ParagraphBox, byPt: number): ParagraphBox => ({
  ...box,
  topPt: box.topPt + byPt,
  lines: box.lines.map((line) => ({
    ...line,
    topPt: line.topPt + byPt,
    baselinePt: line.baselinePt + byPt,
  })),
  marker: box.marker === null ? null : { ...box.marker, baselinePt: box.marker.baselinePt + byPt },
});

export const shiftBoxes = (
  boxes: readonly ParagraphBox[],
  byPt: number,
): readonly ParagraphBox[] => (byPt === 0 ? boxes : boxes.map((box) => shiftBox(box, byPt)));

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

type ParagraphMeasurement =
  | { readonly kind: "measured"; readonly box: ParagraphBox }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

function measureParagraph(
  paragraph: Paragraph,
  context: Context,
  topPt: number,
  frame: Frame,
): ParagraphMeasurement {
  const marks: readonly ParagraphMark[] = [
    resolveParagraphMark(paragraph, context.styles),
    ...resolveRunMarks(paragraph, context.styles),
  ];

  let markHeight = 0;
  for (const mark of marks) {
    const height = heightOf(mark, context.metricsFor);
    if (height.kind === "blocked") {
      return {
        kind: "blocked",
        blocker: blockerFor(mark, context.part, paragraph.index),
      };
    }
    markHeight = Math.max(markHeight, height.value);
  }

  // An object is met where it is anchored, so a paragraph's own floats are already
  // standing there when its first line looks for room.
  if (context.bandsFor !== undefined) {
    context.region.bands.push(...context.bandsFor(paragraph, topPt));
  }

  const paragraphFrame = resolveParagraphFrame(paragraph, context.styles);
  const runs = readRuns(paragraph, context.styles);
  const insets = insetsOf(paragraphFrame);
  const number = context.numbers.get(paragraph.index);
  const widthPt =
    context.wraps === false
      ? Number.POSITIVE_INFINITY
      : frame.widthPt - insets.leftPt - insets.rightPt;

  // A hanging indent leaves its first line wider than the rest, except where a
  // number is what hangs there: then the text starts at the indent like the rest.
  const breaking = breakLines({
    runs,
    widthPt,
    firstLineWidthPt: number === undefined ? widthPt - insets.firstLinePt : widthPt,
    metricsFor: context.metricsFor,
    tabs: {
      stopsPt: tabStopsPt(paragraphFrame),
      originPt: insets.leftPt,
      firstLineOriginPt: number === undefined ? insets.leftPt + insets.firstLinePt : insets.leftPt,
    },
  });

  if (breaking.kind === "unmeasurable") {
    return {
      kind: "blocked",
      blocker: {
        kind: "unmeasurable-text",
        part: context.part,
        paragraphIndex: paragraph.index,
        failure: breaking.failure,
      },
    };
  }

  const measured =
    number === undefined ? null : measureNumber(paragraph, number, context, frame, paragraphFrame);
  if (measured !== null && measured.kind === "blocked") return measured;

  return {
    kind: "measured",
    box: layOutParagraph(paragraph.index, breaking.lines, {
      topPt,
      markHeightPt: markHeight,
      frame,
      paragraphFrame,
      number: measured === null ? null : measured.number,
      bands: context.region.bands,
    }),
  };
}

// A number sits at the hanging position and the text after it starts at whatever
// the level's suffix moves on to, which for a hanging paragraph is the implicit
// stop at its left indent.
type MeasuredNumber = {
  readonly text: string;
  readonly mark: ParagraphMark;
  readonly widthPt: number;
  readonly ascentPt: number;
  readonly leftPt: number;
  readonly textStartPt: number;
};

type NumberMeasurement =
  | { readonly kind: "measured"; readonly number: MeasuredNumber }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

function measureNumber(
  paragraph: Paragraph,
  number: ParagraphNumber,
  context: Context,
  frame: Frame,
  paragraphFrame: ParagraphFrame,
): NumberMeasurement {
  const mark = resolveNumberMark(paragraph, context.styles, number.level);
  const measured = measureText(number.text, mark, context.metricsFor);
  if (measured.kind === "unmeasurable") {
    return {
      kind: "blocked",
      blocker: {
        kind: "unmeasurable-text",
        part: context.part,
        paragraphIndex: paragraph.index,
        failure: measured.failure,
      },
    };
  }

  const insets = insetsOf(paragraphFrame);
  const leftPt = frame.leftPt + insets.leftPt + insets.firstLinePt;
  const endPt = leftPt + measured.widthPt;

  return {
    kind: "measured",
    number: {
      text: number.text,
      mark,
      widthPt: measured.widthPt,
      ascentPt: measured.ascentPt,
      leftPt,
      textStartPt: startOfText(endPt, number.level.suffix, {
        frame,
        paragraphFrame,
        spaceWidthPt: () => widthOfSpace(mark, context.metricsFor),
      }),
    },
  };
}

type SuffixContext = {
  readonly frame: Frame;
  readonly paragraphFrame: ParagraphFrame;
  readonly spaceWidthPt: () => number;
};

function startOfText(endPt: number, suffix: NumberSuffix, context: SuffixContext): number {
  const { frame, paragraphFrame } = context;
  if (suffix === "nothing") return endPt;
  if (suffix === "space") return endPt + context.spaceWidthPt();
  return frame.leftPt + nextTabStopPt(endPt - frame.leftPt, tabStopsPt(paragraphFrame));
}

// A face with no space of its own leaves the number against the text, which is
// still nearer than pretending to a width it does not have.
function widthOfSpace(mark: ParagraphMark, metricsFor: MetricsResolver): number {
  const measured = measureText(" ", mark, metricsFor);
  return measured.kind === "measured" ? measured.widthPt : 0;
}

type Insets = {
  readonly leftPt: number;
  readonly rightPt: number;
  readonly firstLinePt: number;
};

const insetsOf = (frame: ParagraphFrame): Insets => ({
  leftPt: twipsToPoints(frame.indentLeftTwips),
  rightPt: twipsToPoints(frame.indentRightTwips),
  firstLinePt: twipsToPoints(frame.indentFirstLineTwips),
});

type LayOutParagraphInput = {
  readonly topPt: number;
  readonly markHeightPt: number;
  readonly frame: Frame;
  readonly paragraphFrame: ParagraphFrame;
  readonly number: MeasuredNumber | null;
  readonly bands: readonly WrapBand[];
};

// A paragraph with text is as tall as the lines its runs measured to: Word does
// not let the paragraph mark raise a line it shares with a run, however much
// bigger the mark is. An empty paragraph is the mark's height alone.
function layOutParagraph(
  index: number,
  lines: readonly TextLine[],
  input: LayOutParagraphInput,
): ParagraphBox {
  const { paragraphFrame, frame, number } = input;
  const insets = insetsOf(paragraphFrame);
  const beforePt = twipsToPoints(paragraphFrame.spaceBeforeTwips);
  const afterPt = twipsToPoints(paragraphFrame.spaceAfterTwips);

  // An empty paragraph is a line like any other as far as objects are concerned:
  // Word moves it out of their way even though it draws nothing there.
  if (lines.length === 0) {
    const slot = fitLine({
      topPt: input.topPt + beforePt,
      heightPt: input.markHeightPt,
      leftPt: frame.leftPt + insets.leftPt,
      rightPt: frame.leftPt + frame.widthPt - insets.rightPt,
      widthPt: 0,
      bands: input.bands,
    });

    return {
      index,
      topPt: input.topPt,
      heightPt: slot.topPt + input.markHeightPt + afterPt - input.topPt,
      lines: [],
      marker: markerAt(number, slot.topPt + (number?.ascentPt ?? 0)),
    };
  }

  const placed: PlacedLine[] = [];
  let top = input.topPt + beforePt;

  lines.forEach((line, at) => {
    // A line a tab alone holds open has nothing measured on it to give it a
    // height, so it takes the tallest mark the paragraph has, as an empty
    // paragraph does.
    const ownPt =
      line.segments.length === 0 ? Math.max(line.heightPt, input.markHeightPt) : line.heightPt;
    const raisedPt = at === 0 ? liftOfNumber(number, line) : 0;
    const heightPt = spacedHeightPt(ownPt + raisedPt, paragraphFrame);
    const firstLinePt = at === 0 ? insets.firstLinePt : 0;
    // The number takes the first line's own start, so the text after it begins
    // wherever the number's suffix moved on to.
    const startPt =
      at === 0 && number !== null ? number.textStartPt : frame.leftPt + insets.leftPt + firstLinePt;
    const endPt = frame.leftPt + frame.widthPt - insets.rightPt;

    const slot = fitLine({
      topPt: top,
      heightPt,
      leftPt: startPt,
      rightPt: endPt,
      widthPt: line.widthPt,
      bands: input.bands,
    });

    // A justified line fills the room it was fitted into, which an object beside
    // it may have narrowed. Only the paragraph's last line is left as it fell; a
    // line that ended at a manual break is stretched like any other.
    const filled =
      paragraphFrame.alignment === "justify" && at < lines.length - 1
        ? justifyLine(line, slot.rightPt - slot.leftPt)
        : line;

    placed.push({
      line: filled,
      leftPt: lineStartPt(paragraphFrame, slot.leftPt, slot.rightPt, filled.widthPt),
      topPt: slot.topPt,
      heightPt,
      baselinePt: slot.topPt + filled.ascentPt + raisedPt,
    });
    top = slot.topPt + heightPt;
  });

  return {
    index,
    topPt: input.topPt,
    heightPt: top + afterPt - input.topPt,
    lines: placed,
    marker: markerAt(number, placed[0]?.baselinePt ?? input.topPt),
  };
}

// A number lifts the top of the line it sits on by however much its own ascent
// reaches above the line's, and never reaches below the baseline: a Symbol bullet
// raises a 12pt Calibri line from 14.65pt to 15.29pt, which is its ascent over
// that text's descent, while a Courier one with a deeper descent leaves the line
// alone.
const liftOfNumber = (number: MeasuredNumber | null, line: TextLine): number =>
  number === null ? 0 : Math.max(0, number.ascentPt - line.ascentPt);

const markerAt = (number: MeasuredNumber | null, baselinePt: number): ParagraphMarker | null =>
  number === null
    ? null
    : {
        text: number.text,
        mark: number.mark,
        widthPt: number.widthPt,
        leftPt: number.leftPt,
        baselinePt,
      };

function spacedHeightPt(naturalPt: number, frame: ParagraphFrame): number {
  const { lineTwips, lineRule } = frame;
  if (lineTwips === null) return naturalPt;
  if (lineRule === "exact") return twipsToPoints(lineTwips);
  if (lineRule === "atLeast") return Math.max(naturalPt, twipsToPoints(lineTwips));
  return (naturalPt * lineTwips) / LINE_MULTIPLE_UNITS;
}

// w:line counts 240ths of a line when the rule is "auto".
const LINE_MULTIPLE_UNITS = 240;

function lineStartPt(
  frame: ParagraphFrame,
  startPt: number,
  endPt: number,
  contentPt: number,
): number {
  if (frame.alignment === "right") return endPt - contentPt;
  if (frame.alignment === "center") return startPt + (endPt - startPt - contentPt) / 2;
  return startPt;
}

type MarkHeight =
  { readonly kind: "height"; readonly value: number } | { readonly kind: "blocked" };

function heightOf(mark: ParagraphMark, metricsFor: MetricsResolver): MarkHeight {
  if (mark.font.kind === "unresolved") return { kind: "blocked" };
  const lookup = metricsFor(faceRequestFor(mark));
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
