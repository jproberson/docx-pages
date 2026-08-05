import type {
  Block,
  CellVerticalAlign,
  Paragraph,
  TableCell,
  TableInsets,
  TableRow,
} from "../docx/blocks.js";
import {
  borderExtentPt,
  resolveCellBorders,
  NO_BORDERS,
  SIDES,
  type Border,
  type Borders,
} from "../docx/borders.js";
import { numberParagraphs, type ParagraphNumber } from "../docx/list-numbers.js";
import type { NumberSuffix } from "../docx/numbering.js";
import { readRuns, type TextRun } from "../docx/runs.js";
import { W_NS } from "../docx/section.js";
import { roundsAnchorsToTwips, DEFAULT_SETTINGS, type DocumentSettings } from "../docx/settings.js";
import {
  mergeTableBorders,
  resolveNumberMark,
  resolveParagraphFrame,
  resolveParagraphMark,
  resolveRunMarks,
  resolveTableBorders,
  styleIdOf,
  type ParagraphFrame,
  type ParagraphMark,
  type StyleTable,
} from "../docx/styles.js";
import { attribute, firstNamed } from "../docx/xml.js";
import { lineHeightPt } from "./font-metrics.js";
import {
  beginLines,
  faceRequestFor,
  justifyLine,
  measureText,
  type LineFlow,
  type MeasureFailure,
  type MetricsResolver,
  type TextLine,
} from "./lines.js";
import { nextTabStop, tabStopsPt } from "./tab-stops.js";
import { twipsToPoints } from "./units.js";
import { fitLine, type LineSlot, type WrapBand } from "./wrapping.js";

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
  // How far down that room the line of text itself starts, which is nothing until
  // a rule opens room above it. Word answers for a paragraph from here.
  readonly seatPt: number;
  readonly baselinePt: number;
  // Whether a page break in the paragraph's own text put this line at the head of
  // a page. Only the break itself can act on it.
  readonly startsPage: boolean;
};

// The rectangle a paragraph's text is cut to. A row told exactly how tall to be
// gives one to every cell in it, since Word draws what does not fit and then cuts
// it off at the row: text seated below an exact row's foot is written into the
// pdf and painted nowhere.
export type ClipRect = {
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

// The lines drawn round a paragraph and the colour drawn behind it. How far up
// and down they reach is the paragraph's own lines; how far out to either side is
// here, since that is the text area rather than anything the text did with it.
//
// A border a neighbour asks for in the same words is left out: Word joins a run
// of such paragraphs into one box rather than drawing a line between each pair.
export type ParagraphPaint = {
  readonly leftPt: number;
  readonly rightPt: number;
  readonly fillColor: string | null;
  readonly borders: Borders;
};

export type ParagraphBox = {
  readonly index: number;
  readonly topPt: number;
  // Where the objects this paragraph anchors are measured from, which is its own
  // top until a legacy document's rounding drops the paragraph past one of them:
  // the object keeps the place the flow first gave it.
  readonly anchorTopPt: number;
  readonly heightPt: number;
  readonly lines: readonly PlacedLine[];
  readonly marker: ParagraphMarker | null;
  // How far the paragraph reaches across the frame it was laid out in, measured
  // from that frame's own left. A box fitting itself to its text is as wide as the
  // widest paragraph in it, and an empty paragraph still reaches as far as its own
  // mark.
  readonly contentWidthPt: number;
  // Where the paragraph mark itself came to rest, which is the paragraph's own top
  // until an object moves the mark's line down or a line rule seats it lower in
  // one. A paragraph with no text draws nothing there and still holds the room,
  // and Word answers for it.
  readonly markTopPt: number;
  // The foot of the last thing the paragraph draws: its last line, or the room its
  // mark stands in where it has none. A page break is decided by this rather than
  // by the paragraph's whole height, since the room a paragraph keeps below itself
  // never holds it back at the foot of a page.
  readonly contentBottomPt: number;
  // What the paragraph asks of a page break running through it, which only the
  // break itself can act on.
  readonly widowControl: boolean;
  // Whether the paragraph asked for a page of its own, and whether it ended on a
  // break that puts whatever follows it on one.
  readonly startsPage: boolean;
  readonly endsPage: boolean;
  // What the paragraph's text is cut off at, which is the row when a row was told
  // exactly how tall to be and nothing anywhere else.
  readonly clipTo: ClipRect | null;
  readonly paint: ParagraphPaint | null;
};

// A cell as it stands on the page: the rectangle its grid lines mark out, the
// colour behind its text, and the line drawn along each of its edges. Word
// centres a border on the edge it runs along, so half of one falls outside this
// rectangle, and a fill stops at the inner side of each.
//
// This is a peer of a paragraph rather than something hanging off one: a cell
// holds paragraphs, and drawing one is not drawing any of them.
export type PlacedCell = {
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly fillColor: string | null;
  readonly borders: Borders;
};

export type StackMeasurement =
  | {
      readonly kind: "measured";
      readonly boxes: readonly ParagraphBox[];
      readonly cells: readonly PlacedCell[];
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
  readonly settings?: DocumentSettings;
};

export type BandResolver = (paragraph: Paragraph, topPt: number) => readonly WrapBand[];

type Context = Omit<MeasureStackInput, "blocks" | "originPt" | "leftPt" | "widthPt"> & {
  readonly numbers: ReadonlyMap<number, ParagraphNumber>;
  readonly settings: DocumentSettings;
  // A page break inside a cell is no break at all, and a cell is the only place
  // that has to know it.
  readonly inCell: boolean;
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

  const context: Context = {
    ...input,
    numbers: numbered.numbers,
    settings: input.settings ?? DEFAULT_SETTINGS,
    inCell: false,
  };
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
  const cells: PlacedCell[] = [];
  let top = originPt;
  // The objects met so far, which grows as the stack walks forward: every line
  // from an object's own paragraph on has to sit clear of it.
  let standing: readonly WrapBand[] = [];
  // Where the paragraph coming up anchors its objects, which is its own top until
  // the one above it was dropped past one of them.
  let anchoredAtPt: number | null = null;

  for (const [at, block] of blocks.entries()) {
    if (block.kind === "paragraph") {
      const { paragraph } = block;
      const neighbours = {
        above: paragraphAt(blocks, at - 1),
        below: paragraphAt(blocks, at + 1),
      };
      const anchorTopPt = anchoredAtPt ?? top;
      const bands = [...standing, ...bandsOf(paragraph, anchorTopPt, context)];
      const measured = measureParagraph(paragraph, context, top, frame, neighbours, {
        bands,
        ahead: [],
        anchorTopPt,
      });
      if (measured.kind === "blocked") return measured;

      standing = bands;
      anchoredAtPt = null;
      let box = measured.box;

      // Where the anchor rounds down, the object the next paragraph holds stands
      // over the foot of this one, whose last line is then blocked and falls past
      // it. The object itself keeps the place the flow gave it, so the room the
      // line had is left empty.
      const ahead = lookedAhead(neighbours.below, top + box.heightPt, context);
      if (ahead.length > 0) {
        const again = measureParagraph(paragraph, context, top, frame, neighbours, {
          bands,
          ahead,
          anchorTopPt,
        });
        if (again.kind === "blocked") return again;
        anchoredAtPt = top + box.heightPt;
        box = again.box;
      }

      boxes.push(box);
      top += box.heightPt;
      continue;
    }

    const measured = measureTable(block, context, top, frame);
    if (measured.kind === "blocked") return measured;
    boxes.push(...measured.boxes);
    cells.push(...measured.cells);
    anchoredAtPt = null;
    top += measured.heightPt;
  }

  return { kind: "measured", boxes, cells, heightPt: top - originPt };
}

const bandsOf = (paragraph: Paragraph, topPt: number, context: Context): readonly WrapBand[] =>
  context.bandsFor === undefined ? [] : context.bandsFor(paragraph, topPt);

// Half the twip a legacy document's anchors are rounded to, which is as far over
// the paragraph above an object can come to stand by that rounding alone. An
// object reaching further up than that is one nothing has measured, and is left to
// the paragraph that anchors it as it always was.
const ROUNDING_PT = 1 / 40;

// The objects the next paragraph anchors that ended up standing over the foot of
// this one, which only a document Word rounds an anchor's position in can have.
function lookedAhead(
  below: Paragraph | null,
  footPt: number,
  context: Context,
): readonly WrapBand[] {
  if (below === null || !roundsAnchorsToTwips(context.settings)) return [];
  return bandsOf(below, footPt, context).filter(
    (band) => band.topPt < footPt - EPSILON && footPt - band.topPt <= ROUNDING_PT + EPSILON,
  );
}

type Table = Extract<Block, { kind: "table" }>;

// Half of a border falls outside the line it is centred on, so half of the ones
// round the outside of a table falls outside the table. What stands where the
// table goes is that outer edge: its first grid line is half its widest top
// border below the flow, and half its widest left border in from its indent.
function measureTable(
  block: Table,
  context: Context,
  topPt: number,
  frame: Frame,
): StackMeasurement {
  const borders = resolveCellBorders(
    block.rows.map((row) => row.cells.map((cell) => cell.borders)),
    mergeTableBorders(resolveTableBorders(context.styles, block.styleId), block.borders),
  );

  const first = borders[0] ?? [];
  const last = borders[borders.length - 1] ?? [];
  const outerTopPt = halfOf(first.map((cell) => cell.top));
  const outerBottomPt = halfOf(last.map((cell) => cell.bottom));
  const outerLeftPt = halfOf(borders.map((row) => row[0]?.left ?? null));

  const boxes: ParagraphBox[] = [];
  const cells: PlacedCell[] = [];
  const rowFrame = {
    leftPt: frame.leftPt + twipsToPoints(block.insets.indentTwips) + outerLeftPt,
    widthPt: frame.widthPt,
  };

  let top = topPt + outerTopPt;
  for (const [at, row] of block.rows.entries()) {
    const measured = measureRow(row, borders[at] ?? [], context, top, rowFrame, block.insets);
    if (measured.kind === "blocked") return measured;
    boxes.push(...measured.boxes);
    cells.push(...measured.cells);
    top += measured.heightPt;
  }

  return { kind: "measured", boxes, cells, heightPt: top + outerBottomPt - topPt };
}

// How far a line drawn along an edge reaches to either side of it, which is the
// room the cells on both sides of it have to leave.
const halfOf = (borders: readonly (Border | null)[]): number =>
  Math.max(0, ...borders.map((border) => borderExtentPt(border) / 2));

// A paragraph in the next cell or on the other side of a table is not a
// neighbour: only what stands beside it in its own run of blocks is.
function paragraphAt(blocks: readonly Block[], at: number): Paragraph | null {
  const block = blocks[at];
  return block !== undefined && block.kind === "paragraph" ? block.paragraph : null;
}

type MeasuredCell = {
  readonly align: CellVerticalAlign;
  readonly boxes: readonly ParagraphBox[];
  // The cells of a table inside this one, which move with its content rather
  // than with the row.
  readonly inner: readonly PlacedCell[];
  readonly heightPt: number;
  readonly leftPt: number;
  readonly widthPt: number;
  readonly fillColor: string | null;
  readonly borders: Borders;
};

// A row is as tall as its tallest cell, and every cell starts at the row's top;
// cells sit beside each other, so their heights never add up.
//
// How far a cell holds its content off its walls is not asked the same way at
// every side. Left and right are the cell's own business: a cell states its
// margins and its neighbour, which states none, keeps the table's. Above and
// below they are the row's: the largest top margin any cell in the row asks for
// holds every cell in it off the top wall, and the largest bottom margin adds to
// the row under all of them.
//
// A border is not room on top of that but room of its own: the text clears the
// half of the line that falls inside the cell, and where the margin already holds
// it further off than that the border asks for nothing. Case h's rows, held off
// their walls by 2.75pt and lined with a quarter point, come out the height they
// came out with no borders read at all; the same rows with a 6pt line and no
// margin clear the whole 3pt of it.
function measureRow(
  row: TableRow,
  borders: readonly Borders[],
  context: Context,
  topPt: number,
  frame: Frame,
  insets: TableInsets,
): StackMeasurement {
  const measured: MeasuredCell[] = [];
  const topMarginPt = Math.max(
    rowMarginPt(row, insets, "topTwips"),
    halfOf(borders.map((of) => of.top)),
  );
  const bottomMarginPt = Math.max(
    rowMarginPt(row, insets, "bottomTwips"),
    halfOf(borders.map((of) => of.bottom)),
  );
  let contentHeightPt = 0;
  let leftPt = frame.leftPt;

  // A cell is measured from its own origin and only then moved down to the row, so
  // the page coordinates a wrapping object stands in cannot reach inside one.
  const inCell: Context = { ...context, bandsFor: () => [], inCell: true };

  for (const [at, cell] of row.cells.entries()) {
    const widthPt = cellWidthPt(cell) ?? frame.widthPt;
    const own = borders[at] ?? NO_BORDERS;
    const leftMarginPt = Math.max(
      twipsToPoints(cell.margins.leftTwips ?? insets.leftTwips),
      halfOf([own.left]),
    );
    const rightMarginPt = Math.max(
      twipsToPoints(cell.margins.rightTwips ?? insets.rightTwips),
      halfOf([own.right]),
    );
    const cellFrame = {
      leftPt: leftPt + leftMarginPt,
      widthPt: Math.max(0, widthPt - leftMarginPt - rightMarginPt),
    };
    const of = measureBlocks(cell.blocks, inCell, 0, cellFrame);
    if (of.kind === "blocked") return of;
    measured.push({
      align: cell.verticalAlign,
      boxes: of.boxes,
      inner: of.cells,
      heightPt: of.heightPt,
      leftPt,
      widthPt,
      fillColor: cell.fillColor,
      borders: own,
    });
    contentHeightPt = Math.max(contentHeightPt, of.heightPt);
    leftPt += widthPt;
  }

  const heightPt = rowHeightPt(row, topMarginPt + contentHeightPt + bottomMarginPt);
  // A row told exactly how tall to be leaves its cells whatever room is left over
  // once it has held them off its walls, and Word draws what does not fit anyway.
  const roomPt = Math.max(0, heightPt - topMarginPt - bottomMarginPt);

  const boxes: ParagraphBox[] = [];
  const cells: PlacedCell[] = [];
  for (const cell of measured) {
    const offset = topPt + topMarginPt + seatingOffset(cell.align, roomPt, cell.heightPt);
    // Only a row given a height of its own can be shorter than what it holds, so
    // only that row has anything to cut its cells off at.
    const clipTo =
      row.height?.exact === true
        ? { leftPt: cell.leftPt, topPt, widthPt: cell.widthPt, heightPt }
        : null;
    for (const box of cell.boxes) boxes.push({ ...shiftBox(box, offset), clipTo });
    for (const inner of cell.inner) cells.push({ ...inner, topPt: inner.topPt + offset });
    cells.push({
      leftPt: cell.leftPt,
      topPt,
      widthPt: cell.widthPt,
      heightPt,
      fillColor: cell.fillColor,
      borders: cell.borders,
    });
  }

  return { kind: "measured", boxes, cells, heightPt };
}

// The largest margin any cell in the row asks for at that side, which is what
// every cell in it is held off the wall by.
function rowMarginPt(row: TableRow, insets: TableInsets, side: "topTwips" | "bottomTwips"): number {
  const twips = row.cells.map((cell) => cell.margins[side] ?? insets[side]);
  return twipsToPoints(Math.max(insets[side], ...twips));
}

// A stated height is a floor under the row until the row says it is exact, and
// then it is the whole of the row however much its cells hold.
function rowHeightPt(row: TableRow, heldPt: number): number {
  if (row.height === null) return heldPt;
  const askedPt = twipsToPoints(row.height.twips);
  return row.height.exact ? askedPt : Math.max(heldPt, askedPt);
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
  anchorTopPt: box.anchorTopPt + byPt,
  markTopPt: box.markTopPt + byPt,
  contentBottomPt: box.contentBottomPt + byPt,
  clipTo: box.clipTo === null ? null : { ...box.clipTo, topPt: box.clipTo.topPt + byPt },
  lines: box.lines.map((line) => ({
    ...line,
    topPt: line.topPt + byPt,
    baselinePt: line.baselinePt + byPt,
  })),
  marker: box.marker === null ? null : { ...box.marker, baselinePt: box.marker.baselinePt + byPt },
});

export const shiftCells = (cells: readonly PlacedCell[], byPt: number): readonly PlacedCell[] =>
  byPt === 0 ? cells : cells.map((cell) => ({ ...cell, topPt: cell.topPt + byPt }));

export const shiftBoxes = (
  boxes: readonly ParagraphBox[],
  byPt: number,
): readonly ParagraphBox[] => (byPt === 0 ? boxes : boxes.map((box) => shiftBox(box, byPt)));

function seatingOffset(align: CellVerticalAlign, roomPt: number, cellHeightPt: number): number {
  const slack = Math.max(0, roomPt - cellHeightPt);
  if (align === "center") return slack / 2;
  if (align === "bottom") return slack;
  return 0;
}

type ParagraphMeasurement =
  | { readonly kind: "measured"; readonly box: ParagraphBox }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

// What stands either side of the paragraph in the same run of blocks, which is
// all "don't add space between paragraphs of the same style" asks about.
type Neighbours = {
  readonly above: Paragraph | null;
  readonly below: Paragraph | null;
};

// What the objects around the paragraph have already settled: the ones standing in
// its way, and where its own were anchored.
type Standing = {
  readonly bands: readonly WrapBand[];
  // Objects the paragraph after this one anchors that came to stand over its foot,
  // which its last line has to make room for.
  readonly ahead: readonly WrapBand[];
  readonly anchorTopPt: number;
};

function measureParagraph(
  paragraph: Paragraph,
  context: Context,
  topPt: number,
  frame: Frame,
  neighbours: Neighbours,
  standing: Standing,
): ParagraphMeasurement {
  const paragraphMark = resolveParagraphMark(paragraph, context.styles);
  const marks: readonly ParagraphMark[] = [
    paragraphMark,
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

  const paragraphFrame = resolveParagraphFrame(paragraph, context.styles);
  const runs = flowing(readRuns(paragraph, context.styles), context.inCell);
  const insets = insetsOf(paragraphFrame);
  const number = context.numbers.get(paragraph.index);
  const widthPt =
    context.wraps === false
      ? Number.POSITIVE_INFINITY
      : frame.widthPt - insets.leftPt - insets.rightPt;

  const breaking = beginLines({
    runs,
    metricsFor: context.metricsFor,
    tabs: {
      stopsPt: tabStopsPt(paragraphFrame),
      originPt: insets.leftPt,
      firstLineOriginPt: number === undefined ? insets.leftPt + insets.firstLinePt : insets.leftPt,
      defaultStopPt: twipsToPoints(context.settings.defaultTabStopTwips),
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
    box: layOutParagraph(paragraph.index, breaking.flow, {
      topPt,
      anchorTopPt: standing.anchorTopPt,
      markHeightPt: markHeight,
      markWidthPt: widthOfMark(paragraphMark, context.metricsFor),
      frame,
      paragraphFrame,
      spacing: spacingPt(paragraph, paragraphFrame, context, neighbours),
      paint: paintOf(paragraphFrame, context, neighbours, {
        leftPt: frame.leftPt + insets.leftPt,
        rightPt: frame.leftPt + frame.widthPt - insets.rightPt,
      }),
      startsPage: !context.inCell && paragraphFrame.pageBreakBefore,
      number: measured === null ? null : measured.number,
      bands: standing.bands,
      ahead: standing.ahead,
      roomPt: widthPt,
      // A hanging indent leaves its first line wider than the rest, except where a
      // number is what hangs there: then the text starts at the indent like the rest.
      firstLineRoomPt: number === undefined ? widthPt - insets.firstLinePt : widthPt,
    }),
  };
}

// Word ignores a page break inside a cell outright: the text either side of one
// comes out on the same line and the row stands where it always did. Dropping the
// piece is what says so, since a break left in place would still end its line.
function flowing(runs: readonly TextRun[], inCell: boolean): readonly TextRun[] {
  if (!inCell) return runs;
  return runs.map((run) => ({
    ...run,
    pieces: run.pieces.filter((piece) => piece.kind !== "break" || !piece.endsPage),
  }));
}

// What a paragraph has drawn round it, once the neighbours have had their say.
// A run of paragraphs asking for the same border is one box in Word: the line
// between two of them is not drawn, and no room is left for it either.
function paintOf(
  paragraphFrame: ParagraphFrame,
  context: Context,
  neighbours: Neighbours,
  across: { readonly leftPt: number; readonly rightPt: number },
): ParagraphPaint | null {
  const { borders, fillColor } = paragraphFrame;
  if (fillColor === null && SIDES.every((side) => borders[side] === null)) return null;

  const joins = (other: Paragraph | null): boolean =>
    other !== null && sameBorders(borders, resolveParagraphFrame(other, context.styles).borders);

  return {
    ...across,
    fillColor,
    borders: {
      ...borders,
      top: joins(neighbours.above) ? null : borders.top,
      bottom: joins(neighbours.below) ? null : borders.bottom,
    },
  };
}

const sameBorders = (one: Borders, other: Borders): boolean =>
  SIDES.every((side) => sameBorder(one[side], other[side]));

const sameBorder = (one: Border | null, other: Border | null): boolean =>
  one === null || other === null
    ? one === other
    : one.style === other.style &&
      one.widthPt === other.widthPt &&
      one.color === other.color &&
      one.spacePt === other.spacePt;

// The room a line drawn round a paragraph takes from the flow: the line itself and
// the distance it stands off the text.
const borderRoomPt = (border: Border | null): number =>
  border === null ? 0 : borderExtentPt(border) + border.spacePt;

// The room a paragraph keeps above and below itself. `w:contextualSpacing` drops
// whichever of the two faces a paragraph of the same style, which is how a list
// closes up into one block while still standing off the text around it.
type Spacing = {
  readonly beforePt: number;
  readonly afterPt: number;
};

// Where one paragraph's room below meets the next one's room above, Word keeps the
// larger of the two rather than both: a paragraph asking for 11.25pt under it
// followed by one asking for 12pt over it leaves 12pt between them, not 23.25.
// The room below belongs to the paragraph above, so this is taken off the room
// above rather than added anywhere.
function spacingPt(
  paragraph: Paragraph,
  paragraphFrame: ParagraphFrame,
  context: Context,
  neighbours: Neighbours,
): Spacing {
  const own = ownSpacingPt(paragraph, paragraphFrame, context, neighbours);
  const abovePt = roomBelowPt(neighbours.above, paragraph, context);
  return { ...own, beforePt: Math.max(0, own.beforePt - abovePt) };
}

function ownSpacingPt(
  paragraph: Paragraph,
  paragraphFrame: ParagraphFrame,
  context: Context,
  neighbours: Neighbours,
): Spacing {
  const beforePt = twipsToPoints(paragraphFrame.spaceBeforeTwips);
  const afterPt = twipsToPoints(paragraphFrame.spaceAfterTwips);
  if (!paragraphFrame.contextualSpacing) return { beforePt, afterPt };

  const own = styleIdOf(paragraph, context.styles);
  const sameStyle = (other: Paragraph | null): boolean =>
    other !== null && styleIdOf(other, context.styles) === own;

  return {
    beforePt: sameStyle(neighbours.above) ? 0 : beforePt,
    afterPt: sameStyle(neighbours.below) ? 0 : afterPt,
  };
}

// How much room the paragraph above keeps under itself, which is the whole of
// what it already put between the two.
function roomBelowPt(above: Paragraph | null, below: Paragraph, context: Context): number {
  if (above === null) return 0;
  const frame = resolveParagraphFrame(above, context.styles);
  return ownSpacingPt(above, frame, context, { above: null, below }).afterPt;
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
  // The number's own tab starts the text at the stop whatever that stop lines
  // other text up on: it is the number that is being placed, not what follows.
  return frame.leftPt + nextTabStop(endPt - frame.leftPt, tabStopsPt(paragraphFrame)).positionPt;
}

// A face with no space of its own leaves the number against the text, which is
// still nearer than pretending to a width it does not have.
function widthOfSpace(mark: ParagraphMark, metricsFor: MetricsResolver): number {
  const measured = measureText(" ", mark, metricsFor);
  return measured.kind === "measured" ? measured.widthPt : 0;
}

// The mark itself, which is the only thing an empty paragraph holds. Measured
// against Word by shrinking a box that fits itself to an empty paragraph: it came
// out one pilcrow wide in every one of Calibri, Arial, Impact, Times New Roman,
// Verdana and Courier New, each within the twentieth of a point Word rounds its
// answers to. A face that cannot measure the character leaves the paragraph no
// width, which is what a box with nothing in it was given before.
const PILCROW = "¶";

function widthOfMark(mark: ParagraphMark, metricsFor: MetricsResolver): number {
  const measured = measureText(PILCROW, mark, metricsFor);
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
  readonly anchorTopPt: number;
  readonly markHeightPt: number;
  readonly markWidthPt: number;
  readonly frame: Frame;
  readonly paragraphFrame: ParagraphFrame;
  readonly spacing: Spacing;
  readonly paint: ParagraphPaint | null;
  readonly number: MeasuredNumber | null;
  // Whether the paragraph asked for a page of its own.
  readonly startsPage: boolean;
  readonly bands: readonly WrapBand[];
  readonly ahead: readonly WrapBand[];
  // What a line has room for where nothing stands beside it, which a shape that
  // refuses to wrap leaves unbounded.
  readonly roomPt: number;
  readonly firstLineRoomPt: number;
};

// A paragraph with text is as tall as the lines its runs measured to: Word does
// not let the paragraph mark raise a line it shares with a run, however much
// bigger the mark is. An empty paragraph is the mark's height alone.
function layOutParagraph(index: number, flow: LineFlow, input: LayOutParagraphInput): ParagraphBox {
  const across = acrossOf(input);
  return droppedPast(layOutWholeParagraph(index, flow, input), input.ahead, across);
}

const acrossOf = (input: LayOutParagraphInput): Span => {
  const insets = insetsOf(input.paragraphFrame);
  return {
    leftPt: input.frame.leftPt + insets.leftPt,
    rightPt: input.frame.leftPt + input.frame.widthPt - insets.rightPt,
  };
};

type Span = { readonly leftPt: number; readonly rightPt: number };

// What the paragraph draws last: its last line, or the room its mark stands in
// where it has none, taken down to the paragraph's own foot. That is the box an
// object standing over the foot is asked to make room for, and Word answers it as
// it answers any blocked line: measured over paragraphs of one line, of three and
// of none at all, only the last of them falls, it lands on the object's foot, and
// the room the paragraph keeps below itself goes with it. Beside an object narrow
// enough to leave the line somewhere to sit, nothing moves.
function droppedPast(box: ParagraphBox, ahead: readonly WrapBand[], across: Span): ParagraphBox {
  if (ahead.length === 0) return box;

  const last = box.lines[box.lines.length - 1];
  const topPt = last === undefined ? box.markTopPt : last.topPt;
  const slot = fitLine({
    topPt,
    heightPt: box.topPt + box.heightPt - topPt,
    ...across,
    widthPt: last?.line.widthPt ?? 0,
    bands: ahead,
  });

  const byPt = slot.topPt - topPt;
  if (byPt <= EPSILON) return box;
  return {
    ...box,
    heightPt: box.heightPt + byPt,
    contentBottomPt: box.contentBottomPt + byPt,
    markTopPt: box.markTopPt + byPt,
    lines:
      last === undefined
        ? box.lines
        : [
            ...box.lines.slice(0, -1),
            { ...last, topPt: last.topPt + byPt, baselinePt: last.baselinePt + byPt },
          ],
    // A number stands on the paragraph's first line, and moves only when that is
    // the line that fell.
    marker:
      box.marker === null || box.lines.length > 1
        ? box.marker
        : { ...box.marker, baselinePt: box.marker.baselinePt + byPt },
  };
}

function layOutWholeParagraph(
  index: number,
  flow: LineFlow,
  input: LayOutParagraphInput,
): ParagraphBox {
  const { paragraphFrame, frame, number, paint } = input;
  const insets = insetsOf(paragraphFrame);
  const { beforePt, afterPt } = input.spacing;
  // A border round the paragraph is room it takes out of the flow, above the
  // first line and below the last.
  const abovePt = borderRoomPt(paint?.borders.top ?? null);
  const belowPt = borderRoomPt(paint?.borders.bottom ?? null);
  const { lines: laid, endsPage } = layOutLines(flow, input);

  // An empty paragraph is a line like any other as far as objects are concerned:
  // Word moves it out of their way even though it draws nothing there.
  if (laid.length === 0) {
    // A paragraph with nothing in it answers to its line rule as any other does,
    // its mark seated in the room that rule leaves.
    const height = seatedHeight(
      {
        naturalPt: input.markHeightPt,
        ascentPt: number?.ascentPt ?? 0,
        seatPt: 0,
        fontHeightPt: input.markHeightPt,
      },
      paragraphFrame,
    );
    const slot = fitLine({
      topPt: input.topPt + beforePt + abovePt,
      heightPt: height.heightPt,
      leftPt: frame.leftPt + insets.leftPt,
      rightPt: frame.leftPt + frame.widthPt - insets.rightPt,
      widthPt: 0,
      bands: input.bands,
    });

    return {
      index,
      topPt: input.topPt,
      anchorTopPt: input.anchorTopPt,
      heightPt: slot.topPt + height.heightPt + belowPt + afterPt - input.topPt,
      lines: [],
      marker: markerAt(number, slot.topPt + height.baseFromTopPt),
      markTopPt: slot.topPt + height.seatPt,
      contentBottomPt: slot.topPt + height.heightPt,
      widowControl: paragraphFrame.widowControl,
      startsPage: input.startsPage,
      endsPage,
      contentWidthPt: slot.leftPt - frame.leftPt + input.markWidthPt,
      clipTo: null,
      paint,
    };
  }

  const placed = laid.map((each, at) => {
    // A justified line fills the room it was fitted into, which an object beside
    // it may have narrowed. Only the paragraph's last line is left as it fell; a
    // line that ended at a manual break is stretched like any other.
    const filled =
      paragraphFrame.alignment === "justify" && at < laid.length - 1
        ? justifyLine(each.line, each.slot.rightPt - each.slot.leftPt)
        : each.line;

    return {
      line: filled,
      leftPt: lineStartPt(paragraphFrame, each.slot.leftPt, each.slot.rightPt, filled.widthPt),
      topPt: each.slot.topPt,
      heightPt: each.height.heightPt,
      seatPt: each.height.seatPt,
      baselinePt: each.slot.topPt + each.height.baseFromTopPt,
      startsPage: each.startsPage,
    };
  });

  const last = laid[laid.length - 1];
  const bottomPt = last === undefined ? input.topPt : last.slot.topPt + last.height.heightPt;

  return {
    index,
    topPt: input.topPt,
    anchorTopPt: input.anchorTopPt,
    heightPt: bottomPt + belowPt + afterPt - input.topPt,
    lines: placed,
    marker: markerAt(number, placed[0]?.baselinePt ?? input.topPt),
    markTopPt: last === undefined ? input.topPt : last.slot.topPt + last.height.seatPt,
    contentBottomPt: bottomPt,
    widowControl: paragraphFrame.widowControl,
    startsPage: input.startsPage,
    endsPage,
    contentWidthPt: placed.reduce(
      (widest, line) => Math.max(widest, line.leftPt - frame.leftPt + line.line.widthPt),
      0,
    ),
    clipTo: null,
    paint,
  };
}

// A line as it came out of the paragraph's text and where it ended up, before
// anything is asked about the line above or below it.
type LaidLine = {
  readonly line: TextLine;
  readonly slot: LineSlot;
  readonly height: LineHeight;
  readonly startsPage: boolean;
};

// The paragraph's lines, and whether it ended on a page break: a break with
// nothing after it draws no line of its own, so what it has to say is said here.
type LaidLines = {
  readonly lines: readonly LaidLine[];
  readonly endsPage: boolean;
};

// How many goes a line gets at settling on a height. A line broken again at a
// narrower width can come out a different height, which moves the objects it has
// to clear; two rounds settle anything these documents hold, and a third is the
// backstop rather than a rule.
const SETTLING_ROUNDS = 3;

// Each line is broken at the room the frame gives it, placed clear of whatever
// stands beside it, and, where that left it narrower than the frame, broken again
// at the width it was left. Word breaks at the narrower width, so a paragraph
// beside an object runs on in shorter lines rather than dropping past it whole.
function layOutLines(flow: LineFlow, input: LayOutParagraphInput): LaidLines {
  const { paragraphFrame, frame, number } = input;
  const insets = insetsOf(paragraphFrame);
  const laid: LaidLine[] = [];
  let rest: LineFlow = flow;
  let top = input.topPt + input.spacing.beforePt + borderRoomPt(input.paint?.borders.top ?? null);

  for (;;) {
    const at = laid.length;
    const roomPt = at === 0 ? input.firstLineRoomPt : input.roomPt;
    const firstLinePt = at === 0 ? insets.firstLinePt : 0;
    // The number takes the first line's own start, so the text after it begins
    // wherever the number's suffix moved on to.
    const startPt =
      at === 0 && number !== null ? number.textStartPt : frame.leftPt + insets.leftPt + firstLinePt;
    const endPt = frame.leftPt + frame.widthPt - insets.rightPt;

    const leastPt = rest.leastPt;
    const startsPage = rest.startsPage;
    let taken = rest.next(roomPt);
    if (taken === null) return { lines: laid, endsPage: startsPage };

    let height = heightOfLine(taken.line, at, input);
    let slot = slotFor(top, height.heightPt, leastPt, startPt, endPt, input.bands);

    for (let round = 1; round < SETTLING_ROUNDS; round += 1) {
      const narrowedPt = slot.rightPt - slot.leftPt;
      // Only an object beside the line narrows it. A line the frame itself leaves
      // no room for is one a shape refusing to wrap runs past, and is left alone.
      if (narrowedPt >= endPt - startPt - EPSILON) break;

      const again = rest.next(Math.min(roomPt, narrowedPt));
      if (again === null) break;

      const settled = heightOfLine(again.line, at, input);
      taken = again;
      if (settled.heightPt === height.heightPt) break;

      height = settled;
      slot = slotFor(top, height.heightPt, leastPt, startPt, endPt, input.bands);
    }

    laid.push({ line: taken.line, slot, height, startsPage });
    rest = taken.rest;
    top = slot.topPt + height.heightPt;
  }
}

// A line is not asked to fit whole, since it is broken again to whatever width it
// is given: what it asks of a run of space is room for the word it has to start
// with.
const slotFor = (
  topPt: number,
  heightPt: number,
  widthPt: number,
  leftPt: number,
  rightPt: number,
  bands: readonly WrapBand[],
): LineSlot => fitLine({ topPt, heightPt, leftPt, rightPt, widthPt, bands });

const raisedBy = (line: TextLine, at: number, input: LayOutParagraphInput): number =>
  at === 0 ? liftOfNumber(input.number, line) : 0;

// A line a tab alone holds open has nothing measured on it to give it a height, so
// it takes the tallest mark the paragraph has, as an empty paragraph does.
function heightOfLine(line: TextLine, at: number, input: LayOutParagraphInput): LineHeight {
  const raisedPt = raisedBy(line, at, input);
  const held = line.segments.length === 0 ? input.markHeightPt : 0;
  return seatedHeight(
    {
      naturalPt: Math.max(line.heightPt, held) + raisedPt,
      ascentPt: line.ascentPt + raisedPt,
      seatPt: line.seatPt,
      // A number lifts the line by as much as its own face reaches, so it is part
      // of what a multiple over the line is taken of.
      fontHeightPt: Math.max(line.fontHeightPt, held) + raisedPt,
    },
    input.paragraphFrame,
  );
}

// How tall a line stands in the stack, how far down that room its own text sits,
// and where its baseline falls from the line's top.
type LineHeight = {
  readonly heightPt: number;
  readonly seatPt: number;
  readonly baseFromTopPt: number;
};

// A line asked for exactly so much room is a slot the text is dropped into rather
// than one measured from it: the baseline lands four fifths of the way down the
// room asked for, whatever the face and whatever the size, so a slot shorter than
// its text cuts the top off the glyphs. Measured off Word's own pdf over lines of
// 6 to 48pt, in Calibri at two sizes and in Arial, every one of them within the
// quarter point the pdf rounds to.
const EXACT_BASELINE = 0.8;

// What a line measured to before its paragraph's rule was applied to it.
type NaturalLine = {
  readonly naturalPt: number;
  readonly ascentPt: number;
  // Room the line already seated its own content under, before the paragraph's
  // rule opened any more above it.
  readonly seatPt: number;
  readonly fontHeightPt: number;
};

// Where a line rule leaves room its text does not need, `atLeast` takes it above:
// the text keeps a line of its own height at the foot of the room. Under `auto` the
// room falls below and the text keeps the top.
function seatedHeight(line: NaturalLine, frame: ParagraphFrame): LineHeight {
  const heightPt = spacedHeightPt(line, frame);
  if (frame.lineTwips !== null && frame.lineRule === "exact") {
    return { heightPt, seatPt: 0, baseFromTopPt: heightPt * EXACT_BASELINE };
  }

  const openedPt = frame.lineRule === "atLeast" ? Math.max(0, heightPt - line.naturalPt) : 0;
  const seatPt = openedPt + line.seatPt;
  return { heightPt, seatPt, baseFromTopPt: seatPt + line.ascentPt };
}

// Room is a difference of exact ratios, so only the last bits of one need absorbing.
const EPSILON = 1e-9;

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

// A multiple is taken of the line the paragraph's own faces would have made, and
// what it opens is added to whatever the line measured to. For a line of text
// those are the same number and this is the multiple over the whole line; for one
// holding a drawing they are not, and Word gives the chart in a 1.2 line paragraph
// its own height and a fifth of a line of text, not a fifth of the chart.
function spacedHeightPt(line: NaturalLine, frame: ParagraphFrame): number {
  const { lineTwips, lineRule } = frame;
  if (lineTwips === null) return line.naturalPt;
  if (lineRule === "exact") return twipsToPoints(lineTwips);
  if (lineRule === "atLeast") return Math.max(line.naturalPt, twipsToPoints(lineTwips));
  const multiple = lineTwips / LINE_MULTIPLE_UNITS;
  return line.naturalPt + (multiple - 1) * line.fontHeightPt;
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
