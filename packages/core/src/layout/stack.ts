import { blockParagraphs } from "../docx/blocks.js";
import type {
  Block,
  CellVerticalAlign,
  Paragraph,
  TableCell,
  TableInsets,
  TablePositioning,
  TableRow,
} from "../docx/blocks.js";
import {
  borderExtentPt,
  resolveCellBorders,
  NO_BORDERS,
  SIDES,
  type Border,
  type Borders,
  type CellBorders,
} from "../docx/borders.js";
import { numberParagraphs, type ParagraphNumber } from "../docx/list-numbers.js";
import type { NumberSuffix } from "../docx/numbering.js";
import { readRuns, type TextRun } from "../docx/runs.js";
import { W_NS, type SectionClose } from "../docx/section.js";
import {
  measuresTheIndentToTheText,
  roundsAnchorsToTwips,
  DEFAULT_SETTINGS,
  type DocumentSettings,
} from "../docx/settings.js";
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
import type { Column } from "./columns.js";
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
  // What the paragraph asks of a page break running through it, and of one falling
  // between it and the paragraph after it. Only the break itself can act on either.
  readonly widowControl: boolean;
  readonly keepNext: boolean;
  // Whether the paragraph asked for a page of its own, and whether it ended on a
  // break that puts whatever follows it on one.
  readonly startsPage: boolean;
  readonly endsPage: boolean;
  // Which kind of break that was, since the page a section opens keeps the room
  // the paragraph opening it asks for above itself and the page any other break
  // opens does not.
  readonly endsPageAtASection: boolean;
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

// A row a page break is not allowed to run through, which moves whole to the page
// under it instead. Word tears an ordinary row at a line, so what is listed here is
// the two rows that refuse: one saying `w:cantSplit`, and one asking to be taller
// than its own text needs, whose empty foot is what will not come apart.
//
// `opensAt` is the first paragraph in it, which is where the break has to be
// decided: by the time a later one is reached the page it stands on is settled.
export type UntornRow = {
  readonly topPt: number;
  readonly bottomPt: number;
  readonly opensAt: number;
};

// The room a wrapping object takes down the page, and the paragraph anchoring it.
// An object that will not fit in what is left below moves to the page under it and
// takes that paragraph with it, so like a row that refuses to be torn this has to
// be decided at the paragraph rather than at the object.
//
// Only an object text wraps round is listed. One wrapping nothing hangs past the
// foot of the page and moves neither itself nor anything else, which is why the
// band is what this is read off: a band is made for exactly the wraps that move.
export type AnchoredObject = {
  readonly topPt: number;
  readonly bottomPt: number;
  readonly anchoredAt: number;
};

export type StackMeasurement =
  | {
      readonly kind: "measured";
      readonly boxes: readonly ParagraphBox[];
      readonly cells: readonly PlacedCell[];
      readonly untornRows: readonly UntornRow[];
      readonly anchoredObjects: readonly AnchoredObject[];
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
  // The paragraphs closing a section, and what each one's break does. Only the
  // body has any: a `w:sectPr` inside a cell governs the story in that cell and
  // closes no section, and a header has none at all.
  readonly sectionsClosed?: ReadonlyMap<number, SectionClose>;
  // The frame a block of the body's own stands in, where its section states one of
  // its own. A story is one frame from top to bottom until a section break changes
  // it partway down, and a block inside a table cell is framed by the cell.
  readonly frameOf?: (block: Block) => Frame | undefined;
  // Where the columns of the section a block stands in run across that frame. One
  // column, which is what all but sixteen of the corpus documents keep, answers
  // with the frame itself and takes the same path it always did.
  readonly columnsOf?: (block: Block) => readonly Column[];
  // How tall the body's text is, which is how tall a column that fills rather than
  // evens out may be. Measuring is one column with no pages in it, so the foot of a
  // page cannot be read off the stack's own top: a section that fills its columns is
  // one that opens a page, and its own top is the body's. Left out, a column runs on
  // for ever, which is what every story that is not the body does.
  readonly bodyHeightPt?: number;
};

export type BandResolver = (paragraph: Paragraph, topPt: number) => readonly WrapBand[];

type Context = Omit<MeasureStackInput, "blocks" | "originPt" | "leftPt" | "widthPt"> & {
  readonly numbers: ReadonlyMap<number, ParagraphNumber>;
  readonly settings: DocumentSettings;
  // A page break inside a cell is no break at all, and a cell is the only place
  // that has to know it.
  readonly inCell: boolean;
  // Whether these blocks are one column of a section's run, whose own column breaks
  // have already been spent dividing it.
  readonly inColumn: boolean;
  // The style of the table a cell's text stands in, which that text reads between
  // the document's defaults and its own style. Null anywhere but inside a cell.
  readonly tableStyleId: string | null;
};

// Where a story's text runs across the page, which a section break can change
// partway down the body.
export type Frame = {
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
    inColumn: false,
    tableStyleId: null,
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
  storyFrame: Frame,
): StackMeasurement {
  const boxes: ParagraphBox[] = [];
  const cells: PlacedCell[] = [];
  const untornRows: UntornRow[] = [];
  const anchoredObjects: AnchoredObject[] = [];
  let top = originPt;
  // The objects met so far, which grows as the stack walks forward: every line
  // from an object's own paragraph on has to sit clear of it.
  let standing: readonly WrapBand[] = [];
  // Where the paragraph coming up anchors its objects, which is its own top until
  // the one above it was dropped past one of them.
  let anchoredAtPt: number | null = null;

  // The last block of a run of columns already laid out, which the walk below skips
  // past: a column run is measured whole, since every column of it starts at the
  // same top and the run is as tall as the tallest.
  let laidOutTo = -1;

  for (const [at, block] of blocks.entries()) {
    if (at <= laidOutTo) continue;
    const frame = context.inCell ? storyFrame : (context.frameOf?.(block) ?? storyFrame);

    const columns = context.inCell ? [] : (context.columnsOf?.(block) ?? []);
    if (columns.length > 1) {
      const run = columnRunFrom(blocks, at, context);
      const measured = measureColumnRun(
        run.blocks,
        context,
        top,
        columns,
        context.bodyHeightPt ?? Number.POSITIVE_INFINITY,
        run.balanced,
      );
      if (measured.kind === "blocked") return measured;
      boxes.push(...measured.boxes);
      cells.push(...measured.cells);
      untornRows.push(...measured.untornRows);
      anchoredObjects.push(...measured.anchoredObjects);
      standing = [];
      anchoredAtPt = null;
      laidOutTo = run.endsAt;
      top += measured.heightPt;
      continue;
    }

    if (block.kind === "paragraph") {
      const { paragraph } = block;
      const neighbours = {
        above: paragraphAt(blocks, at - 1),
        below: paragraphAt(blocks, at + 1),
      };
      // An object wraps the text on the page its anchor landed on and no other, so
      // an explicit break is where the objects met before it are let go of.
      // Measuring is one column with no pages in it, and a break is the only place
      // in it that a later page is known about: without this a narrow object on one
      // page moves the first line of the next one out of its way, which the page
      // break then leaves standing where nothing is beside it.
      if (opensPage(paragraph, boxes.at(-1), context)) standing = [];
      const anchorTopPt = anchoredAtPt ?? top;
      const own = bandsOf(paragraph, anchorTopPt, context);
      const bands = [...standing, ...own];
      for (const band of own) {
        anchoredObjects.push({
          topPt: band.topPt,
          bottomPt: band.bottomPt,
          anchoredAt: paragraph.index,
        });
      }
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

    // **A table `w:tblpPr` takes out of the flow keeps no room in it**, so the
    // paragraphs under it are drawn where they would have stood had the table never
    // been there, and the table is a band they wrap round instead. Measured on
    // 2026-08-08 by the authored `positioned-table` document.
    if (block.positioning !== null && !context.inCell) {
      const placed = positionedFrame(block, block.positioning, frame);
      const measured = measureTable(block, context, top + placed.topPt, placed.frame);
      if (measured.kind === "blocked") return measured;
      boxes.push(...measured.boxes);
      cells.push(...measured.cells);
      standing = [
        ...standing,
        bandRound(placed, top + placed.topPt, measured.heightPt, block.positioning),
      ];
      anchoredAtPt = null;
      continue;
    }

    const measured = measureTable(block, context, top, frame);
    if (measured.kind === "blocked") return measured;
    boxes.push(...measured.boxes);
    cells.push(...measured.cells);
    untornRows.push(...measured.untornRows);
    anchoredObjects.push(...measured.anchoredObjects);
    anchoredAtPt = null;
    top += measured.heightPt;
  }

  return {
    kind: "measured",
    boxes,
    cells,
    untornRows,
    anchoredObjects,
    heightPt: top - originPt,
  };
}

// The blocks a section's columns hold: everything from the block opening the run up
// to and including the paragraph that closes the section.
function columnRunFrom(
  blocks: readonly Block[],
  from: number,
  context: Context,
): { readonly blocks: readonly Block[]; readonly endsAt: number; readonly balanced: boolean } {
  for (let at = from; at < blocks.length; at += 1) {
    const block = blocks[at];
    const close =
      block !== undefined && block.kind === "paragraph"
        ? context.sectionsClosed?.get(block.paragraph.index)
        : undefined;
    if (close !== undefined) {
      return { blocks: blocks.slice(from, at + 1), endsAt: at, balanced: !close.opensAPage };
    }
  }
  // The body's own last section closes no paragraph, and Word leaves its columns
  // ragged: it is the one section whose break opens no page of anything.
  return { blocks: blocks.slice(from), endsAt: blocks.length - 1, balanced: false };
}

// **A column is a page-height run of the text, and every column of a section starts
// at the same top.** Measured on 2026-08-07 by the authored `columns` document, and
// built on 2026-08-08.
//
// Two things divide a section between its columns, and the second is the one the
// corpus turns on.
//
// **A column fills to the foot of the text and the next one starts at the top**,
// which is what a section closed by a break that opens a page does: fourteen lines
// in two columns filled both columns of one page and put the last two in the first
// column of the next.
//
// **A section closed by a continuous break has its columns evened out** instead:
// four lines in two columns that hold twelve were drawn two and two, not four and
// none. **29 of the 30 column sections in the corpus are closed that way.** Evening
// them out is the same fill asked of the shortest column that still holds the run,
// which is what the search below looks for.
//
// **A break of the section's own says where a column starts whatever the heights
// say.** 23 of the 30 carry enough of them to divide the run outright, and every one
// of those 25 breaks either stands alone in its paragraph or opens one, so the
// division falls between two blocks. **A break between two runs of one paragraph is
// measured and not built**: it costs the authored document's fourth case and nothing
// in the wild.
//
// The run is as tall as its tallest column, which is where the text under it carries
// on: Word started the section below a balanced one under the taller of its two.
function measureColumnRun(
  blocks: readonly Block[],
  context: Context,
  topPt: number,
  columns: readonly Column[],
  heightPt: number,
  balanced: boolean,
): StackMeasurement {
  // A column's own frame stands instead of the section's, and the breaks that
  // divided the run are places between blocks now rather than pieces of text.
  // A column's own frame stands instead of the section's, and nothing inside a
  // column run asks for the columns again.
  const inColumn: Context = {
    ...context,
    inColumn: true,
    frameOf: () => undefined,
    columnsOf: () => [],
  };

  const forced = new Set<number>();
  for (const [at, block] of blocks.entries()) {
    if (at > 0 && opensAColumn(block, inColumn)) forced.add(at);
  }

  if (!balanced) {
    const filled = fillColumns(blocks, inColumn, topPt, columns, forced, heightPt);
    return filled.kind === "blocked" ? filled : filled.measured;
  }

  // Evening the columns out is the same fill asked of the shortest column that still
  // holds the whole run, and the heights worth asking for are the ones the run's own
  // blocks make.
  const whole = fillColumns(blocks, inColumn, topPt, columns, forced, Number.POSITIVE_INFINITY);
  if (whole.kind === "blocked") return whole;

  // The shortest that works is wanted, so the candidates are tried from the shortest
  // up and the first one that holds the run is the answer.
  const candidates = [...new Set(whole.bottomsPt)].sort((one, other) => one - other);
  for (const candidate of candidates) {
    const filled = fillColumns(blocks, inColumn, topPt, columns, forced, candidate);
    if (filled.kind === "blocked") return filled;
    // The last column holds whatever is over, so a height every block found a column
    // at is not yet a height the columns were evened to: the tallest of them has to
    // come in under it as well.
    if (!filled.whole || filled.measured.heightPt > candidate + EPSILON) continue;
    return filled.measured;
  }
  return whole.measured;
}

type ColumnFill =
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker }
  | {
      readonly kind: "filled";
      readonly measured: Extract<StackMeasurement, { readonly kind: "measured" }>;
      // Whether every block of the run found a column, which is what says a height
      // is tall enough to be worth evening out at.
      readonly whole: boolean;
      // How far below the run's top each block reached, which are the heights a
      // column could be cut to.
      readonly bottomsPt: readonly number[];
    };

// The blocks dealt into the columns, each column measured in its own frame and every
// one of them starting at the run's own top. A block that will not fit in the room
// left opens the next column, and one carrying a break of its own opens one whatever
// room is left.
function fillColumns(
  blocks: readonly Block[],
  context: Context,
  topPt: number,
  columns: readonly Column[],
  forced: ReadonlySet<number>,
  heightPt: number,
): ColumnFill {
  const blockOf = new Map<number, number>();
  for (const [at, block] of blocks.entries()) {
    for (const paragraph of blockParagraphs([block])) blockOf.set(paragraph.index, at);
  }

  const boxes: ParagraphBox[] = [];
  const cells: PlacedCell[] = [];
  const untornRows: UntornRow[] = [];
  const bottomsPt: number[] = [];
  let tallestPt = 0;
  let from = 0;

  for (const [at, column] of columns.entries()) {
    if (from >= blocks.length) break;
    const measured = measureBlocks(blocks.slice(from), context, topPt, column);
    if (measured.kind === "blocked") return measured;

    const last = at === columns.length - 1;
    const cut = last
      ? blocks.length
      : cutColumnAt(measured.boxes, blockOf, forced, topPt, heightPt);

    if (cut >= blocks.length) {
      boxes.push(...measured.boxes);
      cells.push(...measured.cells);
      untornRows.push(...measured.untornRows);
      for (const box of measured.boxes) bottomsPt.push(box.topPt + box.heightPt - topPt);
      tallestPt = Math.max(tallestPt, measured.heightPt);
      from = blocks.length;
      break;
    }

    const kept = measureBlocks(blocks.slice(from, cut), context, topPt, column);
    if (kept.kind === "blocked") return kept;
    boxes.push(...kept.boxes);
    cells.push(...kept.cells);
    untornRows.push(...kept.untornRows);
    for (const box of measured.boxes) bottomsPt.push(box.topPt + box.heightPt - topPt);
    tallestPt = Math.max(tallestPt, kept.heightPt);
    from = cut;
  }

  return {
    kind: "filled",
    measured: {
      kind: "measured",
      boxes,
      cells,
      untornRows,
      anchoredObjects: [],
      heightPt: tallestPt,
    },
    whole: from >= blocks.length,
    bottomsPt,
  };
}

// The first block of the run that belongs in the next column: the one carrying a
// break of its own, or the one whose foot passed the height the column was given.
// The block a column opens with always stays, however tall it is, since a column
// that carries nothing forward never ends.
function cutColumnAt(
  boxes: readonly ParagraphBox[],
  blockOf: ReadonlyMap<number, number>,
  forced: ReadonlySet<number>,
  topPt: number,
  heightPt: number,
): number {
  let opened: number | null = null;
  for (const box of boxes) {
    const place = blockOf.get(box.index);
    if (place === undefined) continue;
    if (opened === null) opened = place;
    if (place === opened) continue;
    if (forced.has(place)) return place;
    if (box.topPt + box.heightPt > topPt + heightPt + EPSILON) return place;
  }
  return Number.POSITIVE_INFINITY;
}

// Where a paragraph carries a break asking for the next column. Every one of the 25
// in the corpus either stands alone in its paragraph or opens one, so a break is a
// place between two blocks rather than inside one.
const opensAColumn = (block: Block, context: Context): boolean =>
  block.kind === "paragraph" &&
  readRuns(block.paragraph, context.styles).some((run) =>
    run.pieces.some((piece) => piece.kind === "break" && piece.endsColumn),
  );

const bandsOf = (paragraph: Paragraph, topPt: number, context: Context): readonly WrapBand[] =>
  context.bandsFor === undefined ? [] : context.bandsFor(paragraph, topPt);

// Whether a paragraph opens a page: it asks for one of its own, or the paragraph
// above it ended on a break. A cell is the one place Word ignores both.
const opensPage = (
  paragraph: Paragraph,
  above: ParagraphBox | undefined,
  context: Context,
): boolean =>
  !context.inCell &&
  (above?.endsPage === true ||
    resolveParagraphFrame(paragraph, context.styles, context.tableStyleId).pageBreakBefore);

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

// Where a positioned table stands, and how wide it turned out to be.
type PositionedTable = {
  readonly frame: Frame;
  // How far below the flow the table's own top is, which is what `w:tblpY` asks
  // for and nothing at all where it asks for nothing.
  readonly topPt: number;
  readonly widthPt: number;
};

// The table's own width, which is what its first row's cells come to. Read the same
// way `measureRow` reads them, since a table stating a grid and a table whose cells
// state their own widths both have to come out where Word drew them.
function tableWidthPt(block: Table, frame: Frame): number {
  const cells = block.rows[0]?.cells ?? [];
  return cells.reduce(
    (width, cell, at) =>
      width + (cellWidthPt(cell) ?? columnWidthPt(block.gridTwips, at) ?? frame.widthPt),
    0,
  );
}

// **What `w:tblpX` is measured from is what `w:horzAnchor` names**: the edge of the
// sheet for `page`, and the text frame's own left for `margin` and for the `column`
// a table stating no anchor takes. Measured on 2026-08-08 by the authored
// `positioned-table` document, which places the same table an inch off each of the
// three: Word drew it at 72 for the page and at 108 for the other two, the frame
// there beginning at 36. One column makes the margin and the column the same place,
// and nothing here can tell them apart until there is more than one.
//
// **`w:tblpXSpec` names an edge rather than a distance**, and `right` puts the
// table's own right edge on the frame's.
function positionedFrame(
  block: Table,
  positioning: TablePositioning,
  frame: Frame,
): PositionedTable {
  const widthPt = tableWidthPt(block, frame);
  const originPt = positioning.horizontalAnchor === "page" ? 0 : frame.leftPt;
  const leftPt =
    positioning.xSpec === "right"
      ? frame.leftPt + frame.widthPt - widthPt
      : positioning.xSpec === "center"
        ? frame.leftPt + (frame.widthPt - widthPt) / 2
        : originPt + twipsToPoints(positioning.xTwips);

  return {
    frame: { leftPt, widthPt: frame.widthPt },
    topPt: twipsToPoints(positioning.yTwips),
    widthPt,
  };
}

// The room the text keeps clear of a positioned table: its own rectangle grown by
// the distances it asks the text to stay off each side of it.
const bandRound = (
  placed: PositionedTable,
  topPt: number,
  heightPt: number,
  positioning: TablePositioning,
): WrapBand => ({
  leftPt: placed.frame.leftPt - twipsToPoints(positioning.leftFromTextTwips),
  rightPt: placed.frame.leftPt + placed.widthPt + twipsToPoints(positioning.rightFromTextTwips),
  topPt: topPt - twipsToPoints(positioning.topFromTextTwips),
  bottomPt: topPt + heightPt + twipsToPoints(positioning.bottomFromTextTwips),
});

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
  const outerTopPt = halfOf(first.map((cell) => cell.agreed.top));
  const outerBottomPt = halfOf(last.map((cell) => cell.agreed.bottom));
  const outerLeftPt = halfOf(borders.map((row) => row[0]?.drawn.left ?? null));

  // An old document's indent is measured to the text rather than to the table, so
  // the first column's own margin stands outside the indent instead of inside it
  // and the table's edge moves left by the whole of it.
  const openingCell = block.rows[0]?.cells[0];
  const insetPt =
    measuresTheIndentToTheText(context.settings) && openingCell !== undefined
      ? -leftMarginOf(openingCell, block.insets, first[0]?.drawn ?? NO_BORDERS)
      : outerLeftPt;

  const boxes: ParagraphBox[] = [];
  const cells: PlacedCell[] = [];
  const untornRows: UntornRow[] = [];
  const rowFrame = {
    leftPt: frame.leftPt + twipsToPoints(block.insets.indentTwips) + insetPt,
    widthPt: frame.widthPt,
  };

  // Everything inside the table reads the table's own style, which the paragraphs
  // in its cells sit under.
  const inTable: Context = { ...context, tableStyleId: block.styleId };

  let top = topPt + outerTopPt;
  for (const [at, row] of block.rows.entries()) {
    const measured = measureRow(row, borders[at] ?? [], inTable, top, rowFrame, block.insets, {
      gridTwips: block.gridTwips,
    });
    if (measured.kind === "blocked") return measured;
    boxes.push(...measured.boxes);
    cells.push(...measured.cells);
    untornRows.push(...measured.untornRows);
    top += measured.heightPt;
  }

  // A cell is measured with no bands at all, so nothing inside a table can anchor
  // an object a page break has to make room for.
  return {
    kind: "measured",
    boxes,
    cells,
    untornRows,
    anchoredObjects: [],
    heightPt: top + outerBottomPt - topPt,
  };
}

// How far a cell holds its text off its own left wall: the margin it asks for, or
// the room its border needs where that is the wider of the two.
const leftMarginOf = (cell: TableCell, insets: TableInsets, borders: Borders): number =>
  Math.max(twipsToPoints(cell.margins.leftTwips ?? insets.leftTwips), halfOf([borders.left]));

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
  readonly innerUntorn: readonly UntornRow[];
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
// **A border is room on top of the margin rather than instead of it.** The half of
// the line that falls inside the cell is cleared, and then the margin is cleared
// after it, so two rows lined with 6pt and held off their walls by 5 stand 36pt
// apart and not 30. Measured on 2026-08-07 by the authored `lined-rows` document
// over widths from half a point to six at two margins, and all eleven cases are the
// margins either side plus the whole of the line between them.
//
// What was here before took the larger of the two, which is right only where one of
// them is nought, and every table in a real document is out by a line a row for it.
function measureRow(
  row: TableRow,
  borders: readonly CellBorders[],
  context: Context,
  topPt: number,
  frame: Frame,
  insets: TableInsets,
  table: { readonly gridTwips: readonly number[] },
): StackMeasurement {
  const measured: MeasuredCell[] = [];
  const topMarginPt =
    rowMarginPt(row, insets, "topTwips") + halfOf(borders.map((of) => of.agreed.top));
  // The cell's own margin at the foot, kept apart from the half of the line cleared
  // after it, because a row told exactly how tall to be counts one and not the other.
  const bottomCellMarginPt = rowMarginPt(row, insets, "bottomTwips");
  const bottomMarginPt = bottomCellMarginPt + halfOf(borders.map((of) => of.agreed.bottom));
  let contentHeightPt = 0;
  let leftPt = frame.leftPt;

  // A cell is measured from its own origin and only then moved down to the row, so
  // the page coordinates a wrapping object stands in cannot reach inside one.
  const inCell: Context = { ...context, bandsFor: () => [], inCell: true };
  const untornRows: UntornRow[] = [];

  for (const [at, cell] of row.cells.entries()) {
    const widthPt = cellWidthPt(cell) ?? columnWidthPt(table.gridTwips, at) ?? frame.widthPt;
    const own = borders[at]?.drawn ?? NO_BORDERS;
    const leftMarginPt = leftMarginOf(cell, insets, own);
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
      innerUntorn: of.untornRows,
      heightPt: of.heightPt,
      leftPt,
      widthPt,
      fillColor: cell.fillColor,
      borders: own,
    });
    contentHeightPt = Math.max(contentHeightPt, of.heightPt);
    leftPt += widthPt;
  }

  const heldPt = topMarginPt + contentHeightPt + bottomMarginPt;
  const heightPt = rowHeightPt(row, contentHeightPt, {
    marginsPt: topMarginPt + bottomMarginPt,
    bottomCellMarginPt,
  });
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
    for (const inner of cell.innerUntorn)
      untornRows.push({ ...inner, topPt: inner.topPt + offset, bottomPt: inner.bottomPt + offset });
    cells.push({
      leftPt: cell.leftPt,
      topPt,
      widthPt: cell.widthPt,
      heightPt,
      fillColor: cell.fillColor,
      borders: cell.borders,
    });
  }

  // Word tears an ordinary row at a line, and refuses two: one saying so, and one
  // standing taller than its own text, whose empty foot has no line to be torn at.
  // Both were measured on 2026-08-07 by the authored `tearing` document: a row
  // asking to be 150pt tall with 48pt of text in it moved whole where 102pt was
  // left, and so did the same row with 144pt of text, while a row asking to be
  // 48pt tall with 144pt of text in it was torn like any other.
  const opensAt = boxes[0]?.index;
  if (opensAt !== undefined && (row.cantSplit || heightPt > heldPt + EPSILON)) {
    untornRows.push({ topPt, bottomPt: topPt + heightPt, opensAt });
  }

  return { kind: "measured", boxes, cells, untornRows, anchoredObjects: [], heightPt };
}

// The largest margin any cell in the row asks for at that side, which is what
// every cell in it is held off the wall by.
function rowMarginPt(row: TableRow, insets: TableInsets, side: "topTwips" | "bottomTwips"): number {
  const twips = row.cells.map((cell) => cell.margins[side] ?? insets[side]);
  return twipsToPoints(Math.max(insets[side], ...twips));
}

// **A stated height is a floor under the text and not under the row.** What the row
// asks for stands instead of what its cells hold, and the margins holding them off
// its walls are then cleared on top of that, the half of the line between two rows
// that falls inside each of them included.
//
// Measured on 2026-08-07 by the authored `stated-row-heights` document, which asks
// for 60pt a row against 20pt of text. Rows held off their walls by 5pt stand 70.08
// apart and rows lined at six points stand 66, and rows that are both stand 76.08:
// the stated height, the two margins and the whole of the line between them. Taking
// the larger of the stated height and the whole room the row held cost a real
// three page document a fraction of a point at every row boundary in it.
//
// **A row saying it is exact is the stated height and the cell's own margin at the
// foot, and the line between two rows takes no room in it at all.** The same document
// says so over four cases: 60pt asked for came out 65.04 both with a 6pt line and
// without one, so long as the cells were held off their walls by 5.04; and 60pt
// exactly, line or no line, wherever they were held off by nothing.
//
// Every margin in that document is the same at the top as at the foot, so which of
// the two it is comes from the `tables` document instead: one row there is told to be
// 14.4pt with a cell holding its text 21.6pt off the top wall and nothing off the
// bottom, and Word draws that row at 14.4. So the stated height covers the top margin
// and the text under it, and only the foot is cleared after.
function rowHeightPt(
  row: TableRow,
  contentHeightPt: number,
  margins: { readonly marginsPt: number; readonly bottomCellMarginPt: number },
): number {
  if (row.height === null) return margins.marginsPt + contentHeightPt;
  const askedPt = twipsToPoints(row.height.twips);
  return row.height.exact
    ? margins.bottomCellMarginPt + askedPt
    : margins.marginsPt + Math.max(contentHeightPt, askedPt);
}

// **A cell that states no width of its own is as wide as the column it stands in.**
// The grid is what a table is drawn on and a `w:tcW` is the cell's own preference
// over it, so a table stating a grid and no cell widths at all, which a real
// document does, had every column as wide as the whole text frame before this.
function columnWidthPt(gridTwips: readonly number[], at: number): number | null {
  const twips = gridTwips[at];
  return twips === undefined ? null : twipsToPoints(twips);
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
  const paragraphMark = resolveParagraphMark(paragraph, context.styles, context.tableStyleId);
  const marks: readonly ParagraphMark[] = [
    paragraphMark,
    ...resolveRunMarks(paragraph, context.styles),
  ];

  // **What holds a line open with nothing measured on it is the paragraph's own
  // mark**, and not the tallest thing the paragraph is written in: a line holding
  // one 24pt space, or one 24pt tab, comes out at the 12pt mark behind it. Measured
  // on 2026-08-07 by the authored `trailing-space` document. Every run's mark is
  // still asked for its height, since a face this machine cannot answer for blocks
  // the document whether or not its run ends up on a line.
  let markHeight = 0;
  for (const mark of marks) {
    const height = heightOf(mark, context.metricsFor);
    if (height.kind === "blocked") {
      return {
        kind: "blocked",
        blocker: blockerFor(mark, context.part, paragraph.index),
      };
    }
    if (mark === paragraphMark) markHeight = height.value;
  }

  const paragraphFrame = resolveParagraphFrame(paragraph, context.styles, context.tableStyleId);
  const runs = flowing(readRuns(paragraph, context.styles), context.inCell, context.inColumn);
  const insets = insetsOf(paragraphFrame);
  const number = context.numbers.get(paragraph.index);
  const sectionClose = context.sectionsClosed?.get(paragraph.index);
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
      endsPageAtASection: sectionClose?.opensAPage === true,
      closesASection: sectionClose !== undefined,
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
function flowing(runs: readonly TextRun[], inCell: boolean, inColumn: boolean): readonly TextRun[] {
  if (!inCell && !inColumn) return runs;
  return runs.map((run) => ({
    ...run,
    pieces: run.pieces.filter(
      (piece) =>
        piece.kind !== "break" || !((inCell && piece.endsPage) || (inColumn && piece.endsColumn)),
    ),
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
    other !== null &&
    sameBorders(
      borders,
      resolveParagraphFrame(other, context.styles, context.tableStyleId).borders,
    );

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
  const frame = resolveParagraphFrame(above, context.styles, context.tableStyleId);
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
  readonly startsPage: boolean;
  // Whether a section break stands after the paragraph, which ends a page as a
  // break in its own text would. What the text asked for is read off the text.
  readonly endsPageAtASection: boolean;
  readonly closesASection: boolean;
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
  return droppedPast(layOutWholeParagraph(index, flow, input), input, across);
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
// enough to leave the line somewhere to sit it does not fall at all, and takes
// that room instead: measured on an object wrapped on its largest side, whose line
// came back where it started and three hundred and ninety points across.
function droppedPast(box: ParagraphBox, input: LayOutParagraphInput, across: Span): ParagraphBox {
  const ahead = input.ahead;
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
  const leftPt =
    last === undefined
      ? 0
      : lineStartPt(input.paragraphFrame, slot.leftPt, slot.rightPt, last.line.widthPt);
  if (byPt <= EPSILON && (last === undefined || leftPt === last.leftPt)) return box;
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
            { ...last, leftPt, topPt: last.topPt + byPt, baselinePt: last.baselinePt + byPt },
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
  const brokenByItsText = layOutLines(flow, input);
  const laid = brokenByItsText.lines;
  const endsPage = brokenByItsText.endsPage || input.endsPageAtASection;

  // A paragraph whose whole content is the section break it carries is not laid
  // out at all: it takes no room and holds nothing back. Measured on 2026-08-07 by
  // the authored `section-closer` document. An empty closer offered half a line of
  // room at the foot of a page neither moved on nor pushed the break past it, and
  // one with ten lines of room under it left the paragraph after it exactly where
  // the paragraph above the closer ended. A closer with text in it is an ordinary
  // paragraph: offered the same half line it moved onto the next page and the break
  // then opened a third.
  //
  // A real document closes a section with such a paragraph six points past the foot
  // of its page, and read as an ordinary one it came out a blank page longer than
  // Word drew it.
  if (laid.length === 0 && input.closesASection) {
    return {
      index,
      topPt: input.topPt,
      anchorTopPt: input.anchorTopPt,
      heightPt: 0,
      lines: [],
      marker: null,
      markTopPt: input.topPt,
      contentBottomPt: input.topPt,
      widowControl: paragraphFrame.widowControl,
      keepNext: paragraphFrame.keepNext,
      startsPage: input.startsPage,
      endsPage,
      endsPageAtASection: input.endsPageAtASection,
      contentWidthPt: 0,
      clipTo: null,
      paint: null,
    };
  }

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
    const slot = slotFor({
      topPt: input.topPt + beforePt + abovePt,
      heightPt: height.heightPt,
      roomAbovePt: beforePt,
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
      keepNext: paragraphFrame.keepNext,
      startsPage: input.startsPage,
      endsPage,
      endsPageAtASection: input.endsPageAtASection,
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
    keepNext: paragraphFrame.keepNext,
    startsPage: input.startsPage,
    endsPage,
    endsPageAtASection: input.endsPageAtASection,
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

    // Only the first line has room asked for above it; the rest follow the line
    // before them.
    const roomAbovePt = at === 0 ? input.spacing.beforePt : 0;
    const fit = { roomAbovePt, widthPt: leastPt, leftPt: startPt, rightPt: endPt };

    let height = heightOfLine(taken.line, at, input);
    let slot = slotFor({ ...fit, topPt: top, heightPt: height.heightPt, bands: input.bands });

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
      slot = slotFor({ ...fit, topPt: top, heightPt: height.heightPt, bands: input.bands });
    }

    laid.push({ line: taken.line, slot, height, startsPage });
    rest = taken.rest;
    top = slot.topPt + height.heightPt;
  }
}

type Slot = {
  readonly topPt: number;
  readonly heightPt: number;
  // The room the paragraph asks for above this line, which goes through a wrap with
  // it and is left standing above it wherever it lands.
  readonly roomAbovePt: number;
  readonly widthPt: number;
  readonly leftPt: number;
  readonly rightPt: number;
  readonly bands: readonly WrapBand[];
};

// A line is not asked to fit whole, since it is broken again to whatever width it
// is given: what it asks of a run of space is room for the word it has to start
// with.
//
// **The room a paragraph asks for above itself goes through a wrap with its first
// line.** What has to clear an object is the room and the line together, so a line
// pushed past one stands that room below it rather than against it, and one whose
// room alone reaches an object is drawn beside it as the line itself would be.
// Measured on 2026-08-07 by the authored `space-under-a-wrap` document: a paragraph
// asking 36pt above itself under a box whose foot is at 190 has its line at 226,
// and one whose line stands 44pt clear of the foot of a box beside it is drawn to
// the right of that box all the same.
function slotFor(slot: Slot): LineSlot {
  const { roomAbovePt } = slot;
  const found = fitLine({
    topPt: slot.topPt - roomAbovePt,
    heightPt: slot.heightPt + roomAbovePt,
    leftPt: slot.leftPt,
    rightPt: slot.rightPt,
    widthPt: slot.widthPt,
    bands: slot.bands,
  });
  return { ...found, topPt: found.topPt + roomAbovePt };
}

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
  if (mark.font.kind === "unresolved" && metricsFor.answersForUnresolved !== true)
    return { kind: "blocked" };
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
