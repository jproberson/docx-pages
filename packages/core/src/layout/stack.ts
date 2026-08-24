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
import type { SectionClose } from "../docx/section.js";
import {
  measuresTheIndentToTheText,
  roundsAnchorsToTwips,
  squeezesAJustifiedLine,
  DEFAULT_SETTINGS,
  type DocumentSettings,
} from "../docx/settings.js";
import {
  mergeTableBorders,
  resolveNumberMark,
  resolveParagraphFrame,
  resolveParagraphMark,
  resolveParagraphNumbering,
  resolveRunMarks,
  resolveBandSizes,
  resolveTableBorders,
  resolveTableInsets,
  styleIdOf,
  type CellPosition,
  type InTable,
  type ParagraphFrame,
  type ParagraphMark,
  type StyleTable,
} from "../docx/styles.js";
import { lineHeightPt } from "./font-metrics.js";
import {
  beginLines,
  faceRequestFor,
  justifyLine,
  measureText,
  type LineFlow,
  type LineFlowStart,
  type MeasureFailure,
  type MetricsResolver,
  type TextLine,
} from "./lines.js";
import { nextTabStop, tabStopsPt } from "./tab-stops.js";
import { twipsToPoints } from "./units.js";
import type { Column } from "./columns.js";
import { fitLine, fitMark, type FitLineInput, type LineSlot, type WrapBand } from "./wrapping.js";

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
  /**
   * How much of the line has to fit on the page for the line to stay on it.
   *
   * **The room a multiple opens below the text hangs past the foot rather than
   * moving the line on.** Measured on 2026-08-11 by the authored `twip-grid`
   * document, whose four cases stack lines under a multiple until the page runs
   * out: one keeps 39 lines whose boxes come to 720.46 in a body of 720, its 39th
   * ending 8.8pt above the foot with the multiple's own room hanging past it.
   *
   * A rule stating an exact height is the other way round and answers for the
   * whole of what it asks for, since there the room is a slot the text is dropped
   * into rather than room opened under a line that measured itself.
   */
  readonly fittingHeightPt: number;
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
  // What a page opening at this paragraph's text keeps above that text, which is
  // the furniture the rows it stands in put there again: the table's own top
  // border, and the margin holding a cell's text off its wall. Nought outside a
  // table, where a page opens at the first line and everything above it is left
  // behind.
  //
  // Measured on 2026-08-10 by the authored `resuming` document. A row with neither
  // resumed exactly where an ordinary paragraph did; one held 12pt off the top of
  // its cell resumed 12pt below the top of the body, and one inside a 3pt border
  // 3.12pt below it.
  readonly resumesUnderPt: number;
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
  // The paragraphs laid out in this cell itself, which is how an object anchored
  // in one finds the frame Word places it against. A paragraph inside a table
  // inside this cell belongs to that table's cell and not to this one.
  readonly holds: readonly number[];
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
/**
 * An object standing in the way of text, as the break pass needs it: how far down the
 * column it reaches, the paragraph that anchors it, and the band itself.
 *
 * **The band is carried so that the break pass can tell whose page it belongs to.**
 * Measuring is one column with no pages in it, so a band drawn from an object near the
 * foot of one page goes on keeping text off lines that the break has since moved to
 * the next: `1bd495dddcb2` opened its second page at 314.25, a band's own right edge,
 * where Word opens it at the frame's own left. The rule about which page a line lands
 * on belongs in the break pass, and this is what it will need. **Nothing reads the
 * band yet.**
 *
 * `measureParagraph` already lets go of every object it is standing under when a
 * paragraph opens a page of its own, which is the same fault caught in the one place
 * the measure pass can see it.
 *
 * When the break pass does use this, it will be able to move a line and not to break
 * it again: **a line already broken early by a band that turns out to belong to
 * another page keeps its early break**, since breaking is a measuring question, and
 * the corpus will price whether that matters.
 *
 * Two things to hold on to while writing that. An object hanging past the foot of the
 * text is drawn up in the drawing pass, so the band here is drawn from where its
 * anchor put it rather than from where it lands. And the fault has a reverse: a band
 * whose object lands on a later page covering positions that belong to an earlier one.
 */
export type AnchoredObject = {
  readonly topPt: number;
  readonly bottomPt: number;
  readonly anchoredAt: number;
  readonly band: WrapBand;
};

export type StackMeasurement =
  | {
      readonly kind: "measured";
      readonly boxes: readonly ParagraphBox[];
      readonly cells: readonly PlacedCell[];
      readonly untornRows: readonly UntornRow[];
      readonly anchoredObjects: readonly AnchoredObject[];
      readonly heightPt: number;
      // What is left of the last paragraph where a column was asked to keep only the
      // first of its lines, which the column after it breaks again at its own width.
      readonly rest?: LineFlow | null;
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
  /**
   * How much of its own page is left below where the column run opening at this
   * paragraph begins, for the runs the caller has found that out about.
   *
   * **A column run is the one thing in a stack that has to know where the pages fell**,
   * and measuring is one column with no pages in it, so it cannot know: the run's own
   * section opens a page, but in the stack it sits whereever the text above it ended.
   * The caller breaks the stack, sees which page each run really opened on, and hands
   * that back here to be measured again. See `layOutDocument`, which does the passes.
   *
   * Unanswered, a run is measured as though it stood at the top of a page, which is what
   * the first pass has to assume and what a run under a page-opening break turns out to
   * be.
   */
  readonly roomForRun?: (opensAt: number) => number | undefined;
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
  readonly inTable: InTable | null;
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
    inTable: null,
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
  division: Division = WHOLE,
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
  // A column takes up the paragraph its own first block began and cuts the last of
  // them wherever its room ran out, so a division stands at the two ends of the run
  // of blocks and nowhere between.
  const dividing = (at: number): Division => ({
    resume: at === 0 ? division.resume : null,
    keepLines: at === blocks.length - 1 ? division.keepLines : null,
  });
  let rest: LineFlow | null = null;

  for (const [at, block] of blocks.entries()) {
    if (at <= laidOutTo) continue;
    const frame = context.inCell ? storyFrame : (context.frameOf?.(block) ?? storyFrame);

    const columns = context.inCell ? [] : (context.columnsOf?.(block) ?? []);
    if (columns.length > 1) {
      const run = columnRunFrom(blocks, at, context);
      const pageHeightPt = context.bodyHeightPt ?? Number.POSITIVE_INFINITY;
      const opensAt = blockParagraphs(run.blocks)[0]?.index;
      const roomLeftPt =
        (opensAt === undefined ? undefined : context.roomForRun?.(opensAt)) ?? pageHeightPt;
      const measured = measureColumnRun(
        run.blocks,
        context,
        top,
        columns,
        roomLeftPt,
        pageHeightPt,
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
        closesACellUnderATable:
          context.inCell && at === blocks.length - 1 && blocks[at - 1]?.kind === "table",
        opensWhatHoldsIt: at === 0,
        closesWhatHoldsIt: at === blocks.length - 1,
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
          band,
        });
      }
      const measured = measureParagraph(
        paragraph,
        context,
        top,
        frame,
        neighbours,
        { bands, ahead: [], anchorTopPt },
        dividing(at),
      );
      if (measured.kind === "blocked") return measured;

      standing = bands;
      anchoredAtPt = null;
      let box = measured.box;
      rest = measured.rest;

      // Where the anchor rounds down, the object the next paragraph holds stands
      // over the foot of this one, whose last line is then blocked and falls past
      // it. The object itself keeps the place the flow gave it, so the room the
      // line had is left empty.
      const ahead = lookedAhead(neighbours.below, top + box.heightPt, context);
      if (ahead.length > 0) {
        const again = measureParagraph(
          paragraph,
          context,
          top,
          frame,
          neighbours,
          { bands, ahead, anchorTopPt },
          dividing(at),
        );
        if (again.kind === "blocked") return again;
        anchoredAtPt = top + box.heightPt;
        box = again.box;
        rest = again.rest;
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
      // A row of the table that will not be torn is offered to no page break, where a
      // row of a table in the flow is. The break answers such a row by opening a page at the
      // row's own top, and the top of a table keeping no room in the flow is nowhere
      // in the flow to open one: what follows the table stands where the table never
      // was, and would be carried down with it. Nothing has asked Word whether it
      // tears a `w:cantSplit` row that stands out of the flow.
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
    rest,
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
  roomLeftPt: number,
  pageHeightPt: number,
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

  const boxes: ParagraphBox[] = [];
  const cells: PlacedCell[] = [];
  const untornRows: UntornRow[] = [];

  let from = 0;
  let stretchTopPt = topPt;
  let roomPt = roomLeftPt;
  // The paragraph the columns of the page above were cut through, which the first
  // column of this page carries on: Word tears a run across a page inside a paragraph
  // as readily as between two of them.
  let carried: LineFlow | null = null;

  // The walk always moves on, since a column keeps whatever block it opens with however
  // tall that is, so a stretch that swallows nothing cannot happen.
  for (;;) {
    const rest = blocks.slice(from);
    const forcedHere = new Set([...forced].flatMap((at) => (at > from ? [at - from] : [])));
    const stretch = fillStretch(
      rest,
      inColumn,
      stretchTopPt,
      columns,
      forcedHere,
      roomPt,
      balanced,
      carried,
    );
    if (stretch.kind === "blocked") return stretch;

    boxes.push(
      ...(stretchTopPt === topPt ? stretch.measured.boxes : takenUpUnder(stretch, stretchTopPt)),
    );
    cells.push(...stretch.measured.cells);
    untornRows.push(...stretch.measured.untornRows);

    if (stretch.whole) {
      return {
        kind: "measured",
        boxes,
        cells,
        untornRows,
        // Offered to no page break, for the reason `fillColumns` records.
        anchoredObjects: [],
        // What stands under the run stands under the last stretch of it, and under the
        // tallest column of that stretch: Word started the section below a balanced run
        // under the taller of its two columns.
        heightPt: stretchTopPt - topPt + stretch.measured.heightPt,
      };
    }

    // What the room would not hold carries on in the columns of the next page, which
    // begins in the stack where this page's own room ran out. Asked of Word directly on
    // 2026-08-12: seventy lines in two columns filled both columns of their page outright
    // and the last ten went on to the first column of the next.
    from += stretch.consumed;
    carried = stretch.rest;
    stretchTopPt += roomPt;
    roomPt = pageHeightPt;
  }
}

/**
 * The stretch of a run that carries on at the top of the next page, with the room the
 * paragraph opening it asks for above itself put back.
 *
 * **A page a run carries on to opens at the run and not at a paragraph.** A page an
 * ordinary break opens leaves the space above the paragraph behind, since it opens at
 * that paragraph's first line, and `breakStack` reads every page that way. A run torn
 * across a page has its columns' own top at the top of the body, so the space the first
 * paragraph asks for stands under it and is drawn.
 *
 * Read on 2026-08-17 off `395ea6c2f664`, whose run crosses onto its second page: Word
 * draws the first line of that page 4.8pt below the body's top, which is exactly the
 * space that paragraph asks for, and this drew it on the top itself. Only the box the
 * page opens at is answered for, which is the first of the stretch: no other box of it
 * can reach the foot of the page above, since a column keeps no more than the room.
 */
function takenUpUnder(stretch: Filled, stretchTopPt: number): readonly ParagraphBox[] {
  const boxes = stretch.measured.boxes;
  const first = boxes[0];
  if (first === undefined) return boxes;
  const underPt = (first.lines[0]?.topPt ?? first.topPt) - stretchTopPt;
  if (underPt <= EPSILON) return boxes;
  return [{ ...first, resumesUnderPt: first.resumesUnderPt + underPt }, ...boxes.slice(1)];
}

// One page's worth of a column run: the blocks the room held, dealt into the columns, and
// how many of them that was.
//
// **Only the stretch that holds the rest of the run is evened out.** Word fills every page
// a balanced run crosses and evens the remainder on the page it ran on to, which is the
// same measurement: three lines in each column at the foot of the page it filled, one in
// each at the top of the next.
function fillStretch(
  blocks: readonly Block[],
  context: Context,
  topPt: number,
  columns: readonly Column[],
  forced: ReadonlySet<number>,
  roomPt: number,
  balanced: boolean,
  resume: LineFlow | null,
): ColumnFill {
  const filled = fillColumns(blocks, context, topPt, columns, forced, roomPt, roomPt, resume);
  if (filled.kind === "blocked" || !filled.whole || !balanced) return filled;

  // **A run that states a break of its own is not evened out at all.** The document has
  // said where its columns divide, and Word takes it at its word: measured on 2026-08-17
  // by `column-room-probe`, five cases three times each, against the same six blocks that
  // come out three and three where nothing is stated. Six with a break after the second
  // come out **2 and 4**, with a break after the fourth **4 and 2**, and eight over three
  // columns with a break after the second come out **2 and 6 with the third column
  // empty**. Evening those out is what put a corpus template's eleven lines of a first
  // column into three, and everything under them on the wrong page.
  if (forced.size > 0) return filled;

  // Evening the columns out is the same fill asked of a shorter column, and the heights
  // worth asking for are the ones the run's own blocks make. Nothing taller than the room
  // is worth asking for, since the fill above already came in under it.
  const candidates = [...new Set(filled.bottomsPt)]
    .filter((each) => each <= roomPt + EPSILON)
    .sort((one, other) => one - other);
  let evenest: Filled | null = null;
  for (const candidate of candidates) {
    const evened = fillColumns(blocks, context, topPt, columns, forced, candidate, roomPt, resume);
    if (evened.kind === "blocked") return evened;
    if (!evened.whole) continue;
    if (evenest === null || evensBetter(evened, evenest)) evenest = evened;
  }
  // A run the room will not hold whole fills what it has and carries the rest into the
  // columns of the next page, which is what every page of it but the last does.
  return evenest ?? filled;
}

/**
 * Which of two divisions of a run Word settles on.
 *
 * **The tallest column is made as short as it can be, and where two divisions stand
 * equally tall the earlier column takes more.** Measured on 2026-08-13 by the authored
 * cases A to K, each read off Word's own pdf by where the paragraph under the run came to
 * sit: four one-line blocks in two columns cost 48 and not 72, five come out three and two
 * rather than two and three, and four blocks in three columns leave the third empty.
 *
 * **A fill is judged by the height it comes out at and not by the height it was asked
 * for.** Case F, four blocks in columns of 108pt and 234pt: Word divided it two and two,
 * its last column coming out at 72 against a first column of 48, where cutting every
 * column at one height can only reach 96 and 48. A corpus run of 109.35pt and 230.55pt
 * says the same, its columns coming out 23.8 and 24.3.
 *
 * **The division is one part in as many columns as there are, and not in proportion to
 * how wide they are.** Case K: five one-line blocks in columns of 108pt and 234pt came out
 * three in the narrow column and two in the wide, where a share of the width would have
 * put two in the narrow.
 *
 * A division is judged by how tall its columns stand and not by what the run then costs
 * the page, which is a different quantity: see `drawnHeightPt`.
 */
function evensBetter(one: Filled, than: Filled): boolean {
  const tallest = (fill: Filled): number => Math.max(0, ...fill.columnHeightsPt);
  if (Math.abs(tallest(one) - tallest(than)) > EPSILON) return tallest(one) < tallest(than);

  for (const [at, heightPt] of one.columnHeightsPt.entries()) {
    const other = than.columnHeightsPt[at] ?? 0;
    if (Math.abs(heightPt - other) > EPSILON) return heightPt > other;
  }
  return false;
}

/**
 * What the run costs the page it stands on: **the tallest of its columns up to and
 * including the last one that draws anything. The columns past that cost nothing,
 * however much they hold.**
 *
 * Measured on 2026-08-13 by the authored cases H and I: three lines of 24pt followed by
 * three empty paragraphs of 30pt, in two equal columns, put the lines in the first column
 * and the empties in the second, and the run cost **72**, the first column alone. The same
 * with 26pt empties cost 72 as well, so the second column's 90 and its 78 were both worth
 * nothing. **Case J is the control**: put a word in each of those three paragraphs and
 * they cost 90, drawn in the second column at 60, 90 and 120. The only thing that changed
 * is that they draw something.
 *
 * **A column standing before one that draws is worth its whole height, drawn or not**,
 * which is what says this is about the columns after the last drawn thing and not about
 * empty columns anywhere. The corpus run `28ef8aa08b34` carries a stretch of three columns
 * whose first two hold nothing but empty paragraphs, 28.17 and 27.60, and whose third
 * draws one line and stands 23.85: Word put the paragraph under it at 28.17, the tallest
 * of the three.
 *
 * **A column that draws anything is measured to the foot of everything in it**, its own
 * empty paragraphs included. Case G: four empty paragraphs falling in the column that also
 * drew three blocks counted for their 96, and the run cost 168.
 *
 * Case E said the same of a column that receives nothing at all, and B could not tell the
 * difference, since its empties came to exactly the 72 its drawn column already cost.
 *
 * Where no column draws anything the run keeps the room its tallest column takes, which is
 * what this did before any of it was measured and what nothing has asked Word about.
 */
function drawnHeightPt(
  columnHeightsPt: readonly number[],
  columnsDrawing: readonly boolean[],
): number {
  const lastDrawing = columnsDrawing.lastIndexOf(true);
  if (lastDrawing < 0) return Math.max(0, ...columnHeightsPt);
  return Math.max(0, ...columnHeightsPt.slice(0, lastDrawing + 1));
}

// Whether anything of the paragraph is drawn where it stands: its text, the number a
// list puts in front of it, or the borders and shading behind it. A paragraph holding
// none of those is drawn nowhere at all.
const drawsSomething = (box: ParagraphBox): boolean =>
  box.lines.length > 0 || box.marker !== null || box.paint !== null;

type Filled = {
  readonly kind: "filled";
  readonly measured: Extract<StackMeasurement, { readonly kind: "measured" }>;
  // Whether every block handed in found a column, which is what says a height is
  // tall enough to be worth evening out at.
  readonly whole: boolean;
  // How many of them did, which is where the next page's columns carry on from.
  readonly consumed: number;
  // What is left of the paragraph the last column cut, which the columns of the next
  // page take up.
  readonly rest: LineFlow | null;
  // How far below the run's top each block reached, which are the heights a
  // column could be cut to.
  readonly bottomsPt: readonly number[];
  // How tall each column came out, which is what one fill is judged against another by,
  // and whether each of them draws anything, which is what says what the run costs.
  readonly columnHeightsPt: readonly number[];
  readonly columnsDrawing: readonly boolean[];
};

type ColumnFill = { readonly kind: "blocked"; readonly blocker: LayoutBlocker } | Filled;

// The blocks dealt into the columns, each column measured in its own frame and every
// one of them starting at the run's own top. A block that will not fit in the room
// left opens the next column, and one carrying a break of its own opens one whatever
// room is left.
//
// **The last column takes what is left, up to the room, rather than being cut at the
// height the others are evened to.** Word's own division of the four-block run read on
// 2026-08-14 has its last column standing 0.5pt taller than the height the first two were
// evened to, which no fill cutting every column at that height can reach.
//
// **What follows the run's last break of its own stays in the column that break opened.**
// Measured on 2026-08-13 by the authored case G: two blocks, a break, three blocks and
// four empty paragraphs in three columns came out two in the first column, the other seven
// in the second, nothing in the third, and the run as tall as that second column. Once a
// break has said where the text goes, nothing after it is divided again.
function fillColumns(
  blocks: readonly Block[],
  context: Context,
  topPt: number,
  columns: readonly Column[],
  forced: ReadonlySet<number>,
  heightPt: number,
  roomPt: number,
  resume: LineFlow | null,
): ColumnFill {
  const blockOf = new Map<number, number>();
  for (const [at, block] of blocks.entries()) {
    for (const paragraph of blockParagraphs([block])) blockOf.set(paragraph.index, at);
  }
  const lastBreak = forced.size === 0 ? null : Math.max(...forced);
  // Only a paragraph is cut between two columns. A table is torn at a row and a row at
  // a line, which is a rule about pages that nothing has asked Word about for columns.
  const divides = (at: number): boolean => blocks[at]?.kind === "paragraph";

  const boxes: ParagraphBox[] = [];
  const cells: PlacedCell[] = [];
  const untornRows: UntornRow[] = [];
  const bottomsPt: number[] = [];
  const columnHeightsPt: number[] = [];
  const columnsDrawing: boolean[] = [];
  let from = 0;
  // What the column above left of the paragraph the run carries on with, which the
  // column being filled breaks again at its own width.
  let carried = resume;

  for (const [at, column] of columns.entries()) {
    const afterTheLastBreak = lastBreak !== null && from >= lastBreak;
    if (from >= blocks.length) {
      columnHeightsPt.push(0);
      columnsDrawing.push(false);
      continue;
    }
    const measured = measureBlocks(blocks.slice(from), context, topPt, column, {
      resume: carried,
      keepLines: null,
    });
    if (measured.kind === "blocked") return measured;

    // **No column keeps more than the room its page left.** A column used to keep
    // whatever was over, which is what put the tail of a run below the foot of its page
    // and left `breakStack` to cut inside the run; what is over goes to the next page's
    // columns instead. The last column is held to the room alone, since it takes what the
    // columns evened out above it did not, and so is the column the run's last break
    // opened.
    const cutAtPt = at === columns.length - 1 || afterTheLastBreak ? roomPt : heightPt;
    const cut = cutColumnAt(
      measured.boxes,
      blockOf,
      forced,
      divides,
      topPt,
      cutAtPt,
      cutAtPt >= roomPt - EPSILON,
    );
    for (const box of measured.boxes) reachesOf(box, topPt, bottomsPt);

    if (cut.at >= blocks.length) {
      boxes.push(...measured.boxes);
      cells.push(...measured.cells);
      untornRows.push(...measured.untornRows);
      columnHeightsPt.push(measured.heightPt);
      columnsDrawing.push(measured.boxes.some(drawsSomething));
      from = blocks.length;
      carried = null;
      continue;
    }

    // The paragraph the cut runs through is measured with this column, which keeps the
    // lines it had room for, and stays where the next column takes up.
    const upTo = cut.keepLines === null ? cut.at : cut.at + 1;
    const kept = measureBlocks(blocks.slice(from, upTo), context, topPt, column, {
      resume: carried,
      keepLines: cut.keepLines,
    });
    if (kept.kind === "blocked") return kept;
    boxes.push(...kept.boxes);
    cells.push(...kept.cells);
    untornRows.push(...kept.untornRows);
    // A column ending inside a paragraph leaves no break's line behind it: the break
    // that paragraph carries is what opened this column, and was answered for there.
    columnHeightsPt.push(
      kept.heightPt +
        (cut.keepLines === null ? nearSideOfABreak(measured.boxes, blockOf, forced, cut.at) : 0),
    );
    columnsDrawing.push(kept.boxes.some(drawsSomething));
    from = cut.at;
    carried = kept.rest ?? null;
  }

  return {
    kind: "filled",
    measured: {
      kind: "measured",
      boxes,
      cells,
      untornRows,
      // **An object a column's text wraps round is offered to no page break.** How
      // much of a run a page holds is settled here, by `cutColumnAt` against the room
      // the page left, so `breakStack` never cuts inside a run and an object moving
      // its paragraph to the page below would have to be answered by the column
      // holding it. A row that will not be torn reaches the break because a run
      // carries on into the columns of the next page whole blocks at a time, which is
      // where such a row can still move. Nothing has asked Word what it does with an
      // object a column has not the room for.
      anchoredObjects: [],
      heightPt: drawnHeightPt(columnHeightsPt, columnsDrawing),
    },
    whole: from >= blocks.length,
    consumed: from,
    rest: carried,
    bottomsPt,
    columnHeightsPt,
    columnsDrawing,
  };
}

// How far below the run's top a block reaches, and how far each of its lines does: the
// heights a column could be cut at, which are places inside a paragraph as well as
// between two of them.
function reachesOf(box: ParagraphBox, topPt: number, into: number[]): void {
  into.push(box.topPt + box.heightPt - topPt);
  if (box.lines.length < 2) return;
  for (const line of box.lines) into.push(line.topPt + line.fittingHeightPt - topPt);
}

/**
 * The room the paragraph carrying a column break leaves at the foot of the column it
 * breaks out of.
 *
 * **A paragraph stands on both sides of its own break**: an empty line closing the
 * column it leaves, and what follows the break opening the next. Measured on
 * 2026-08-13 by the authored cases L and M, whose runs cost 72 where the column the
 * break opens holds 48: a paragraph holding nothing but the break and one carrying its
 * own text after it both leave a line of their own height behind them.
 *
 * The corpus said it first, in tenths of a point: `299724db7cc1` divides a three-column
 * run with two such paragraphs, half a point tall apiece, and the paragraph under the
 * run stands 0.39pt lower in Word than it did here. That 0.39 carries the last line of
 * its page from a fifth of a point inside the body to a fifth past it, which is where
 * Word puts it on the next page and this did not.
 *
 * It is one line and not the whole paragraph: what follows the break may run to several
 * lines in the column it opens, and only the empty piece before it stays behind.
 */
function nearSideOfABreak(
  boxes: readonly ParagraphBox[],
  blockOf: ReadonlyMap<number, number>,
  forced: ReadonlySet<number>,
  cut: number,
): number {
  if (!forced.has(cut)) return 0;
  const box = boxes.find((each) => blockOf.get(each.index) === cut);
  if (box === undefined) return 0;
  return box.lines[0]?.heightPt ?? box.heightPt;
}

// Where a column ends: the first block of the run that belongs in the next one, and,
// where the cut falls inside a paragraph, how many of that paragraph's lines this
// column keeps.
type ColumnCut = {
  readonly at: number;
  readonly keepLines: number | null;
};

const WHOLE_COLUMN: ColumnCut = { at: Number.POSITIVE_INFINITY, keepLines: null };

// The first block of the run that belongs in the next column: the one carrying a
// break of its own, or the one whose foot passed the height the column was given. A
// paragraph the room ran out inside is cut at the last of its lines that fitted rather
// than moved whole; a block the column opens with and cannot cut stays there however
// tall it is, since a column that carries it forward whole never ends.
//
// **A foot is what the paragraph draws, and not the room it keeps under itself**, which
// is the rule `breakStack` already holds a page to. `2c1289b95c31` ends the second column
// of its first page with an empty paragraph standing 13.80 in a column with 15.92 left,
// which asks for 3.15 more after itself: judged whole it passed the foot by a point and
// went to the next page, taking its anchored drawing with it and leaving every line of
// that page 7pt low.
function cutColumnAt(
  boxes: readonly ParagraphBox[],
  blockOf: ReadonlyMap<number, number>,
  forced: ReadonlySet<number>,
  divides: (at: number) => boolean,
  topPt: number,
  heightPt: number,
  askedForByTheRoom: boolean,
): ColumnCut {
  let opened: number | null = null;
  for (const box of boxes) {
    const place = blockOf.get(box.index);
    if (place === undefined) continue;
    if (opened === null) opened = place;
    if (place !== opened && forced.has(place)) return { at: place, keepLines: null };
    if (box.contentBottomPt <= topPt + heightPt + EPSILON) continue;

    const kept = divides(place)
      ? linesKeptOf(box, topPt + heightPt, place === opened, askedForByTheRoom)
      : null;
    if (kept !== null) return { at: place, keepLines: kept };
    if (place === opened) continue;
    return { at: place, keepLines: null };
  }
  return WHOLE_COLUMN;
}

/**
 * How many of a paragraph's lines stand inside the column, where that is a cut and not the
 * whole paragraph either way.
 *
 * Nothing where every line fits and only the room the paragraph keeps under itself passed
 * the foot, and nothing where not even the first line does, unless the column opens with
 * it and has nowhere to send it.
 *
 * **A cut made to even a run out leaves at least two lines behind it; one the room forces
 * leaves whatever it had room for.** Read on 2026-08-17 off three corpus documents:
 *
 * - `8010f77cdeee` divides a run whose fifth block is two lines. Cut between them the
 *   columns stand 82.9 and 88.0, and left whole they stand 97.6 and 73.4. **Word draws the
 *   taller**, so it will not leave one line of a paragraph at the foot of a column to even
 *   the run out, even where doing so would make the run shorter.
 * - `395ea6c2f664` divides a three-line paragraph **two and one**, standing 81.8 and 102.8
 *   against the 107.4 and 77.2 it comes to whole, and Word draws the cut one. So it is the
 *   near side of the cut that has to hold two lines and not the far side.
 * - `52342f52bfb1` ends a column of its first page with **one line** of the paragraph its
 *   next page carries on. That column ran out of page rather than being evened to a
 *   height, which is the difference: a run torn across a page takes what the room holds.
 */
function linesKeptOf(
  box: ParagraphBox,
  footPt: number,
  opensTheColumn: boolean,
  askedForByTheRoom: boolean,
): number | null {
  if (box.lines.length < 2) return null;
  let kept = 0;
  for (const line of box.lines) {
    if (line.topPt + line.fittingHeightPt > footPt + EPSILON) break;
    kept += 1;
  }
  if (kept >= box.lines.length) return null;
  if (kept >= 2) return kept;
  // **A paragraph that opens the column is cut all the same**, since the line the rule
  // would send forward is the only thing the column has: refusing the cut leaves the
  // column empty and the run undivided. `2c1289b95c31` ends its two-column section with
  // one paragraph of two lines and Word draws them side by side, one to a column, where
  // `8010f77cdeee` had blocks above its cut paragraph and Word moved that paragraph whole
  // rather than leave a line of it behind. Neither document asks for widow control, so
  // this is Word's own arithmetic and not `w:widowControl`.
  if (kept === 1 && opensTheColumn) return 1;
  if (!askedForByTheRoom) return null;
  // A column with room for less than a line of the paragraph it opens with keeps a line
  // of it all the same, since a column that carries the whole of it forward never ends.
  return kept === 1 || opensTheColumn ? 1 : null;
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
    resolveParagraphFrame(paragraph, context.styles, context.inTable).pageBreakBefore);

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
  // A width needs no bands: every cell of the first row is as wide as its columns
  // however the style formats it.
  const plans = planCells(block, { rowBandSize: 1, columnBandSize: 1 }, frame);
  const first = plans[0] ?? [];
  return first.reduce((width, plan) => width + plan.widthPt, 0);
}

// Where one cell stands on the table, before anything knows how tall a row is.
type CellPlan = {
  readonly cell: TableCell;
  readonly at: number;
  readonly leftPt: number;
  readonly widthPt: number;
  // Where the cell stands in the table, which is what the style's conditional
  // formats are chosen by.
  readonly position: CellPosition;
  // The last row the cell reaches, which is its own unless a merge opens at it. A
  // swallowed cell reaches nowhere: it draws nothing and is worth nothing.
  readonly throughRow: number | null;
};

// **A cell spanning grid columns is worth the columns under it whether it says so or
// not**, and the cell after it starts at the column the span ended on rather than at
// the next one along. Measured on 2026-08-10 by the authored `merged-cells`
// document, whose cases h and i put a third column's text at 180 either way.
//
// **A `w:vMerge` continuation draws nothing and is worth nothing.** The same
// document's case g puts a 40pt line in every swallowed cell of a table of 20pt
// rows: Word drew no ink for any of them and left every row at 20pt. A continuation
// with no merge open above it is a cell of its own, which is what an orphan can only
// be.
// Which of the table's conditional formats a cell standing here answers to. A
// switch the table's own `w:tblLook` turns off takes the row or the column out of the
// question altogether, and a table that turns banding off has no bands at all.
//
// **Banding counts from the first row the header did not take**, so a table whose
// first row is its header has its second row in band 1.
function positionOf(
  block: Table,
  bands: BandSizes,
  rowAt: number,
  gridAt: number,
  span: number,
): CellPosition {
  const look = block.look;
  const columns = Math.max(block.gridTwips.length, ...block.rows.map((row) => row.cells.length));
  const lastRowAt = block.rows.length - 1;

  const firstRow = look.firstRow && rowAt === 0;
  const lastRow = look.lastRow && rowAt === lastRowAt && lastRowAt > 0;
  const firstColumn = look.firstColumn && gridAt === 0;
  const lastColumn = look.lastColumn && gridAt + span >= columns && columns > 1;

  const bandOf = (index: number, size: number): 1 | 2 | null =>
    index < 0 ? null : Math.floor(index / size) % 2 === 0 ? 1 : 2;

  return {
    firstRow,
    lastRow,
    firstColumn,
    lastColumn,
    rowBand:
      !look.horizontalBanding || firstRow || lastRow
        ? null
        : bandOf(rowAt - (look.firstRow ? 1 : 0), bands.rowBandSize),
    columnBand:
      !look.verticalBanding || firstColumn || lastColumn
        ? null
        : bandOf(gridAt - (look.firstColumn ? 1 : 0), bands.columnBandSize),
  };
}

// A band is as deep as the table style says unless the table says otherwise, and
// one row and one column deep where neither does.
type BandSizes = { readonly rowBandSize: number; readonly columnBandSize: number };

function planCells(block: Table, bands: BandSizes, frame: Frame): readonly (readonly CellPlan[])[] {
  // A merge learns how far down it reached only once the rows under it have been
  // read, so the cell that opened one is written to as they are.
  type Planning = { -readonly [K in keyof CellPlan]: CellPlan[K] };

  const plans: Planning[][] = [];
  const open = new Map<number, Planning>();

  for (const [rowAt, row] of block.rows.entries()) {
    const planned: Planning[] = [];
    // The columns a merge still reaches past this row: the ones it was continued in,
    // and the ones it opened at.
    const alive = new Set<number>();
    let gridAt = 0;
    let leftPt = frame.leftPt;

    for (const cell of row.cells) {
      const widthPt = plannedWidthPt(block, cell, gridAt) ?? frame.widthPt;
      const opened = cell.merge === "continue" ? open.get(gridAt) : undefined;
      const position = positionOf(block, bands, rowAt, gridAt, cell.gridSpan);
      if (opened === undefined) {
        const plan: Planning = { cell, at: gridAt, leftPt, widthPt, position, throughRow: rowAt };
        planned.push(plan);
        if (cell.merge === "restart") {
          open.set(gridAt, plan);
          alive.add(gridAt);
        } else open.delete(gridAt);
      } else {
        opened.throughRow = rowAt;
        alive.add(gridAt);
        planned.push({ cell, at: gridAt, leftPt, widthPt, position, throughRow: null });
      }
      gridAt += cell.gridSpan;
      leftPt += widthPt;
    }

    for (const column of [...open.keys()]) if (!alive.has(column)) open.delete(column);
    plans.push(planned);
  }

  return plans;
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
  const insets = resolveTableInsets(context.styles, block.styleId, block.statedInsets);

  const first = borders[0] ?? [];
  const last = borders[borders.length - 1] ?? [];
  const outerTopPt = halfOf(first.map((cell) => cell.agreed.top));
  const outerBottomPt = halfOf(last.map((cell) => cell.agreed.bottom));
  const outerLeftPt = halfOf(borders.map((row) => row[0]?.drawn.left ?? null));

  // An old document's indent is measured to the text rather than to the table, so
  // the first column's own margin stands outside the indent instead of inside it
  // and the table's edge moves left by the whole of it.
  //
  // **A table inside a cell stands where a modern document's does either way.**
  // Measured on 2026-08-10 by the authored `resuming` document and its legacy twin:
  // the two put the outer table's text 5.28pt apart and the table inside a cell in
  // the same place, 6.96pt inside the cell's own text in both. A nested table
  // states no indent of its own there, which is what a document in the wild writes.
  const openingCell = block.rows[0]?.cells[0];
  const insetPt =
    measuresTheIndentToTheText(context.settings) && !context.inCell && openingCell !== undefined
      ? -leftMarginOf(openingCell, insets, first[0]?.drawn ?? NO_BORDERS)
      : outerLeftPt;

  const rowFrame = {
    leftPt: frame.leftPt + twipsToPoints(insets.indentTwips) + insetPt,
    widthPt: frame.widthPt,
  };

  // Everything inside the table reads the table's own style, which the paragraphs
  // in its cells sit under.
  const inTable: Context = { ...context, inTable: { styleId: block.styleId, at: null } };

  const stated = resolveBandSizes(context.styles, block.styleId);
  const plans = planCells(
    block,
    {
      rowBandSize: block.look.rowBandSize ?? stated.rowBandSize ?? 1,
      columnBandSize: block.look.columnBandSize ?? stated.columnBandSize ?? 1,
    },
    rowFrame,
  );
  const measured: MeasuredCell[][] = [];
  const margins: RowMargins[] = [];
  for (const [at, row] of block.rows.entries()) {
    const of = measureRowCells(row, plans[at] ?? [], borders[at] ?? [], inTable, insets);
    if (of.kind === "blocked") return of;
    measured.push([...of.cells]);
    margins.push(of.margins);
  }

  const heightsPt = rowHeights(block, measured, margins);
  const placed = placeRows(block, measured, margins, heightsPt, topPt + outerTopPt, outerTopPt);

  // A cell is measured with no bands at all, so nothing inside a table can anchor
  // an object a page break has to make room for.
  return {
    kind: "measured",
    boxes: placed.boxes,
    cells: placed.cells,
    untornRows: placed.untornRows,
    anchoredObjects: [],
    heightPt: outerTopPt + heightsPt.reduce((total, each) => total + each, 0) + outerBottomPt,
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
  // The last row this cell reaches, which is the row it stands in unless a merge
  // opens at it.
  readonly throughRow: number;
  readonly fillColor: string | null;
  readonly borders: Borders;
};

// How far a row holds every cell in it off its own walls, which is the row's
// business rather than any one cell's.
type RowMargins = {
  readonly topPt: number;
  // The cell's own margin at the foot, kept apart from the half of the line cleared
  // after it, because a row told exactly how tall to be counts one and not the other.
  readonly bottomCellPt: number;
  readonly bottomPt: number;
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
//
// What each cell of a row holds, before anything knows how tall the row is: a cell
// merged down a run of rows is measured here and seated once all of them are.
function measureRowCells(
  row: TableRow,
  plans: readonly CellPlan[],
  borders: readonly CellBorders[],
  context: Context,
  insets: TableInsets,
):
  | {
      readonly kind: "measured";
      readonly cells: readonly MeasuredCell[];
      readonly margins: RowMargins;
    }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker } {
  const measured: MeasuredCell[] = [];
  const bottomCellPt = rowMarginPt(row, insets, "bottomTwips");
  const margins: RowMargins = {
    topPt: rowMarginPt(row, insets, "topTwips") + halfOf(borders.map((of) => of.agreed.top)),
    bottomCellPt,
    bottomPt: bottomCellPt + halfOf(borders.map((of) => of.agreed.bottom)),
  };

  // A cell is measured from its own origin and only then moved down to the row, so
  // the page coordinates a wrapping object stands in cannot reach inside one.
  const inCell: Context = { ...context, bandsFor: () => [], inCell: true };

  for (const [at, plan] of plans.entries()) {
    if (plan.throughRow === null) continue;
    const cell = plan.cell;
    const inCellHere: Context = {
      ...inCell,
      inTable: { styleId: context.inTable?.styleId ?? null, at: plan.position },
    };
    const own = borders[at]?.drawn ?? NO_BORDERS;
    const leftMarginPt = leftMarginOf(cell, insets, own);
    const rightMarginPt = Math.max(
      twipsToPoints(cell.margins.rightTwips ?? insets.rightTwips),
      halfOf([own.right]),
    );
    const cellFrame = {
      leftPt: plan.leftPt + leftMarginPt,
      widthPt: Math.max(0, plan.widthPt - leftMarginPt - rightMarginPt),
    };
    const of = measureBlocks(cell.blocks, inCellHere, 0, cellFrame);
    if (of.kind === "blocked") return of;
    measured.push({
      align: cell.verticalAlign,
      boxes: of.boxes,
      inner: of.cells,
      innerUntorn: of.untornRows,
      heightPt: of.heightPt,
      leftPt: plan.leftPt,
      widthPt: plan.widthPt,
      throughRow: plan.throughRow,
      fillColor: cell.fillColor,
      borders: own,
    });
  }

  return { kind: "measured", cells: measured, margins };
}

// How tall each row of the table came out.
//
// **What a merge is short falls whole on the last row of the merge**, and on the last
// row of the merge rather than the last row of the table. Measured on 2026-08-10 by
// the authored `merged-cells` document: six 20pt lines merged down four 20pt rows
// left the first three at 20 and made the fourth 60, ten lines made it 140, and the
// same six lines merged down only the first two rows made the second 99.84 while the
// two ordinary rows under it stayed at 20. Sharing the shortfall out would have put
// every one of those rows at the same height, and it does not.
//
// Nothing has asked Word what a merge ending on a row told exactly how tall to be
// does. It is left as that row already is: unable to grow for anything, overflowed
// by a merge the way it would be by its own text.
function rowHeights(
  block: Table,
  measured: readonly (readonly MeasuredCell[])[],
  margins: readonly RowMargins[],
): readonly number[] {
  const heightsPt = block.rows.map((row, at) => {
    const own = measured[at] ?? [];
    const contentHeightPt = Math.max(
      0,
      ...own.filter((cell) => cell.throughRow === at).map((cell) => cell.heightPt),
    );
    return rowHeightPt(row, contentHeightPt, {
      marginsPt: (margins[at]?.topPt ?? 0) + (margins[at]?.bottomPt ?? 0),
      bottomCellMarginPt: margins[at]?.bottomCellPt ?? 0,
    });
  });

  // A merge is read once the rows it reaches are all as tall as their own text, and
  // the ones ending lowest are read last so that a merge above them has already had
  // whatever it was short.
  const merges = measured
    .flatMap((row, at) => row.filter((cell) => cell.throughRow > at).map((cell) => ({ cell, at })))
    .sort((left, right) => left.cell.throughRow - right.cell.throughRow);

  for (const { cell, at } of merges) {
    const through = cell.throughRow;
    if (block.rows[through]?.height?.exact === true) continue;
    const requiredPt =
      (margins[at]?.topPt ?? 0) + cell.heightPt + (margins[through]?.bottomPt ?? 0);
    let availablePt = 0;
    for (let row = at; row <= through; row += 1) availablePt += heightsPt[row] ?? 0;
    if (requiredPt > availablePt + EPSILON)
      heightsPt[through] = (heightsPt[through] ?? 0) + requiredPt - availablePt;
  }

  return heightsPt;
}

// Where everything in the table ended up, once every row knows how tall it is.
//
// **A merged cell is seated in the whole run of rows it reaches** rather than in the
// one it opens. The same document's case f centres one 20pt line in a merge of four
// 20pt rows, and Word drew it 30pt down: half of the 60pt the run had over it.
function placeRows(
  block: Table,
  measured: readonly (readonly MeasuredCell[])[],
  margins: readonly RowMargins[],
  heightsPt: readonly number[],
  topPt: number,
  outerTopPt: number,
): {
  readonly boxes: readonly ParagraphBox[];
  readonly cells: readonly PlacedCell[];
  readonly untornRows: readonly UntornRow[];
} {
  const topsPt: number[] = [];
  let running = topPt;
  for (const heightPt of heightsPt) {
    topsPt.push(running);
    running += heightPt;
  }
  const spannedPt = (from: number, through: number): number => {
    let total = 0;
    for (let row = from; row <= through; row += 1) total += heightsPt[row] ?? 0;
    return total;
  };

  // What each row's own text asks of it, a merge reaching down into it counted
  // against the row it lands in rather than the one it opened at.
  const heldPt = block.rows.map(() => 0);
  for (const [at, row] of measured.entries())
    for (const cell of row) {
      const through = cell.throughRow;
      const askedPt =
        (margins[at]?.topPt ?? 0) +
        cell.heightPt +
        (margins[through]?.bottomPt ?? 0) -
        spannedPt(at, through - 1);
      heldPt[through] = Math.max(heldPt[through] ?? 0, askedPt);
    }

  const boxes: ParagraphBox[] = [];
  const cells: PlacedCell[] = [];
  const untornRows: UntornRow[] = [];

  for (const [at, row] of block.rows.entries()) {
    const rowTopPt = topsPt[at] ?? topPt;
    const ownHeightPt = heightsPt[at] ?? 0;
    const opened: ParagraphBox[] = [];

    for (const cell of measured[at] ?? []) {
      const through = cell.throughRow;
      const heldToPt = spannedPt(at, through);
      const topMarginPt = margins[at]?.topPt ?? 0;
      const bottomMarginPt = margins[through]?.bottomPt ?? 0;
      // A row told exactly how tall to be leaves its cells whatever room is left over
      // once it has held them off its walls, and Word draws what does not fit anyway.
      const roomPt = Math.max(0, heldToPt - topMarginPt - bottomMarginPt);
      const offset = rowTopPt + topMarginPt + seatingOffset(cell.align, roomPt, cell.heightPt);
      // Only a row given a height of its own can be shorter than what it holds, so
      // only that row has anything to cut its cells off at.
      const clipTo =
        row.height?.exact === true
          ? { leftPt: cell.leftPt, topPt: rowTopPt, widthPt: cell.widthPt, heightPt: heldToPt }
          : null;
      // What the row draws above its own text where a page opens at it, which is
      // what it drew above it on the page the table opened on: the table's own top
      // border and the margin holding the cell's text off its wall. A table inside
      // a cell adds its own to the one round it, which is what case e of `resuming`
      // asks: the outer row states neither, and the page below the tear opens on
      // the nested table's 3pt border.
      const resumesUnderPt = outerTopPt + topMarginPt;
      for (const box of cell.boxes) {
        const placed = {
          ...shiftBox(box, offset),
          clipTo,
          resumesUnderPt: box.resumesUnderPt + resumesUnderPt,
        };
        boxes.push(placed);
        opened.push(placed);
      }
      for (const inner of cell.inner) cells.push({ ...inner, topPt: inner.topPt + offset });
      for (const inner of cell.innerUntorn)
        untornRows.push({
          ...inner,
          topPt: inner.topPt + offset,
          bottomPt: inner.bottomPt + offset,
        });
      const nested = new Set(cell.inner.flatMap((inner) => inner.holds));
      cells.push({
        leftPt: cell.leftPt,
        topPt: rowTopPt,
        widthPt: cell.widthPt,
        heightPt: heldToPt,
        fillColor: cell.fillColor,
        borders: cell.borders,
        holds: cell.boxes.map((box) => box.index).filter((index) => !nested.has(index)),
      });
    }

    // Word tears an ordinary row at a line, and refuses two: one saying so, and one
    // standing taller than its own text, whose empty foot has no line to be torn at.
    // Both were measured on 2026-08-07 by the authored `tearing` document: a row
    // asking to be 150pt tall with 48pt of text in it moved whole where 102pt was
    // left, and so did the same row with 144pt of text, while a row asking to be
    // 48pt tall with 144pt of text in it was torn like any other.
    const opensAt = opened[0]?.index;
    if (opensAt !== undefined && (row.cantSplit || ownHeightPt > (heldPt[at] ?? 0) + EPSILON))
      untornRows.push({ topPt: rowTopPt, bottomPt: rowTopPt + ownHeightPt, opensAt });
  }

  return { boxes, cells, untornRows };
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
function columnWidthPt(gridTwips: readonly number[], at: number, span: number): number | null {
  const spanned = gridTwips.slice(at, at + span);
  if (spanned.length === 0) return null;
  return twipsToPoints(spanned.reduce((total, twips) => total + twips, 0));
}

/**
 * How wide a cell is planned, which is the grid where the table states fixed columns
 * and the cell's own `w:tcW` where it does not.
 *
 * **A table stating `w:tblLayout w:type="fixed"` lays its columns out on the grid
 * and its cells' stated widths are ignored.** Read on 2026-08-18 off Word's own pdf.
 * Two tables in three corpus documents state the two at once: a grid of
 * 111+2093+104+113+2160+103 twips and two cells of `w:gridSpan="3"` stating `w:tcW`
 * 2318 and 2385, which come to 4703 against the grid's 4684. Word draws their
 * shading between borders at 340.80 and 456.24 on a page 0.48pt of border wide, so
 * the cells run 340.56 to 456.00 to 574.80: **115.44 and 118.80 against the grid's
 * 115.40 and 118.80**, where the stated widths would give 115.90 and 119.25. Four
 * pictures anchored in the second cell say the same from the other side, landing a
 * flat 115.39 right of the ones in the first.
 *
 * Nothing is changed for a table that states no layout. **Five of those disagree
 * with their own grid as well, by as much as 5.55pt down a row**, and what Word does
 * with an autofit table is a rule of its own that has not been asked.
 */
function plannedWidthPt(block: Table, cell: TableCell, gridAt: number): number | null {
  const fromGrid = columnWidthPt(block.gridTwips, gridAt, cell.gridSpan);
  return block.fixedColumns ? (fromGrid ?? cellWidthPt(cell)) : (cellWidthPt(cell) ?? fromGrid);
}

const cellWidthPt = (cell: TableCell): number | null =>
  cell.statedWidthTwips === null ? null : twipsToPoints(cell.statedWidthTwips);

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
  | { readonly kind: "measured"; readonly box: ParagraphBox; readonly rest: LineFlow | null }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

/**
 * A paragraph a column run cut between two of its columns: where the column takes it up,
 * and how many of its lines that column keeps.
 *
 * **Word cuts a paragraph across a column boundary and breaks what is left again at the
 * width of the column it lands in.** Read on 2026-08-17 off two drawings of
 * `52342f52bfb1`, whose page one ends a 58.1pt column with the first line of a paragraph
 * and lays the rest of it out at page two's 115.75pt column, where it comes to two lines
 * rather than the three it takes here. The line widths either side of the cut agree with
 * Word to a tenth of a point. `column-room-probe` cases E and F say the same of an
 * authored document: a paragraph of twelve words in a 109pt column before a 231pt one
 * comes out two lines and one, and the run takes the 48pt its tallest column stands.
 *
 * **A piece is not a paragraph.** The room the paragraph asks for above itself stands
 * above the first piece and the room below it under the last, its number is drawn in
 * front of the piece holding its first line, and its mark sits at the end of the piece
 * holding its last. That is the same division `partOf` makes where a page break runs
 * through a paragraph, which is the one this has to agree with.
 */
type Division = {
  readonly resume: LineFlow | null;
  readonly keepLines: number | null;
};

const WHOLE: Division = { resume: null, keepLines: null };

// What stands either side of the paragraph in the same run of blocks, which is
// all "don't add space between paragraphs of the same style" asks about.
type Neighbours = {
  readonly above: Paragraph | null;
  readonly below: Paragraph | null;
  // Whether this is the paragraph a cell holding a table has to end with, which is
  // the one paragraph Word leaves no room for.
  readonly closesACellUnderATable: boolean;
  // Whether nothing at all stands above or below it where it is written, which is
  // the edge an automatic space collapses against. A table on that side is a block
  // like any other and is not an edge.
  readonly opensWhatHoldsIt: boolean;
  readonly closesWhatHoldsIt: boolean;
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
  division: Division,
): ParagraphMeasurement {
  const paragraphMark = resolveParagraphMark(paragraph, context.styles, context.inTable);
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

  const paragraphFrame = resolveParagraphFrame(paragraph, context.styles, context.inTable);
  const runs = flowing(readRuns(paragraph, context.styles), context.inCell, context.inColumn);
  const insets = insetsOf(paragraphFrame);
  const resumed = division.resume !== null;
  // A number stands in front of the paragraph's first line, which is in the column
  // this piece was cut from.
  const number = resumed ? undefined : context.numbers.get(paragraph.index);
  // The break a section carries is the paragraph's own end, so it belongs to the piece
  // holding that end and not to one a column cut short of it.
  const sectionClose =
    division.keepLines === null ? context.sectionsClosed?.get(paragraph.index) : undefined;
  const widthPt =
    context.wraps === false
      ? Number.POSITIVE_INFINITY
      : frame.widthPt - insets.leftPt - insets.rightPt;

  // A piece the column above cut off is broken again from where that column left it,
  // and the text behind it was measured once already when the paragraph first flowed.
  const breaking: LineFlowStart =
    division.resume !== null
      ? { kind: "flow", flow: division.resume }
      : beginLines({
          runs,
          metricsFor: context.metricsFor,
          // A justified line may take a word it has not the room for, so where a line
          // breaks depends on how it is aligned and on how old the document is.
          justified:
            paragraphFrame.alignment === "justify" && squeezesAJustifiedLine(context.settings),
          tabs: {
            stopsPt: tabStopsPt(paragraphFrame),
            originPt: insets.leftPt,
            firstLineOriginPt:
              number === undefined ? insets.leftPt + insets.firstLinePt : insets.leftPt,
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
    ...layOutParagraph(paragraph.index, breaking.flow, {
      topPt,
      anchorTopPt: standing.anchorTopPt,
      markHeightPt: markHeight,
      markWidthPt: widthOfMark(paragraphMark, context.metricsFor),
      frame,
      paragraphFrame,
      spacing: roomEitherSideOf(
        spacingPt(paragraph, paragraphFrame, context, neighbours),
        division,
      ),
      paint: paintOf(paragraphFrame, context, neighbours, {
        leftPt: frame.leftPt + insets.leftPt,
        rightPt: frame.leftPt + frame.widthPt - insets.rightPt,
      }),
      startsPage: !resumed && !context.inCell && paragraphFrame.pageBreakBefore,
      endsPageAtASection: sectionClose?.opensAPage === true,
      closesASection: sectionClose !== undefined,
      closesACellUnderATable: neighbours.closesACellUnderATable,
      number: measured === null ? null : measured.number,
      bands: standing.bands,
      ahead: standing.ahead,
      roomPt: widthPt,
      // A hanging indent leaves its first line wider than the rest, and a numbered
      // paragraph's first line is as wide as the run from where its text starts to
      // the right indent, wherever the number's suffix left that text.
      //
      // **What stood here gave a numbered first line the room between the two
      // indents**, on the reading that a number hanging in front of the text leaves
      // that text starting at the left indent like every line under it. It does
      // where nothing says otherwise, and one of the 966 says otherwise: its number
      // tabs to a stop 26.7pt short of the indent, and Word fits 29 characters on
      // that line where this fitted 11.
      //
      // Measured on 2026-08-10 by the authored `numbered-first-line` document and
      // its legacy twin, whose pdfs are identical. Seven cases in a column 126pt
      // wide from the left indent: a number tabbing to a stop 36pt in front of it
      // took **seven** of the 21pt words where every line under it took six, a
      // suffix of one space took eight from 12.45pt further out again, and a suffix
      // of nothing took nine from the number's own place. Room measured from the
      // left indent would have given all four of them six.
      firstLineRoomPt: firstLineRoomOf(measured, widthPt, frame, insets),
      resumed,
      keepLines: division.keepLines,
    }),
  };
}

// **The room a paragraph asks for above and below itself stands at its own two ends
// and not at a cut a column made in it.** The same reading `partOf` takes of a page
// break running through a paragraph, whose pieces keep the space before at the first
// and the space after at the last.
const roomEitherSideOf = (spacing: Spacing, division: Division): Spacing => ({
  beforePt: division.resume === null ? spacing.beforePt : 0,
  afterPt: division.keepLines === null ? spacing.afterPt : 0,
});

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
    sameBorders(borders, resolveParagraphFrame(other, context.styles, context.inTable).borders);

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
  const automatic = automaticSpacingPt(paragraph, paragraphFrame, context, neighbours);
  const beforePt = automatic.beforeIsDropped ? 0 : twipsToPoints(paragraphFrame.spaceBeforeTwips);
  const afterPt = automatic.afterIsDropped ? 0 : twipsToPoints(paragraphFrame.spaceAfterTwips);
  if (!paragraphFrame.contextualSpacing) return { beforePt, afterPt };

  const own = styleIdOf(paragraph, context.styles);
  const sameStyle = (other: Paragraph | null): boolean =>
    other !== null && styleIdOf(other, context.styles) === own;

  return {
    beforePt: sameStyle(neighbours.above) ? 0 : beforePt,
    afterPt: sameStyle(neighbours.below) ? 0 : afterPt,
  };
}

// Where an automatic space is worth nothing at all, which is not where a stated
// one is. Measured on 2026-08-13 against Word's own drawing, every case three
// times over:
//
// - **Against the top of what holds the paragraph.** The first paragraph of the
//   body draws its line where a plain one does, while the same paragraph stating
//   14pt draws it 13.92 lower and one stating 24pt draws it 24 lower. The first and
//   last paragraph of a table cell answer the same way, and neither lifts nor grows
//   its row.
// - **Between two paragraphs of one list**, at any level of it and whether or not
//   the other one asks for a space of its own. Two lists meeting keep it, and so
//   does a numbered paragraph standing beside an unnumbered one.
//
// **A table beside the paragraph is not an edge**: a paragraph under a table keeps
// its fourteen points, and so does one over it, which is what tells this apart from
// a rule about having no paragraph on that side.
function automaticSpacingPt(
  paragraph: Paragraph,
  paragraphFrame: ParagraphFrame,
  context: Context,
  neighbours: Neighbours,
): { readonly beforeIsDropped: boolean; readonly afterIsDropped: boolean } {
  if (!paragraphFrame.automaticSpaceBefore && !paragraphFrame.automaticSpaceAfter) {
    return { beforeIsDropped: false, afterIsDropped: false };
  }

  const listOf = (other: Paragraph | null): string | null =>
    other === null ? null : (resolveParagraphNumbering(other, context.styles)?.numId ?? null);
  const own = listOf(paragraph);
  const sameList = (other: Paragraph | null): boolean => own !== null && listOf(other) === own;

  return {
    beforeIsDropped:
      paragraphFrame.automaticSpaceBefore &&
      (neighbours.opensWhatHoldsIt || sameList(neighbours.above)),
    afterIsDropped:
      paragraphFrame.automaticSpaceAfter &&
      (neighbours.closesWhatHoldsIt || sameList(neighbours.below)),
  };
}

// How much room the paragraph above keeps under itself, which is the whole of
// what it already put between the two.
function roomBelowPt(above: Paragraph | null, below: Paragraph, context: Context): number {
  if (above === null) return 0;
  const frame = resolveParagraphFrame(above, context.styles, context.inTable);
  return ownSpacingPt(above, frame, context, {
    above: null,
    below,
    closesACellUnderATable: false,
    // Whatever else it opens, something stands under it: the paragraph asking.
    opensWhatHoldsIt: false,
    closesWhatHoldsIt: false,
  }).afterPt;
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

// A box that never wraps is measured at no width at all, and a first line has to
// stay as unbounded as the rest of them.
function firstLineRoomOf(
  measured: NumberMeasurement | null,
  widthPt: number,
  frame: Frame,
  insets: Insets,
): number {
  if (measured === null || measured.kind === "blocked") return widthPt - insets.firstLinePt;
  if (!Number.isFinite(widthPt)) return widthPt;
  return frame.leftPt + frame.widthPt - insets.rightPt - measured.number.textStartPt;
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
  readonly closesACellUnderATable: boolean;
  readonly bands: readonly WrapBand[];
  readonly ahead: readonly WrapBand[];
  // What a line has room for where nothing stands beside it, which a shape that
  // refuses to wrap leaves unbounded.
  readonly roomPt: number;
  readonly firstLineRoomPt: number;
  // Whether the column above cut the paragraph in two and this is the second piece
  // of it, and how many lines this piece keeps before the column below takes over.
  readonly resumed: boolean;
  readonly keepLines: number | null;
};

// A paragraph with text is as tall as the lines its runs measured to: Word does
// not let the paragraph mark raise a line it shares with a run, however much
// bigger the mark is. An empty paragraph is the mark's height alone.
function layOutParagraph(
  index: number,
  flow: LineFlow,
  input: LayOutParagraphInput,
): LaidParagraph {
  const across = acrossOf(input);
  const laid = layOutWholeParagraph(index, flow, input);
  return { box: droppedPast(laid.box, input, across), rest: laid.rest };
}

// The paragraph as it stands, and what a column that kept only the first of its lines
// left for the column after it.
type LaidParagraph = {
  readonly box: ParagraphBox;
  readonly rest: LineFlow | null;
};

const acrossOf = (input: LayOutParagraphInput): Span => {
  const insets = insetsOf(input.paragraphFrame);
  return {
    leftPt: input.frame.leftPt + insets.leftPt,
    rightPt: input.frame.leftPt + input.frame.widthPt - insets.rightPt,
  };
};

type Span = { readonly leftPt: number; readonly rightPt: number };

// Where a paragraph with no text begins, and what it asks of the room there.
//
// Its whole line is its mark: a number standing at the hanging position and
// reaching to wherever its suffix moves the text on to, or, where there is no
// number, a paragraph mark drawing nothing at the left indent. **Measured against
// Word on 2026-08-12**, over a numbered paragraph indented 144pt with the number
// hanging 18pt in front of it and a box put down to the quarter point beside it:
// the paragraph stays where it is with 20.25pt of room before the box and falls to
// the box's foot with 14.25pt, and hanging the number 36pt instead it falls with
// 22pt, which no least run of free space explains and the reach to the text start
// does. The same paragraph without a number keeps 2.25pt of room and stays.
const markSpanOf = (
  frame: Frame,
  insets: Insets,
  number: MeasuredNumber | null,
): { readonly leftPt: number; readonly widthPt: number } =>
  number === null
    ? { leftPt: frame.leftPt + insets.leftPt, widthPt: 0 }
    : { leftPt: number.leftPt, widthPt: number.textStartPt - number.leftPt };

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
  const mark = markSpanOf(input.frame, insetsOf(input.paragraphFrame), input.number ?? null);
  const fit = last === undefined ? fitMark : fitLine;
  const slot = fit({
    topPt,
    heightPt: box.topPt + box.heightPt - topPt,
    leftPt: last === undefined ? mark.leftPt : across.leftPt,
    rightPt: across.rightPt,
    widthPt: last?.line.widthPt ?? mark.widthPt,
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
): LaidParagraph {
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
  const laidWith = (box: ParagraphBox): LaidParagraph => ({ box, rest: brokenByItsText.rest });
  // What a paragraph holds is held by the piece of it carrying its end, which is the
  // reading `partOf` takes of a page break running through one.
  const keepNext = input.keepLines === null && paragraphFrame.keepNext;

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
  // **The paragraph a cell holding a table has to end with is not laid out
  // either.** Measured on 2026-08-10 by the authored `resuming` document: the row
  // under such a cell opened where the table above it left off, and the same
  // paragraph with a word in it took the whole of its line. A paragraph closing a
  // cell after ordinary text keeps its line too, so what empties this one is the
  // table above it and not the cell it ends.
  if (laid.length === 0 && (input.closesASection || input.closesACellUnderATable)) {
    return laidWith({
      index,
      topPt: input.topPt,
      anchorTopPt: input.anchorTopPt,
      heightPt: 0,
      lines: [],
      marker: null,
      markTopPt: input.topPt,
      contentBottomPt: input.topPt,
      resumesUnderPt: 0,
      widowControl: paragraphFrame.widowControl,
      keepNext,
      startsPage: input.startsPage,
      endsPage,
      endsPageAtASection: input.endsPageAtASection,
      contentWidthPt: 0,
      clipTo: null,
      paint: null,
    });
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
    const mark = markSpanOf(frame, insets, number ?? null);
    const slot = slotFor(
      {
        topPt: input.topPt + beforePt + abovePt,
        heightPt: height.fittingHeightPt,
        roomAbovePt: beforePt,
        leftPt: mark.leftPt,
        rightPt: frame.leftPt + frame.widthPt - insets.rightPt,
        widthPt: mark.widthPt,
        bands: input.bands,
      },
      fitMark,
    );

    return laidWith({
      index,
      topPt: input.topPt,
      anchorTopPt: input.anchorTopPt,
      heightPt: slot.topPt + height.heightPt + belowPt + afterPt - input.topPt,
      lines: [],
      marker: markerAt(number, slot.topPt + height.baseFromTopPt),
      markTopPt: slot.topPt + height.seatPt,
      // **The room a multiple opens below a mark hangs past the foot, as it does
      // below a line of text**, which is what `breakStack` has always read
      // `fittingHeightPt` for. Measured on 2026-08-24, one document a case, by
      // whether Word made a second page for a trailing empty paragraph: a body of
      // 720pt holding 48 lines of 14.6484 leaves 16.90, and a mark under a rule of
      // 1.3 asking 19.04 for a line of its own 14.65 stayed on the page, while the
      // same mark under no multiple with 49 lines above it, where its own line does
      // not fit either, opened a second page. Two corpus documents of one template
      // turn on it, each opening a page for a mark that missed the foot by two
      // tenths of a point.
      contentBottomPt: slot.topPt + height.fittingHeightPt,
      resumesUnderPt: 0,
      widowControl: paragraphFrame.widowControl,
      keepNext,
      startsPage: input.startsPage,
      endsPage,
      endsPageAtASection: input.endsPageAtASection,
      contentWidthPt: slot.leftPt - frame.leftPt + input.markWidthPt,
      clipTo: null,
      paint,
    });
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
      fittingHeightPt: each.height.fittingHeightPt,
      baselinePt: each.slot.topPt + each.height.baseFromTopPt,
      startsPage: each.startsPage,
    };
  });

  const last = laid[laid.length - 1];
  const bottomPt = last === undefined ? input.topPt : last.slot.topPt + last.height.heightPt;

  return laidWith({
    index,
    topPt: input.topPt,
    anchorTopPt: input.anchorTopPt,
    heightPt: bottomPt + belowPt + afterPt - input.topPt,
    lines: placed,
    marker: markerAt(number, placed[0]?.baselinePt ?? input.topPt),
    markTopPt: last === undefined ? input.topPt : last.slot.topPt + last.height.seatPt,
    contentBottomPt: bottomPt,
    resumesUnderPt: 0,
    widowControl: paragraphFrame.widowControl,
    keepNext,
    startsPage: input.startsPage,
    endsPage,
    endsPageAtASection: input.endsPageAtASection,
    contentWidthPt: placed.reduce(
      (widest, line) => Math.max(widest, line.leftPt - frame.leftPt + line.line.widthPt),
      0,
    ),
    clipTo: null,
    paint,
  });
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
  // What the paragraph has left where a column kept only the first of its lines,
  // which the next column breaks again at its own width.
  readonly rest: LineFlow | null;
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
    // A piece of a paragraph the column above began has no first line of its own:
    // the indent, the room and the number all stood in the column it was cut from.
    const opens = at === 0 && !input.resumed;
    const roomPt = opens ? input.firstLineRoomPt : input.roomPt;
    const firstLinePt = opens ? insets.firstLinePt : 0;
    // The number takes the first line's own start, so the text after it begins
    // wherever the number's suffix moved on to.
    const startPt =
      opens && number !== null ? number.textStartPt : frame.leftPt + insets.leftPt + firstLinePt;
    const endPt = frame.leftPt + frame.widthPt - insets.rightPt;

    const leastPt = rest.leastPt;
    const startsPage = rest.startsPage;
    let taken = rest.next(roomPt);
    if (taken === null) return { lines: laid, endsPage: startsPage, rest: null };

    // Only the first line has room asked for above it; the rest follow the line
    // before them.
    const roomAbovePt = at === 0 ? input.spacing.beforePt : 0;
    const fit = { roomAbovePt, widthPt: leastPt, leftPt: startPt, rightPt: endPt };

    let height = heightOfLine(taken.line, at, input);
    let slot = slotFor({
      ...fit,
      topPt: top,
      heightPt: height.fittingHeightPt,
      bands: input.bands,
    });

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
      slot = slotFor({ ...fit, topPt: top, heightPt: height.fittingHeightPt, bands: input.bands });
    }

    laid.push({ line: taken.line, slot, height, startsPage });
    // The column asked for only so many lines, and what is left of the paragraph goes
    // to the column after it, which breaks it again at its own width. A paragraph
    // ends nowhere here, so it holds no page back either.
    if (laid.length === input.keepLines) return { lines: laid, endsPage: false, rest: taken.rest };
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
// with. **Nor does it stand clear of a band with the whole of the room its line
// rule opened**: what has to clear the top of a band is the text seated in that
// room, and room a multiple opened below the text hangs into the band as it hangs
// past the foot of a page. Measured on 2026-08-11 by the authored
// `line-into-a-band` document, whose lines under a multiple of two hold 14.65 of
// text in a box of 29.30: a band opening 7.3pt below the fourth line's text keeps
// four lines above it, and one opening 6.5pt above that text keeps three.
//
// **The room a paragraph asks for above itself goes through a wrap with its first
// line.** What has to clear an object is the room and the line together, so a line
// pushed past one stands that room below it rather than against it, and one whose
// room alone reaches an object is drawn beside it as the line itself would be.
// Measured on 2026-08-07 by the authored `space-under-a-wrap` document: a paragraph
// asking 36pt above itself under a box whose foot is at 190 has its line at 226,
// and one whose line stands 44pt clear of the foot of a box beside it is drawn to
// the right of that box all the same.
function slotFor(slot: Slot, fit: (input: FitLineInput) => LineSlot = fitLine): LineSlot {
  const { roomAbovePt } = slot;
  const found = fit({
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

// A line a tab or a space alone holds open has nothing measured on it to give it a
// height, so it is held open by the run whose break ends it, and by the paragraph's
// own mark where no break does. **Measured on 2026-08-11 by the authored
// `empty-line-size` document**: a paragraph holding one break comes out one line of
// the break's run and one of the mark, whichever of the two is the larger, so a
// mark twice the size raises the second line alone and a break's run twice the size
// raises the first.
function heightOfLine(line: TextLine, at: number, input: LayOutParagraphInput): LineHeight {
  const raisedPt = raisedBy(line, at, input);
  const held = line.segments.length === 0 ? (line.heldOpenPt ?? input.markHeightPt) : 0;
  return seatedHeight(
    {
      naturalPt: Math.max(line.heightPt, held) + raisedPt,
      ascentPt: line.ascentPt + raisedPt,
      seatPt: line.seatPt,
      // **The room a number adds is not part of what a multiple is taken of**: Word
      // multiplies the line the paragraph's own faces made and adds the number's
      // room on top of that. Measured on 2026-08-24 over twelve paragraphs of Arial
      // 10pt under a rule of 1.1, bulleted from a Symbol level: multiplying the
      // lifted line puts their span at 210.97 and Word drew 210.48, which is what
      // adding the room after the multiple gives to a hundredth of a point. Three
      // corpus documents of one template turn on it, each making a page more than
      // Word does on the half point it costs a page.
      fontHeightPt: Math.max(line.fontHeightPt, held),
    },
    input.paragraphFrame,
  );
}

// How tall a line stands in the stack, how far down that room its own text sits,
// and where its baseline falls from the line's top.
type LineHeight = {
  readonly heightPt: number;
  readonly seatPt: number;
  readonly fittingHeightPt: number;
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
    return {
      heightPt,
      seatPt: 0,
      fittingHeightPt: heightPt,
      baseFromTopPt: heightPt * EXACT_BASELINE,
    };
  }

  const openedPt = frame.lineRule === "atLeast" ? Math.max(0, heightPt - line.naturalPt) : 0;
  const seatPt = openedPt + line.seatPt;
  return {
    heightPt,
    seatPt,
    fittingHeightPt: seatPt + line.naturalPt,
    baseFromTopPt: seatPt + line.ascentPt,
  };
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
