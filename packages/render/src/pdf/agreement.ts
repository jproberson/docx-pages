import {
  paragraphBoxesOn,
  type LaidOutDocument,
  type ParagraphBox,
  type PlacedLine,
} from "@docx-pages/core";

import type { DrawnText, TextPlacement } from "./text.js";

// How far a laid-out document agrees with Word's own drawing of it.
//
// Word's pdf says both which lines it broke each paragraph into and where every run
// of every line sat, so this answers two questions at once: did we break the text
// where Word did, and did we then put it where Word put it. Breaking is asked of
// the whole document and placement of the page, since a line Word drew on another
// page was still broken the way Word broke it.
//
// This began as the body of `text.reference.test.ts` and answers for the eight
// reference documents still. It is here because the corpus asks the same question
// of documents nobody has measured: with Word's pdf beside each one, a corpus stops
// being a pile that can only say a document changed and becomes one that can say a
// document is wrong.

// What one page's lines say when they are read together rather than counted.
//
// **A count of misplaced lines cannot tell a page that is moved from a page that is
// broken**, and for a preview those are opposite verdicts: a page whose every line
// is 13pt low is the page Word drew, put down a little wrong, and a page whose lines
// each disagree by a different amount is text over text and unusable. The same two
// pages score the same on `placed`, which asks only whether each line landed inside
// a point of Word's own, and the same on the raster, which asks only how many cells
// differ. The displacements were being computed and thrown away; keeping them
// answers it.
//
// - `agrees`: nothing on the page is visibly out of place.
// - `shifted`: one offset explains nearly all of them. Constant down with nothing
//   across is a page break in the wrong place.
// - `drifting`: no one offset explains it, but every line still stands under the line
//   above it, so the page was squeezed or stretched and can still be read.
// - `deformed`: a line crossed another. Nothing that moves a page as a whole can do
//   that, and it is what text drawn over text looks like from here.
// - `missing`: Word drew visibly more of the page than we did.
//
// **Every threshold here was chosen against the raster and not guessed at**, since a
// reading that calls a page wrong where Word drew it cell for cell is worse than no
// reading, and the first cut of this did it three ways. The thresholds are what joining
// the two readings of the corpus page by page settled, on 2026-08-12.
export type PageShape = "agrees" | "shifted" | "drifting" | "deformed" | "missing";

// How far our drawing of a page stands from Word's, where one offset explains it.
// Positive is Word's drawing further right and further down than ours.
export type Offset = { readonly leftPt: number; readonly downPt: number };

export type PageAgreement = {
  readonly index: number;
  readonly shape: PageShape;
  // Lines we drew on this page with something on them, and how many of those Word
  // drew the same text for on this same page.
  readonly lines: number;
  readonly matched: number;
  // Lines of ours whose text Word drew on some other page, which is what a page
  // break in the wrong place leaves behind, and lines of ours Word drew nowhere at
  // all.
  readonly onAnotherPage: number;
  readonly oursAlone: number;
  // Items Word drew on this page that spell something, and how many of those no
  // line of ours claimed.
  //
  // **These two count matching failures as readily as missing content**, and the
  // raster proved it: pages where a fifth to three quarters of one side went
  // unclaimed came out cell for cell equal to Word's own drawing, because a line
  // Word broke into pieces this cannot join is a line nobody drew wrong. They are
  // kept because they say where to look, and `missing` is decided by the ink below.
  readonly theirs: number;
  readonly theirsAlone: number;
  // How much advance width each side put on the page, over the text that spells
  // something. **This is what says content is missing**, and it says it whether or
  // not a line could be matched: a page whose text we drew somewhere Word did not
  // still carries the same ink, and a page whose chart or table we never drew at all
  // does not.
  readonly inkOursPt: number;
  readonly inkTheirsPt: number;
  // The offset explaining a shifted page, and how fast the displacement grows with
  // the line's own baseline on a drifting one. Null where the page is neither.
  readonly offsetPt: Offset | null;
  readonly driftPerPt: number | null;
  // What share of the matched lines the shape accounts for, and the worst
  // displacement on the page whatever its shape.
  readonly explained: number;
  readonly worstPt: number;
};

export type Agreement = {
  // Lines we laid out with something on them, which is what the rest are counted
  // against. Every line the comparison looked for, header and footer and text box
  // alike, and not the body's alone: counting a narrower set than was searched is
  // how a share comes out above all of them.
  readonly lines: number;
  // Lines whose text Word drew somewhere, and how many of those we put where Word
  // put them.
  readonly matched: number;
  readonly placed: number;
  // The same of every run of every line: a line whose runs differ in face or size
  // is several items in Word's pdf, and each has a start of its own.
  readonly runsMatched: number;
  readonly runsPlaced: number;
  // And of the number a list draws in front of a line.
  readonly numbersMatched: number;
  readonly numbersPlaced: number;
  // What Word drew the text at against the size the document asks for, over the
  // lines whose text matched, taken as the middle of them. One is a drawing to
  // compare against; anything else is a drawing of the whole page shrunk, which
  // nothing here can be held to. Null where nothing matched.
  readonly drawnScale: number | null;
  // Each of our pages read as a whole, and how many pages Word's own drawing holds.
  // A page Word drew that we never made appears in neither list, so the two counts
  // are what says so.
  readonly pages: readonly PageAgreement[];
  readonly pagesDrawn: number;
};

const textOf = (placed: PlacedLine): string =>
  placed.line.segments.map((segment) => (segment.kind === "text" ? segment.text : "")).join("");

// The size the line's own text is set in, which is the tallest of its runs: what
// Word drew there is compared against it to see whether the two are the same
// drawing at all.
function sizeOf(placed: PlacedLine): number | null {
  const sizes = placed.line.segments.flatMap((segment) =>
    segment.kind === "text" && inkOf(segment.text) !== "" ? [segment.mark.fontSizePt] : [],
  );
  return sizes.length === 0 ? null : Math.max(...sizes);
}

const middleOf = (values: readonly number[]): number | null => {
  if (values.length === 0) return null;
  const sorted = [...values].sort((one, other) => one - other);
  return sorted[Math.floor(sorted.length / 2)] ?? null;
};

// Where the byte a symbol face shadows stands for a character Unicode already
// names, Word says so on the way into the pdf and the character is what comes
// back: Symbol's 0xb7 is the bullet, and arrives as one.
const NAMED_IN_UNICODE = new Map([[0xf0b7, 0x2022]]);

// A symbol face's characters reach the page through the private use page Word
// writes them in, and come back out of the pdf in the low byte they shadow.
const outOfSymbolPage = (text: string): string =>
  Array.from(text, (character) => {
    const codePoint = character.codePointAt(0) ?? 0;
    const named = NAMED_IN_UNICODE.get(codePoint);
    if (named !== undefined) return String.fromCodePoint(named);
    return codePoint >= 0xf020 && codePoint <= 0xf0ff
      ? String.fromCodePoint(codePoint - 0xf000)
      : character;
  }).join("");

// A glyph whose face names no character for it comes back out of the pdf as a nul.
const UNNAMED_IN_THE_PDF = String.fromCodePoint(0);

const normalise = (text: string): string =>
  outOfSymbolPage(text.replaceAll(UNNAMED_IN_THE_PDF, "").replace(/\s+/g, " ").trim());

// A stretch of text drawn in one face at one place, which is what Word writes an
// item for and what a line is made of on our side.
type Run = {
  readonly text: string;
  readonly leftPt: number;
  readonly baselinePt: number;
};

// Runs line up by the characters that carry ink, since the spaces around them are
// drawn by whichever run happens to hold them. A run opening on a space has no
// inked character of its own to line up at, so it is left out.
const inkOf = (text: string): string => text.replace(/\s+/g, "");

// Word writes an item per run, so a line whose runs differ in face or size arrives
// as several items in a row. The line is the shortest stretch of items from some
// starting point whose ink spells it out.
//
// Ink, because the whitespace Word draws is not the whitespace the document states.
// Where the faces either side differ, the space between two runs is a run of its
// own and Word draws it as a text object showing nothing but a space. A tab is
// drawn as a space as well, one stretched to the width of the gap it opened, though
// the line it lands on holds no character for it at all. Spelling a line out of what
// carries ink is what the runs of it are already lined up by, and it asks the one
// question Word's own breaking answers: which characters landed here.
function itemsFor(
  text: string,
  drawn: readonly TextPlacement[],
  taken: ReadonlySet<TextPlacement>,
): readonly TextPlacement[] | null {
  const wanted = inkOf(text);
  for (const [start, first] of drawn.entries()) {
    // A stretch starts where a line does, at ink: one starting at a space could
    // begin by claiming the space the line above ended on.
    if (inkOf(first.text) === "") continue;
    let joined = "";
    let reached = first.leftPt;
    for (const [end, item] of drawn.slice(start).entries()) {
      if (taken.has(item)) break;
      // Word draws a line from its own start towards its end, so an item standing
      // left of the one before it is on some other line however the two read
      // together. Without this a line ending in the same letters another begins
      // with is spelled out of the two of them, and lands on neither.
      if (item.leftPt < reached) break;
      reached = item.leftPt;
      joined += item.text;
      const spelled = inkOf(normalise(joined));
      if (spelled === wanted) return drawn.slice(start, start + end + 1);
      if (!wanted.startsWith(spelled)) break;
    }
  }
  return null;
}

function startsOf(runs: readonly Run[]): ReadonlyMap<number, number> {
  const found = new Map<number, number>();
  let at = 0;
  for (const run of runs) {
    if (inkOf(run.text) !== "" && !/^\s/.test(run.text) && !found.has(at)) {
      found.set(at, run.leftPt);
    }
    at += inkOf(run.text).length;
  }
  return found;
}

const runsOf = (placed: PlacedLine): readonly Run[] =>
  placed.line.segments.flatMap((segment) =>
    segment.kind === "text"
      ? [
          {
            text: outOfSymbolPage(segment.text),
            leftPt: placed.leftPt + segment.offsetPt,
            baselinePt: placed.baselinePt - segment.mark.raisePt,
          },
        ]
      : [],
  );

// Where the first run carrying ink was drawn, which is the line's own baseline
// unless that run is raised off it: Word writes the raised run as an item of its
// own and the line is compared against the first item there is.
const firstBaselineOf = (placed: PlacedLine): number =>
  runsOf(placed).find((run) => inkOf(run.text) !== "" && !/^\s/.test(run.text))?.baselinePt ??
  placed.baselinePt;

// How near a number's line has to sit to Word's own before the number drawn there
// can be taken for the same one.
const SAME_LINE_PT = 3;

// The number is claimed as a line is, and until 2026-08-12 it was not: it is drawn
// by the list rather than by the paragraph, so no line of ours ever spells it and
// every bullet in the document was left over as an item Word drew and we did not.
// Over the eight reference documents, whose pages the raster says are Word's cell for
// cell, that alone was 213 items and up to 38% of a page.
function compareNumbers(
  tolerancePt: number,
  boxes: readonly ParagraphBox[],
  drawn: readonly TextPlacement[],
  taken: Set<TextPlacement>,
): { readonly numbersMatched: number; readonly numbersPlaced: number } {
  let numbersMatched = 0;
  let numbersPlaced = 0;

  for (const box of boxes) {
    const marker = box.marker;
    const line = box.lines[0];
    if (marker === null || line === undefined) continue;

    const text = normalise(marker.text);
    if (text === "") continue;

    const near = drawn.filter(
      (item) =>
        !taken.has(item) &&
        normalise(item.text) === text &&
        Math.abs(item.baselinePt - line.baselinePt) <= SAME_LINE_PT,
    );
    const nearest = [...near].sort(
      (one, other) => Math.abs(one.leftPt - marker.leftPt) - Math.abs(other.leftPt - marker.leftPt),
    )[0];
    if (nearest === undefined) continue;
    numbersMatched += 1;
    taken.add(nearest);

    if (Math.abs(nearest.leftPt - marker.leftPt) <= tolerancePt) numbersPlaced += 1;
  }

  return { numbersMatched, numbersPlaced };
}

// How far one of our lines stands from the item Word drew it as, and where down the
// page the line itself sits, which is what a drift is measured against.
type Displacement = {
  readonly atPt: number;
  readonly leftOffPt: number;
  readonly downOffPt: number;
};

// What share of a page's lines one offset, or one drift, has to account for before
// the page is called moved rather than deformed. A page whose lines nearly all agree
// about how far out they are is one fault; a page where a twentieth of them dissent
// is still that fault plus a line of its own.
const EXPLAINS_THE_PAGE = 0.95;

// Under this many matched lines nothing can be told from clustering: two lines agree
// about an offset whatever put them there, and three are needed before a slope is
// anything but the line through two points.
const ENOUGH_TO_CLUSTER = 3;

// How far out of place the worst line on a page has to be before the page is called
// anything but agreed.
//
// **A quarter of a line, because under that the reading cried wolf and the raster
// caught it.** Of the 97 pages the first cut of this called deformed over the clean
// corpus, 17 had nothing on them further out than 3pt: ten of those seventeen are
// drawn by us cell for cell as Word drew them, and the median of the group is a tenth
// of a percent of the page. Past 3pt the picture inverts, immediately and completely:
// of the 82 pages left, two are cell for cell equal and the median differs by a fifth
// of the page. So a page whose lines are all within a quarter of a line of Word's is
// the page Word drew, and whether the residue clusters is a question about the lines
// and not about the page. `placed` is where a line 1.4pt low is still counted wrong.
const VISIBLY_OUT_PT = 3;

// How much more of the page Word has to have drawn than we did before the page is
// called missing.
//
// **Measured in ink rather than in lines that failed to match**, and that was the
// second wolf-cry the raster caught. The first cut called a page missing when a fifth
// of one side went unclaimed, and over the clean corpus 9 of the 44 pages it named
// were drawn cell for cell as Word drew them, with no share of unclaimed lines telling
// the two apart: three quarters of a page can go unmatched with nothing wrong on it,
// because a line Word broke into pieces this cannot spell is a line both sides drew.
// The advance width each side put down does not care whether a line could be matched.
//
// **A twentieth, because the ink itself is measured to a half of a percent and the
// noise in it is one-sided.** Over the twenty pages of the eight documents known to be
// right, fifteen come out inside 0.5% and none of the rest has Word drawing more than
// 2.3%. Two of those three are pages holding a drawing whose labels Word writes as text
// items while we draw them inside the picture, so this cannot be tightened past them
// without calling a page missing what it plainly draws; `h` page 1's bar chart is the
// same shape of thing and is reported as `unknown-drawing` and seen by the raster.
// **This is the instrument for a preview missing content a reader would notice**, and
// two percent of the ink is not that.
//
// Only Word drawing more is asked. The other direction is the noisy one, since our own
// line carries the trailing spaces Word draws as items of their own and this reading
// throws away: `h` page 2 comes out 2.7% heavier on our side with nothing wrong. Text
// we drew that Word drew nowhere is a real fault and nobody has measured it here, so it
// is not claimed.
const DREW_MORE = 0.05;

const meanOf = (values: readonly number[]): number =>
  values.length === 0 ? 0 : values.reduce((sum, each) => sum + each, 0) / values.length;

const insideOf = (
  spots: readonly Displacement[],
  offset: Offset,
  tolerancePt: number,
): readonly Displacement[] =>
  spots.filter(
    (spot) =>
      Math.abs(spot.leftOffPt - offset.leftPt) <= tolerancePt &&
      Math.abs(spot.downOffPt - offset.downPt) <= tolerancePt,
  );

// The one offset that accounts for most of a page. The middle of the displacements
// is where it starts, since a middle is unmoved by however far the dissenting few
// are out; the mean of what that gathers is then tried against it, which catches a
// page whose middle line happens to sit at the edge of the cluster rather than in it.
function offsetExplaining(
  spots: readonly Displacement[],
  tolerancePt: number,
): { readonly offset: Offset; readonly explained: number } {
  const middle: Offset = {
    leftPt: middleOf(spots.map((spot) => spot.leftOffPt)) ?? 0,
    downPt: middleOf(spots.map((spot) => spot.downOffPt)) ?? 0,
  };
  const gathered = insideOf(spots, middle, tolerancePt);
  const meaned: Offset = {
    leftPt: meanOf(gathered.map((spot) => spot.leftOffPt)),
    downPt: meanOf(gathered.map((spot) => spot.downOffPt)),
  };

  const better = insideOf(spots, meaned, tolerancePt).length > gathered.length ? meaned : middle;
  return {
    offset: better,
    explained: spots.length === 0 ? 0 : insideOf(spots, better, tolerancePt).length / spots.length,
  };
}

// How fast the vertical displacement grows with the line's own baseline, taken as a
// straight line through the page. Reported rather than tested: what it is worth is
// saying how fast a page is being squeezed, since a height wrong by a fraction is paid
// once a line and reaches the foot of the page as a slope.
function driftPerPointOf(spots: readonly Displacement[]): number {
  const meanAt = meanOf(spots.map((spot) => spot.atPt));
  const meanDown = meanOf(spots.map((spot) => spot.downOffPt));
  const spread = spots.reduce((sum, spot) => sum + (spot.atPt - meanAt) ** 2, 0);
  if (spread === 0) return 0;
  return (
    spots.reduce((sum, spot) => sum + (spot.atPt - meanAt) * (spot.downOffPt - meanDown), 0) /
    spread
  );
}

// Whether every line still stands under the line above it, which is the whole question
// a preview has to answer.
//
// **Not a straight line through the page, and the corpus is what said so.** The first
// cut fitted a slope to the displacements and called anything that did not fit
// deformed. Page 7 of `cf4e0f837c83` is a table of 216 cells whose rows are each 0.95pt
// shorter than Word's, accumulating to 24pt by the foot of the page: a squeeze, drawn in
// the right order, perfectly readable. It was called deformed, because the squeeze
// starts at the table rather than at the top of the sheet and no straight line fits both
// the lines above it and the rows below.
//
// What tells a squeeze from a wreck is not the shape of the drift but whether the
// vertical order survived it. Take our lines in the order we drew them down the page: if
// Word's own baselines are in that same order, then whatever moved them moved them
// monotonically, no line crossed another, and every reader can still read the page in
// the order it was written. A line that jumped past another is the case a preview cannot
// be shown for, and it is what text drawn over text looks like from here.
function crossingsIn(spots: readonly Displacement[]): number {
  const down = [...spots].sort((one, other) => one.atPt - other.atPt);
  let reached = -Infinity;
  let crossings = 0;
  for (const spot of down) {
    const theirs = spot.atPt + spot.downOffPt;
    if (theirs < reached - VISIBLY_OUT_PT) crossings += 1;
    reached = Math.max(reached, theirs);
  }
  return crossings;
}

function shapeOf(
  spots: readonly Displacement[],
  tolerancePt: number,
): {
  readonly shape: Exclude<PageShape, "missing">;
  readonly offsetPt: Offset | null;
  readonly driftPerPt: number | null;
  readonly explained: number;
  readonly worstPt: number;
} {
  const worstPt = Math.max(
    0,
    ...spots.map((spot) => Math.max(Math.abs(spot.leftOffPt), Math.abs(spot.downOffPt))),
  );
  if (worstPt <= VISIBLY_OUT_PT)
    return { shape: "agrees", offsetPt: null, driftPerPt: null, explained: 1, worstPt };

  const across = offsetExplaining(spots, tolerancePt);
  if (spots.length < ENOUGH_TO_CLUSTER || across.explained >= EXPLAINS_THE_PAGE)
    return {
      shape: "shifted",
      offsetPt: across.offset,
      driftPerPt: null,
      explained: across.explained,
      worstPt,
    };

  // A line out of order is what a preview cannot be shown for, and how far the rest
  // drifted is what a preview can. So the order is asked first and the drift is only
  // ever a description of a page that kept it.
  //
  // **Across the page there is nothing to drift into.** A squeeze is a rule about
  // heights and moves nothing sideways, so a page whose lines are scattered across it
  // is not a squeezed page however well its order survived: those are columns landing
  // in the wrong place, which reads as nonsense without a line ever touching another.
  const acrossKept =
    spots.filter((spot) => Math.abs(spot.leftOffPt - across.offset.leftPt) <= tolerancePt).length /
    spots.length;
  const kept = Math.min(1 - crossingsIn(spots) / spots.length, acrossKept);
  if (kept >= EXPLAINS_THE_PAGE)
    return {
      shape: "drifting",
      offsetPt: null,
      driftPerPt: driftPerPointOf(spots),
      explained: kept,
      worstPt,
    };

  return {
    shape: "deformed",
    offsetPt: null,
    driftPerPt: null,
    explained: kept,
    worstPt,
  };
}

// Word draws an item for a glyph whose face names no character for it, and the pdf
// hands that back as a nul: it spells nothing, so no line of ours can ever claim it
// and counting it as content we missed is crying wolf.
const spellsSomething = (item: TextPlacement): boolean => inkOf(normalise(item.text)) !== "";

export function agreementWith(
  layout: LaidOutDocument,
  drawing: DrawnText,
  tolerancePt: number,
): Agreement {
  // How many pages the file holds, which is the drawing's own answer and not the highest
  // page an item stands on: a page drawing nothing but a picture holds no text item.
  const drawn = drawing.placements;
  const taken = new Set<TextPlacement>();
  const elsewhere: { readonly pageIndex: number; readonly text: string }[] = [];
  const scales: number[] = [];
  const readings: {
    readonly index: number;
    readonly onPage: readonly TextPlacement[];
    readonly lines: number;
    readonly matched: number;
    readonly ink: number;
    readonly spots: readonly Displacement[];
  }[] = [];
  let lines = 0;
  let matched = 0;
  let placed = 0;
  let runsMatched = 0;
  let runsPlaced = 0;
  let numbersMatched = 0;
  let numbersPlaced = 0;

  for (const page of layout.pages) {
    const onPage = drawn.filter((item) => item.pageIndex === page.index);
    const boxes = paragraphBoxesOn(page);
    const spots: Displacement[] = [];
    let linesHere = 0;
    let matchedHere = 0;
    let inkHere = 0;

    // What this page puts on the sheet, against what Word's own items put there. The
    // number a list draws is ink as much as the line is, and Word writes an item for
    // it, so a page of bullets is not a page missing every bullet.
    for (const box of boxes) inkHere += box.marker?.widthPt ?? 0;

    for (const line of boxes.flatMap((box) => box.lines)) {
      const text = normalise(textOf(line));
      if (text === "") continue;
      lines += 1;
      linesHere += 1;
      inkHere += line.line.widthPt;

      const items = itemsFor(text, onPage, taken);
      const found = items?.[0];
      if (items === null || found === undefined) {
        elsewhere.push({ pageIndex: page.index, text });
        continue;
      }
      for (const item of items) taken.add(item);
      matched += 1;
      matchedHere += 1;

      const size = sizeOf(line);
      if (size !== null && size > 0)
        scales.push(Math.max(...items.map((i) => i.fontSizePt)) / size);

      // Only where the run starts is asked of it. How far off the line's baseline a
      // raised one sits is asked of the first of them alone, which is the item the
      // line itself is pinned against.
      const ours = startsOf(runsOf(line));
      const theirs = startsOf(items);

      // Where the line lies is asked of its first inked character rather than of
      // its own start, since a tab opening the line leaves a gap ahead of the first
      // item Word drew and the two starts are then not the same place.
      const baselinePt = firstBaselineOf(line);
      const spot = {
        atPt: baselinePt,
        leftOffPt: (theirs.get(0) ?? found.leftPt) - (ours.get(0) ?? line.leftPt),
        downOffPt: found.baselinePt - baselinePt,
      };
      spots.push(spot);

      const off = Math.max(Math.abs(spot.leftOffPt), Math.abs(spot.downOffPt));
      if (off <= tolerancePt) placed += 1;

      for (const [at, leftPt] of theirs) {
        const mine = ours.get(at);
        if (mine === undefined) continue;
        runsMatched += 1;
        if (Math.abs(mine - leftPt) <= tolerancePt) runsPlaced += 1;
      }
    }

    const numbers = compareNumbers(tolerancePt, boxes, onPage, taken);
    numbersMatched += numbers.numbersMatched;
    numbersPlaced += numbers.numbersPlaced;

    readings.push({
      index: page.index,
      onPage,
      lines: linesHere,
      matched: matchedHere,
      ink: inkHere,
      spots,
    });
  }

  // A line Word drew on another page was still broken the way Word broke it. Which
  // page of ours it came off is kept, since a page whose lines Word drew on the page
  // before is a page break in the wrong place and not text nobody drew.
  const foundElsewhere = new Map<number, number>();
  for (const { pageIndex, text } of elsewhere) {
    const items = itemsFor(text, drawn, taken);
    if (items === null) continue;
    for (const item of items) taken.add(item);
    matched += 1;
    foundElsewhere.set(pageIndex, (foundElsewhere.get(pageIndex) ?? 0) + 1);
  }

  // Read after the whole document, not during it: what Word drew and we did not is
  // whatever no line of ours claimed by the end, and a line of ours reaches for an
  // item on another page in the pass above.
  const pages = readings.map((reading): PageAgreement => {
    const onAnotherPage = foundElsewhere.get(reading.index) ?? 0;
    const oursAlone = reading.lines - reading.matched - onAnotherPage;
    const theirs = reading.onPage.filter(spellsSomething);
    const theirsAlone = theirs.filter((item) => !taken.has(item)).length;

    const read = shapeOf(reading.spots, tolerancePt);

    // How much of the page each side drew. A line whose text Word drew on the page
    // before is still ink this page carries, so nothing here is about matching: it is
    // the advance width put down on the sheet either way.
    const inkOursPt = reading.ink;
    const inkTheirsPt = theirs.reduce((sum, item) => sum + item.widthPt, 0);
    // A page Word put no text on at all is the one place the other direction is not a
    // question of noise: there is nothing there to be noisy about, and whatever we drew
    // on it Word drew nowhere.
    const drewMore =
      inkTheirsPt === 0 ? (inkOursPt > 0 ? 1 : 0) : (inkTheirsPt - inkOursPt) / inkTheirsPt;

    // Deformed before missing, and both before anything else. A page that is both is
    // the worse of the two, and neither is a page a preview can show.
    const shape: PageShape =
      read.shape === "deformed" ? "deformed" : drewMore > DREW_MORE ? "missing" : read.shape;

    return {
      index: reading.index,
      shape,
      lines: reading.lines,
      matched: reading.matched,
      onAnotherPage,
      oursAlone,
      theirs: theirs.length,
      theirsAlone,
      inkOursPt,
      inkTheirsPt,
      offsetPt: read.offsetPt,
      driftPerPt: read.driftPerPt,
      explained: read.explained,
      worstPt: read.worstPt,
    };
  });

  return {
    lines,
    matched,
    placed,
    runsMatched,
    runsPlaced,
    numbersMatched,
    numbersPlaced,
    drawnScale: middleOf(scales),
    pages,
    pagesDrawn: drawing.pages,
  };
}
