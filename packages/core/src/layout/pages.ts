import { borderExtentPt } from "../docx/borders.js";
import type { Border } from "../docx/borders.js";
import type { AnchoredObject, ParagraphBox, PlacedCell, UntornRow } from "./stack.js";

// The run of a page the body's text may stand in: where it begins under the
// header, and where the footer or the bottom margin ends it.
export type PageBody = {
  readonly topPt: number;
  readonly bottomPt: number;
};

export type PageStack = {
  readonly index: number;
  readonly boxes: readonly ParagraphBox[];
  readonly cells: readonly PlacedCell[];
  // The paragraph whose text opened the page, which is what says which section
  // made it. Null on a page no paragraph reached.
  readonly openedBy: number | null;
};

export type BreakStackInput = {
  // A stack measured from `topPt` with no bottom, as `measureStack` produces it.
  readonly boxes: readonly ParagraphBox[];
  readonly cells: readonly PlacedCell[];
  readonly untornRows?: readonly UntornRow[];
  readonly anchoredObjects?: readonly AnchoredObject[];
  readonly topPt: number;
  readonly bottomPt: number;
  // What the section a paragraph stands in keeps for the body, where the sections
  // of a document make different pages. `topPt` and `bottomPt` are what a document
  // of one section keeps and what this falls back to.
  //
  // **A page belongs to the section whose text opened it**, so a page is asked for
  // once, of the paragraph that opens it, and holds what it answered to its foot: a
  // continuous section beginning partway down a page is drawn on the page it found,
  // and makes only the pages its own text runs on to.
  readonly bodyOf?: (box: ParagraphBox) => PageBody;
};

// Where a page started in the stack, what it keeps for the body, and the paragraph
// whose text opened it.
type Opening = {
  readonly shiftPt: number;
  readonly body: PageBody;
  readonly openedBy: number | null;
};

// A paragraph asking to stand with the one after it moves onto the page that one
// begins, so the stack is broken over and over: moving one leaves the paragraph
// above it split from what it holds and moves that too, until the break has walked
// back to the head of the run and a whole chain of them travels together.
//
// **Word moves a paragraph at most once**, which is what ends this, and it is
// measured rather than a guard against looping: where a chain is followed by
// something that can never stand with it, the second of the chain stayed where the
// first move put it, mid page, with room below it to move into and nothing to gain
// by it.
export function breakStack(input: BreakStackInput): readonly PageStack[] {
  const moved = new Set<number>();
  for (;;) {
    const broken = breakOnce(input, moved);
    if (broken.split === null) return broken.pages;
    moved.add(broken.split);
  }
}

// Every page repeats the one body box, so a page is the stack shifted back up by
// however much of it came before: what carried on to the next page starts again at
// the body's top. What a paragraph put above its first line, its space before, is
// what the break leaves behind.
//
// `moved` is the paragraphs `keepNext` has already carried forward. Only the first
// of a run of them leaves a page: the rest are carried by it, and a second break
// would put each of them on a page of its own. `split` in the answer is the first
// paragraph this pass found parted from the one it holds, which is the next to
// move.
function breakOnce(
  input: BreakStackInput,
  moved: ReadonlySet<number>,
): { readonly pages: readonly PageStack[]; readonly split: number | null } {
  const everyPage: PageBody = { topPt: input.topPt, bottomPt: input.bottomPt };
  const bodyOf = (box: ParagraphBox | undefined): PageBody =>
    box === undefined ? everyPage : (input.bodyOf?.(box) ?? everyPage);

  const pages: ParagraphBox[][] = [[]];
  const first = input.boxes[0];
  // What the page being filled keeps for the body, which is what the paragraph that
  // opened it asked for.
  let body = bodyOf(first);
  let shiftPt = 0;
  // Where in the stack each page started and what it kept, which is what the cells
  // are cut by once the text has said where the pages fall.
  const opened: Opening[] = [{ shiftPt, body, openedBy: first?.index ?? null }];

  const put = (box: ParagraphBox): void => {
    pages[pages.length - 1]?.push(box);
  };

  const open = (next: PageBody, openedBy: number): void => {
    body = next;
    opened.push({ shiftPt, body, openedBy });
    pages.push([]);
  };

  const overflows = (topPt: number, heightPt: number): boolean =>
    topPt - shiftPt + heightPt > body.bottomPt && topPt - shiftPt > body.topPt;

  // The page is left where the break asked, so what comes after it starts again at
  // the top of the page the paragraph opening it makes. A break with nothing left to
  // move makes no page: a paragraph asking for one of its own while already standing
  // at the top of one stays where it is, which is what keeps the first paragraph of
  // a document off page two.
  const leave = (topPt: number, next: PageBody, openedBy: number): boolean => {
    if (topPt - shiftPt <= body.topPt) return false;
    shiftPt = topPt - next.topPt;
    open(next, openedBy);
    return true;
  };

  // Keyed by the paragraph a row opens with, which is the last moment the whole of
  // it can still be moved.
  const opening = new Map<number, UntornRow>();
  for (const row of input.untornRows ?? []) opening.set(row.opensAt, row);

  // Keyed the same way, and a paragraph may anchor several: a page of a real
  // document holds three objects on one paragraph.
  const anchoring = new Map<number, AnchoredObject[]>();
  for (const object of input.anchoredObjects ?? []) {
    const already = anchoring.get(object.anchoredAt);
    if (already === undefined) anchoring.set(object.anchoredAt, [object]);
    else already.push(object);
  }

  // Whether the paragraph before this one ended on a page break, which draws no
  // line of its own and so has nothing here to be seen at, and whether that break
  // was a section's.
  let broken = false;
  let brokenAtASection = false;
  // The first paragraph this pass found parted from the one it holds.
  let split: number | null = null;

  for (const [place, box] of input.boxes.entries()) {
    // What a page this paragraph opens keeps for the body, which is its own
    // section's and not the page it may be standing at the foot of.
    const opens = bodyOf(box);
    const carriedForward = moved.has(box.index) && !moved.has(input.boxes[place - 1]?.index ?? -1);
    if (broken || box.startsPage || carriedForward) {
      // **A page a section break opens keeps the room the paragraph opening it asks
      // for above itself, and a page any other break opens does not.** Measured on
      // 2026-08-08 by the authored `space-above-a-break` document: of the four kinds
      // of break that open a page, the foot of the page filling, a break in the
      // paragraph's own text, a paragraph asking for a page of its own and a section
      // break, only the last drew its first line the 18pt it asked for below the top
      // of the page. So the paragraph's own top goes to the top of the page there,
      // and its first line's does everywhere else.
      const opensAt =
        brokenAtASection && !box.startsPage && !carriedForward
          ? box.topPt
          : (box.lines[0]?.topPt ?? box.topPt) - box.resumesUnderPt;
      leave(opensAt, opens, box.index);
    }
    broken = box.endsPage;
    brokenAtASection = box.endsPageAtASection;

    // A row that will not come apart is decided here, at the paragraph that opens
    // it: what does not fit below moves whole. Moved, its top stands at the top of
    // a page, and a row still too tall for a whole page is then torn where any
    // other would be, which is what Word did with one it was told not to split.
    const row = opening.get(box.index);
    if (row !== undefined && overflows(row.topPt, row.bottomPt - row.topPt)) {
      leave(row.topPt, opens, box.index);
    }

    // **An object text wraps round moves to the page under it when it will not fit
    // in the room left, and takes the paragraph anchoring it with it.** Measured on
    // 2026-08-07 by the authored `objects-past-the-foot` document. An object
    // wrapping nothing hangs past the foot instead, however far past, and moves
    // neither itself nor anything else, which is why only the wraps a band is made
    // for reach this. The room is measured to the foot of the text and not to the
    // edge of the sheet: an object reaching into the bottom margin and no further
    // moved like any other.
    //
    // The page is left at the paragraph's own top rather than at its first line's,
    // which is the one place this parts from every other break here. An object is
    // anchored to the paragraph, so it is the paragraph's top that has to land at
    // the top of the new page for the object to land there too: a wrap taking the
    // whole width with it holds the paragraph's own first line 300pt below the
    // anchor, and leaving the page at that line would put the object above the top
    // of the page.
    //
    // The room is asked for from where the object may be drawn up to and not from
    // where it was put, since an object hanging below its paragraph is drawn back up
    // to the foot of the text before anything moves. **What it may be drawn up to is
    // the foot of the line anchoring it and not the paragraph's own top**, measured
    // on 2026-08-08 by the authored `objects-and-the-footer` document: a box hung
    // 100pt below a 24pt line with 172pt of room under it moved, though rising to
    // the foot of the text would have left its top 22pt below the paragraph's own.
    // It would have left it 2pt above the foot of its line, and an object is never
    // drawn over the line that anchors it.
    //
    // A paragraph answers for every object at once. The corpus template that found
    // this anchors three to one paragraph and only the first will not fit, and Word
    // takes the paragraph and all three on to the next page.
    const objects = anchoring.get(box.index);
    if (objects?.some((each) => overflowsHighest(each, box, overflows)) === true) {
      leave(box.topPt, opens, box.index);
    }

    // A paragraph with nothing in it is judged by the room its mark stands in, as
    // one with lines is judged by its lines: the room it keeps below itself hangs
    // past the foot of the page rather than moving it on.
    if (box.lines.length === 0) {
      if (overflows(box.topPt, box.contentBottomPt - box.topPt)) {
        shiftPt = box.topPt - opens.topPt - box.resumesUnderPt;
        open(opens, box.index);
      }
      put(partOf(box, 0, 0, shiftPt));
    } else {
      let from = 0;
      let at = 0;
      while (at < box.lines.length) {
        const line = box.lines[at];
        if (line === undefined) {
          at += 1;
          continue;
        }
        // A line a break of the paragraph's own put at the head of a page goes there
        // whatever room was left below, and widow control has no say in where a break
        // the document asked for falls.
        const asked = line.startsPage;
        if (!asked && !overflows(line.topPt, line.heightPt)) {
          at += 1;
          continue;
        }

        // The lines between the cut and the line that overflowed are on the next
        // page now, so they are looked at again from there. A line inside a table
        // lands as far below the top of the page as its row draws furniture above
        // it, which is what a torn row resumes under.
        const cut = asked ? at : cutFor(box, { from, at, shiftPt, topPt: opens.topPt });
        if (cut > from) put(partOf(box, from, cut, shiftPt));
        shiftPt = (box.lines[cut]?.topPt ?? line.topPt) - opens.topPt - box.resumesUnderPt;
        open(opens, box.index);
        from = cut;
        at = cut + 1;
      }

      put(partOf(box, from, box.lines.length, shiftPt));
    }

    // Where the page break falls between this paragraph and the next: either the
    // next one overflows the page, or it is one already being carried forward and
    // this one is not going with it, which is how a chain is walked back a
    // paragraph at a time.
    const next = input.boxes[place + 1];
    const partsHere =
      next !== undefined &&
      ((moved.has(next.index) && !moved.has(box.index)) || overflowsAtItsStart(next, overflows));

    if (split === null && !moved.has(box.index) && partsHere && holdsAcross(box, next)) {
      split = box.index;
    }
  }

  return {
    pages: pages.map((boxes, index) => ({
      index,
      boxes,
      cells: cellsOn(input, opened, index),
      openedBy: opened[index]?.openedBy ?? null,
    })),
    split,
  };
}

// Whether an object will not fit even drawn as high as it is allowed to go.
//
// An object is drawn up towards the foot of the text and never above the paragraph
// anchoring it, so the highest its top can reach is its own top or that paragraph's,
// whichever is lower: one already standing above its anchor cannot rise at all, and
// a real document anchors one 348pt above the paragraph it hangs off.
const overflowsHighest = (
  object: AnchoredObject,
  box: ParagraphBox,
  overflows: (topPt: number, heightPt: number) => boolean,
): boolean =>
  overflows(Math.min(object.topPt, anchorLineFootPt(box)), object.bottomPt - object.topPt);

// The foot of the line an object is anchored to, which is as high as one is ever
// drawn. A paragraph with nothing in it anchors from the room its mark stands in.
export const anchorLineFootPt = (box: ParagraphBox): number => {
  const first = box.lines[0];
  return first === undefined ? box.contentBottomPt : first.topPt + first.heightPt;
};

// Whether the paragraph asks to stand with the one after it over a break neither
// of them asked for.
//
// A break either asked for is not one this undoes, which is measured from both
// sides: a paragraph holding one that starts a page of its own stayed where it was,
// and so did one holding the text after a break of its own.
//
// It is the paragraph's **end** that is held rather than the paragraph. One the
// break already runs through has its last line on the page its next one begins and
// so is not parted from it at all: Word left such a paragraph broken where widow
// control broke it and moved nothing.
function holdsAcross(box: ParagraphBox, next: ParagraphBox | undefined): boolean {
  if (!box.keepNext || box.endsPage) return false;
  return next !== undefined && !next.startsPage && next.lines[0]?.startsPage !== true;
}

// A paragraph begins where its first line does, and one with nothing in it where
// the room its mark stands in does.
function overflowsAtItsStart(
  box: ParagraphBox,
  overflows: (topPt: number, heightPt: number) => boolean,
): boolean {
  const first = box.lines[0];
  return first === undefined
    ? overflows(box.topPt, box.contentBottomPt - box.topPt)
    : overflows(first.topPt, first.heightPt);
}

// A cell is cut by the pages the text broke into rather than breaking them: the
// piece of one standing on a page is what it covers between where that page
// started in the stack and where the next page did.
function cellsOn(
  input: BreakStackInput,
  opened: readonly Opening[],
  index: number,
): readonly PlacedCell[] {
  const page = opened[index];
  if (page === undefined) return [];
  const { shiftPt, body } = page;
  const fromPt = body.topPt + shiftPt;
  const next = opened[index + 1];
  const nextPt = next === undefined ? Number.POSITIVE_INFINITY : next.body.topPt + next.shiftPt;
  const footPt = fromPt + (body.bottomPt - body.topPt);

  return input.cells.flatMap((cell) => {
    // **A cell the break runs through runs on to the foot of the page**, and one
    // that begins after the break stands on the next page rather than leaving a
    // sliver of itself at the foot of this one. The two are the same thing only
    // where a page ends exactly where the next one takes up, which a torn row is
    // the case against: the room the row keeps above its text on the page below is
    // room the page above never had.
    const cutPt = cell.topPt < nextPt && cell.topPt + cell.heightPt > nextPt ? footPt : nextPt;
    const topPt = Math.max(cell.topPt, fromPt + halfPt(cell.borders.top));
    const bottomPt = Math.min(cell.topPt + cell.heightPt, cutPt - halfPt(cell.borders.bottom));
    if (bottomPt - topPt <= 0) return [];
    return [{ ...cell, topPt: topPt - shiftPt, heightPt: bottomPt - topPt }];
  });
}

// **A row a break runs through keeps the line closing it inside the page**, at
// both ends: the edge a cut leaves is half a border in from where the page ends
// rather than on it. Measured on 2026-08-10 by the authored `resuming` document,
// whose 3pt-bordered row was torn with 96pt of it on each page: Word drew the line
// closing the first piece over the last 2.88pt of the body and the one opening the
// second over the first 2.88pt of it.
const halfPt = (border: Border | null): number => borderExtentPt(border) / 2;

type Cut = {
  // The first of the paragraph's lines still to be placed, and the one whose own
  // height took it over the bottom of the page.
  readonly from: number;
  readonly at: number;
  readonly shiftPt: number;
  readonly topPt: number;
};

// Word will not leave a paragraph's first line alone at the foot of a page, nor
// its last line alone at the top of the next one: `w:widowControl` moves a break
// that would do either back a line, and takes the whole paragraph over when a
// single line is all that would be left above it. A break with nowhere to move to
// is left where it fell, since a page that carries nothing forward never ends.
function cutFor(box: ParagraphBox, cut: Cut): number {
  if (!box.widowControl) return cut.at;

  const alone = cut.at === box.lines.length - 1 ? cut.at - 1 : cut.at;
  const moved = cut.from === 0 && alone < 2 ? 0 : alone;
  const shiftPt = (box.lines[moved]?.topPt ?? 0) - cut.topPt;
  return moved < cut.from || shiftPt <= cut.shiftPt ? cut.at : moved;
}

// A paragraph the break ran through keeps its index on both pages: it is the one
// paragraph, drawn in two places. Only the part holding its first line carries the
// number the list put in front of it.
function partOf(box: ParagraphBox, from: number, to: number, shiftPt: number): ParagraphBox {
  const lines = box.lines.slice(from, to);
  const first = lines[0];
  const last = lines[lines.length - 1];
  const topPt = (from === 0 || first === undefined ? box.topPt : first.topPt) - shiftPt;
  const bottomPt =
    to === box.lines.length || last === undefined
      ? box.topPt + box.heightPt - shiftPt
      : last.topPt + last.heightPt - shiftPt;

  return {
    index: box.index,
    topPt,
    anchorTopPt: box.anchorTopPt - shiftPt,
    heightPt: bottomPt - topPt,
    lines: lines.map((line) => ({
      ...line,
      topPt: line.topPt - shiftPt,
      baselinePt: line.baselinePt - shiftPt,
    })),
    marker:
      box.marker === null || from > 0
        ? null
        : { ...box.marker, baselinePt: box.marker.baselinePt - shiftPt },
    // The mark stands at the end of the paragraph, so it goes with the part that
    // holds the last of it.
    markTopPt: to === box.lines.length ? box.markTopPt - shiftPt : bottomPt,
    contentBottomPt:
      to === box.lines.length || last === undefined
        ? box.contentBottomPt - shiftPt
        : last.topPt + last.heightPt - shiftPt,
    resumesUnderPt: box.resumesUnderPt,
    widowControl: box.widowControl,
    // What the paragraph asked of the pages either side of it belongs to the part
    // of it that stands there, and what it holds is held by the part carrying its
    // end.
    keepNext: to === box.lines.length && box.keepNext,
    startsPage: from === 0 && box.startsPage,
    endsPage: to === box.lines.length && box.endsPage,
    endsPageAtASection: to === box.lines.length && box.endsPageAtASection,
    contentWidthPt: box.contentWidthPt,
    clipTo: box.clipTo === null ? null : { ...box.clipTo, topPt: box.clipTo.topPt - shiftPt },
    paint: box.paint,
  };
}
