import type { LaidOutDocument, LaidOutPage } from "./document.js";
import type { ParagraphBox, PlacedLine } from "./stack.js";
import { twipsToPoints } from "./units.js";

/**
 * Where a page came out wrong on its own terms, with nothing to compare it against.
 *
 * **This is not `unhonoured`, and the two must not be read as one list.** `unhonoured`
 * says what the document asked for and did not get: a wave border drawn as a plain
 * line, a face stood in, a kerning rule nothing here builds. Every entry of it is a
 * statement about fidelity, most of them are invisible, and a document stating a dozen
 * of them can still be drawn exactly as Word drew it. What follows is the opposite kind
 * of statement: **no document can ask for text off the sheet, text above the top of its
 * own page, or text over other text**, so a page doing one of those is wrong however
 * the document was written and whatever Word would have done with it. Nothing here
 * needs Word, a pdf, a corpus or a rasteriser: the layout says it about itself, in the
 * browser, on one document, which is what lets an application show a preview or fall
 * back rather than show a broken page.
 *
 * `unshowable` because that is the decision it exists to answer, and the word already
 * means this in `corpus/deformed.ts`: a page that is moved is the page Word drew put
 * down wrong and can be shown, and a page that is deformed cannot. This is the same
 * judgement reached without Word.
 *
 * **The naive version of this check cries wolf, and the calibration is the work.** Read
 * over the 718 corpus documents on 2026-08-12, "a line above the body top" named 81 of
 * them, and the raster said 45 of the 72 clean offenders are drawn cell for cell as
 * Word drew them: **whatever put the line up there, nobody can see it.** A check that
 * flags a page Word drew identically is worse than no check. Two things were wrong with
 * it, and both are measured rather than argued below: it read every page against the
 * body top of the **first** page, which is why 43 of the 81 were out by the same 93.8pt,
 * the height of a header; and a line above the top by less than a line of text is
 * invisible. Named against the raster after both, it comes to **11 documents of the clean
 * 580, and not one of the pages it names is drawn as Word drew it**: the mildest differs
 * by two tenths of a percent of its cells and the rest by 7% to 67%.
 */
export type UnshowableKind = "text-off-the-sheet" | "text-above-the-body" | "text-over-text";

export type Unshowable = {
  readonly kind: UnshowableKind;
  readonly page: number;
  // How many of the page's lines are wrong this way, and by how far the worst of them
  // is: off the sheet, above the top, or into another line.
  readonly lines: number;
  readonly worstPt: number;
};

// How far a line may stand outside where it belongs before anybody could see it.
//
// A point and a half, which is the grid Word's own positions are read on and under a
// fifth of a line. It is what keeps a line resting exactly on an edge out of this.
const NOBODY_SEES_PT = 1.5;

// How far above the top of its own page a line has to be drawn before the page is worth
// refusing to show.
//
// **A whole line, and the corpus is what chose it.** Read over the 580 clean documents
// on 2026-08-12 and joined to the raster page by page, the two populations do not
// overlap and there is nothing in between: every page named at under 7pt is one the
// raster says is drawn as Word drew it (five of them cell for cell, the rest inside two
// percent of the page), and every page where it matters is out by **49pt or more**, up
// to 270. Sixteen documents named at a point and a half become eight named at a line,
// and the raster calls every one of those eight pages between a third and two thirds
// wrong.
//
// The pages this now lets through are not pages nothing is wrong with. They are pages
// where what is wrong is invisible, which is the raster's business and the line score's,
// and this exists to answer one question: is the preview fit to put in front of somebody.
const ABOVE_THE_TOP_PT = 12;

// How much of two lines has to lie in the same place before the two are over each
// other rather than near each other.
//
// **Text drawn over text is the fault this whole reading exists for**, and it is also
// the easiest to imagine where there is none: adjacent lines share an edge, a line's
// descender reaches into the next line's ascent by design, and a paragraph told exactly
// how tall its lines are has Word itself drawing them into each other. So a pair counts
// only where the overlap is a real share of the shorter of the two, in both directions
// at once.
const OVER_EACH_OTHER = 0.35;

// The box the line's own text is drawn in. The line's room is not its ink: a rule can
// open room above the text and a multiple hangs room below it, and neither is drawn.
type Ink = {
  readonly leftPt: number;
  readonly rightPt: number;
  readonly topPt: number;
  readonly bottomPt: number;
};

const carriesInk = (line: PlacedLine): boolean =>
  line.line.segments.some((segment) => segment.kind === "text" && segment.text.trim() !== "");

function inkOf(line: PlacedLine): Ink {
  const aboveTheBaseline = line.line.ascentPt - line.line.seatPt;
  return {
    leftPt: line.leftPt,
    rightPt: line.leftPt + line.line.widthPt,
    topPt: line.baselinePt - aboveTheBaseline,
    bottomPt: line.baselinePt + (line.line.heightPt - line.line.ascentPt),
  };
}

const overlapOf = (one: number, other: number, ownerOne: number, ownerOther: number): number =>
  Math.min(ownerOne, ownerOther) - Math.max(one, other);

// Two lines are over each other where a real share of both lies in the same place. The
// share is taken of the smaller of the two, so a word drawn over a whole page of text
// is caught as surely as two pages of text drawn over each other.
function overEachOther(one: Ink, other: Ink): boolean {
  const down = overlapOf(one.topPt, other.topPt, one.bottomPt, other.bottomPt);
  const across = overlapOf(one.leftPt, other.leftPt, one.rightPt, other.rightPt);
  if (down <= 0 || across <= 0) return false;

  const shorter = Math.min(one.bottomPt - one.topPt, other.bottomPt - other.topPt);
  const narrower = Math.min(one.rightPt - one.leftPt, other.rightPt - other.leftPt);
  return (
    shorter > 0 &&
    narrower > 0 &&
    down / shorter >= OVER_EACH_OTHER &&
    across / narrower >= OVER_EACH_OTHER
  );
}

// The lines the flow of the page drew, which is what this asks about and not everything
// drawn on it. A header lives in the top margin and a box may be anchored anywhere at
// all, both by the document's own instruction; the flow is the story that has a top and
// a foot to be outside of.
const flowLinesOf = (page: LaidOutPage): readonly PlacedLine[] =>
  page.body.flatMap((box: ParagraphBox) => box.lines.filter(carriesInk));

// Lines of one paragraph are left out of the comparison with each other. A paragraph
// told exactly how tall its lines are draws them into each other and Word draws them
// into each other too, which is the document being honoured rather than a page coming
// out wrong.
//
// **Read down the page rather than pair by pair.** A page of the corpus holds 295 lines
// and a page built out of one long story holds thousands; comparing every pair of them
// is a run this cannot afford in a browser, and the first cut of it spent half an hour
// and four gigabytes on the corpus before it was killed. Sorted by their tops, a line
// can only be over the ones that start before its own foot, so the walk stops at the
// first that starts below it.
function overOtherTextOn(page: LaidOutPage): { lines: number; worstPt: number } {
  const inks = page.body
    .flatMap((box, at) => box.lines.filter(carriesInk).map((line) => ({ at, ink: inkOf(line) })))
    .sort((one, other) => one.ink.topPt - other.ink.topPt);

  const over = new Set<Ink>();
  let worstPt = 0;

  for (let at = 0; at < inks.length; at += 1) {
    const one = inks[at];
    if (one === undefined) continue;
    for (let next = at + 1; next < inks.length; next += 1) {
      const other = inks[next];
      if (other === undefined || other.ink.topPt >= one.ink.bottomPt) break;
      if (other.at === one.at) continue;
      if (!overEachOther(one.ink, other.ink)) continue;
      over.add(one.ink);
      over.add(other.ink);
      worstPt = Math.max(
        worstPt,
        overlapOf(one.ink.topPt, other.ink.topPt, one.ink.bottomPt, other.ink.bottomPt),
      );
    }
  }

  return { lines: over.size, worstPt };
}

// Only a line drawn **wholly** outside the sheet counts, and the near edge is what it
// is measured by.
//
// A line hanging over an edge is not this fault: a table wider than its page is
// ordinary in the wild, Word draws it past the right edge exactly as we do, and the pdf
// cuts both of them off at the same place. So a page cannot be called unshowable for
// something Word does too. A line drawn entirely off the sheet is a different thing:
// nobody can see any part of it, so whatever it says is lost.
function offTheSheetOn(page: LaidOutPage): { lines: number; worstPt: number } {
  const widthPt = twipsToPoints(page.geometry.widthTwips);
  const heightPt = twipsToPoints(page.geometry.heightTwips);

  let lines = 0;
  let worstPt = 0;
  for (const line of flowLinesOf(page)) {
    const ink = inkOf(line);
    const outside = Math.max(
      -ink.bottomPt,
      ink.topPt - heightPt,
      -ink.rightPt,
      ink.leftPt - widthPt,
    );
    if (outside <= NOBODY_SEES_PT) continue;
    lines += 1;
    worstPt = Math.max(worstPt, outside);
  }
  return { lines, worstPt };
}

function aboveTheBodyOn(page: LaidOutPage): { lines: number; worstPt: number } {
  let lines = 0;
  let worstPt = 0;
  for (const line of flowLinesOf(page)) {
    const above = page.bodyTopPt - inkOf(line).topPt;
    if (above <= ABOVE_THE_TOP_PT) continue;
    lines += 1;
    worstPt = Math.max(worstPt, above);
  }
  return { lines, worstPt };
}

/**
 * Every way the pages of this layout came out wrong on their own terms, one entry per
 * page and kind. Empty is the answer for a page worth showing.
 *
 * An application shows the preview when this is empty and falls back when it is not,
 * so **an entry here has to be worth a fallback**: everything that is merely
 * unfaithful belongs in `unhonoured`, and everything invisible belongs in neither.
 */
export function unshowableIn(layout: LaidOutDocument): readonly Unshowable[] {
  const found: Unshowable[] = [];

  for (const page of layout.pages) {
    const readings: readonly (readonly [UnshowableKind, { lines: number; worstPt: number }])[] = [
      ["text-off-the-sheet", offTheSheetOn(page)],
      ["text-above-the-body", aboveTheBodyOn(page)],
      ["text-over-text", overOtherTextOn(page)],
    ];
    for (const [kind, read] of readings) {
      if (read.lines > 0) {
        found.push({ kind, page: page.index, lines: read.lines, worstPt: read.worstPt });
      }
    }
  }

  return found;
}
