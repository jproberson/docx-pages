import type { LaidOutDocument, LaidOutPage, ParagraphBox, PlacedLine } from "@docx-pages/core";

import type { TextPlacement } from "./text.js";

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

// The header and the footer are drawn again on every page, so every page holds
// their boxes as well as its own.
function boxesOnPage(layout: LaidOutDocument, page: LaidOutPage): readonly ParagraphBox[] {
  const floats = [...layout.headerFloats, ...layout.footerFloats, ...page.floats];
  const inBoxes = floats.flatMap((float) =>
    float.content.kind === "text-box" && float.content.text !== null
      ? [...float.content.text.boxes]
      : [],
  );
  return [...layout.header, ...layout.footer, ...page.body, ...inBoxes];
}

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

function compareNumbers(
  tolerancePt: number,
  boxes: readonly ParagraphBox[],
  drawn: readonly TextPlacement[],
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
        normalise(item.text) === text &&
        Math.abs(item.baselinePt - line.baselinePt) <= SAME_LINE_PT,
    );
    if (near.length === 0) continue;
    numbersMatched += 1;

    const off = Math.min(...near.map((item) => Math.abs(item.leftPt - marker.leftPt)));
    if (off <= tolerancePt) numbersPlaced += 1;
  }

  return { numbersMatched, numbersPlaced };
}

export function agreementWith(
  layout: LaidOutDocument,
  drawn: readonly TextPlacement[],
  tolerancePt: number,
): Agreement {
  const taken = new Set<TextPlacement>();
  const elsewhere: string[] = [];
  const scales: number[] = [];
  let lines = 0;
  let matched = 0;
  let placed = 0;
  let runsMatched = 0;
  let runsPlaced = 0;
  let numbersMatched = 0;
  let numbersPlaced = 0;

  for (const page of layout.pages) {
    const onPage = drawn.filter((item) => item.pageIndex === page.index);
    const boxes = boxesOnPage(layout, page);

    for (const line of boxes.flatMap((box) => box.lines)) {
      const text = normalise(textOf(line));
      if (text === "") continue;
      lines += 1;

      const items = itemsFor(text, onPage, taken);
      const found = items?.[0];
      if (items === null || found === undefined) {
        elsewhere.push(text);
        continue;
      }
      for (const item of items) taken.add(item);
      matched += 1;

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
      const off = Math.max(
        Math.abs((theirs.get(0) ?? found.leftPt) - (ours.get(0) ?? line.leftPt)),
        Math.abs(found.baselinePt - firstBaselineOf(line)),
      );
      if (off <= tolerancePt) placed += 1;

      for (const [at, leftPt] of theirs) {
        const mine = ours.get(at);
        if (mine === undefined) continue;
        runsMatched += 1;
        if (Math.abs(mine - leftPt) <= tolerancePt) runsPlaced += 1;
      }
    }

    const numbers = compareNumbers(tolerancePt, boxes, onPage);
    numbersMatched += numbers.numbersMatched;
    numbersPlaced += numbers.numbersPlaced;
  }

  // A line Word drew on another page was still broken the way Word broke it.
  for (const text of elsewhere) {
    const items = itemsFor(text, drawn, taken);
    if (items === null) continue;
    for (const item of items) taken.add(item);
    matched += 1;
  }

  return {
    lines,
    matched,
    placed,
    runsMatched,
    runsPlaced,
    numbersMatched,
    numbersPlaced,
    drawnScale: middleOf(scales),
  };
}
