import { describe, expect, it } from "vitest";

import type { SectionGeometry } from "../docx/section.js";
import type { ParagraphMark } from "../docx/styles.js";
import type { LaidOutDocument, LaidOutPage } from "./document.js";
import { WHOLE_FRAME } from "../docx/anchors.js";
import { NO_PAINT, type TextBoxBody } from "../docx/drawing.js";
import { UNPAINTED, type PlacedContent, type PlacedFloat } from "./floats.js";
import type { ParagraphBox, PlacedLine } from "./stack.js";
import { unshowableIn } from "./unshowable.js";

// A page read against itself, over layouts written by hand. Everything here is
// geometry, so the cases are geometry: a document laid out from a `.docx` cannot be
// made to draw a line 400pt above its own top without a rule that puts it there, and
// what is being pinned is the reading rather than any of those rules.

const LETTER: SectionGeometry = {
  widthTwips: 12240,
  heightTwips: 15840,
  margin: {
    topTwips: 1440,
    rightTwips: 1440,
    bottomTwips: 1440,
    leftTwips: 1440,
    headerTwips: 720,
    footerTwips: 720,
  },
};

const MARK: ParagraphMark = {
  font: { kind: "named", name: "Twin Sans" },
  fontSizePt: 12,
  bold: false,
  italic: false,
  underline: false,
  raisePt: 0,
  lineSizePt: 12,
  lineRaisePt: 0,
  color: null,
  characterSpacingPt: 0,
  characterScale: 1,
  kernFromHalfPoints: null,
};

// Twelve points tall, its baseline ten below its own top, so the ink of a line asked
// for at 100 runs from 100 to 112 and nothing has to be worked out twice.
const HEIGHT_PT = 12;
const ASCENT_PT = 10;

const lineAt = (topPt: number, leftPt: number, widthPt = 100, text = "something"): PlacedLine => ({
  line: {
    segments: [{ kind: "text", mark: MARK, text, widthPt, offsetPt: 0 }],
    widthPt,
    heightPt: HEIGHT_PT,
    ascentPt: ASCENT_PT,
    seatPt: 0,
    fontHeightPt: HEIGHT_PT,
    heldOpenPt: null,
  },
  leftPt,
  topPt,
  heightPt: HEIGHT_PT,
  seatPt: 0,
  fittingHeightPt: HEIGHT_PT,
  baselinePt: topPt + ASCENT_PT,
  startsPage: false,
});

const paragraphOf = (lines: readonly PlacedLine[], index = 0): ParagraphBox => ({
  index,
  topPt: lines[0]?.topPt ?? 0,
  anchorTopPt: lines[0]?.topPt ?? 0,
  resumesUnderPt: 0,
  heightPt: HEIGHT_PT * lines.length,
  lines,
  marker: null,
  markTopPt: lines[0]?.topPt ?? 0,
  contentBottomPt: (lines[0]?.topPt ?? 0) + HEIGHT_PT * lines.length,
  widowControl: true,
  keepNext: false,
  startsPage: false,
  endsPage: false,
  endsPageAtASection: false,
  contentWidthPt: 100,
  clipTo: null,
  paint: null,
});

const EMPTY_BODY: TextBoxBody = {
  blocks: [],
  insets: { leftEmu: 0, topEmu: 0, rightEmu: 0, bottomEmu: 0 },
  anchor: "top",
  wraps: true,
  fitsText: false,
};

const boxHolding = (boxes: readonly ParagraphBox[]): PlacedContent => ({
  kind: "text-box",
  body: EMPTY_BODY,
  text: { boxes, cells: [], inlines: [], contentHeightPt: 0, contentWidthPt: 0 },
  paint: UNPAINTED,
});

const floatOf = (content: PlacedContent): PlacedFloat => ({
  anchor: {
    paragraphIndex: 0,
    name: "Object",
    widthEmu: 0,
    heightEmu: 0,
    turnDegrees: 0,
    flip: { horizontal: false, vertical: false },
    content: { kind: "shape", paint: NO_PAINT },
    horizontal: { kind: "offset", from: "column", offsetEmu: 0 },
    vertical: { kind: "offset", from: "paragraph", offsetEmu: 0 },
    wrap: "none",
    side: "bothSides",
    area: WHOLE_FRAME,
    distances: { topEmu: 0, rightEmu: 0, bottomEmu: 0, leftEmu: 0 },
    behindDoc: false,
    relativeHeight: 0,
  },
  content,
  leftPt: 0,
  topPt: 0,
  widthPt: 100,
  heightPt: 100,
  turnDegrees: 0,
  flip: { horizontal: false, vertical: false },
});

type PageParts = {
  readonly body?: readonly ParagraphBox[];
  readonly header?: readonly ParagraphBox[];
  readonly floats?: readonly PlacedFloat[];
  readonly bodyTopPt?: number;
};

const pageOf = (index: number, parts: PageParts): LaidOutPage => ({
  index,
  geometry: LETTER,
  body: parts.body ?? [],
  cells: [],
  floats: parts.floats ?? [],
  inlines: [],
  headerTopPt: 36,
  headerHeightPt: 0,
  footerTopPt: 756,
  bodyTopPt: parts.bodyTopPt ?? 72,
  bodyBottomPt: 720,
  header: parts.header ?? [],
  footer: [],
  headerCells: [],
  footerCells: [],
  headerFloats: [],
  footerFloats: [],
  headerInlines: [],
  footerInlines: [],
});

const laidOut = (pages: readonly LaidOutPage[]): LaidOutDocument => ({
  kind: "laid-out",
  page: LETTER,
  unhonoured: [],
  headerTopPt: 36,
  bodyTopPt: 72,
  bodyBottomPt: 720,
  pages,
});

const readingOf = (parts: PageParts) => unshowableIn(laidOut([pageOf(0, parts)]));

describe("a page nothing is wrong with", () => {
  it("says nothing at all", () => {
    const body = [
      paragraphOf([lineAt(72, 72), lineAt(84, 72)], 0),
      paragraphOf([lineAt(96, 72)], 1),
    ];

    expect(readingOf({ body })).toStrictEqual([]);
  });

  // The line drawn against the very top of the body is the commonest line there is,
  // and a reading that rounds it the wrong way names every document ever written.
  it("says nothing about the line standing exactly on the body top", () => {
    expect(readingOf({ body: [paragraphOf([lineAt(72, 72)])] })).toStrictEqual([]);
  });

  it("says nothing about a line whose text is nothing but spaces", () => {
    const body = [paragraphOf([lineAt(-200, 72, 100, "   ")])];

    expect(readingOf({ body })).toStrictEqual([]);
  });
});

describe("text above the top of its own page", () => {
  it("names the page and says how far above", () => {
    const body = [paragraphOf([lineAt(52, 72)]), paragraphOf([lineAt(72, 72)], 1)];

    expect(readingOf({ body })).toStrictEqual([
      { kind: "text-above-the-body", page: 0, lines: 1, worstPt: 20 },
    ]);
  });

  // **The two populations do not overlap, and this is the gap between them.** Over the
  // clean corpus every page named at under 7pt is one the raster says is drawn as Word
  // drew it, and every page where it matters is out by 49pt or more. A line above its
  // own top by half a line is a line nobody sees.
  it("says nothing about a line above the top by less than a line of text", () => {
    expect(readingOf({ body: [paragraphOf([lineAt(66, 72)])] })).toStrictEqual([]);
  });

  // **The reading that cried wolf.** 43 of the 81 offenders were out by the same
  // 93.8pt, which is the height of a header: the check was comparing every page
  // against the body top of the first, and a page that draws no header, or one of
  // another section, starts its body somewhere else entirely.
  it("reads each page against its own body top rather than the document's", () => {
    const pages = [
      pageOf(0, { bodyTopPt: 165.8, body: [paragraphOf([lineAt(165.8, 72)])] }),
      pageOf(1, { bodyTopPt: 72, body: [paragraphOf([lineAt(72, 72)])] }),
    ];

    expect(unshowableIn(laidOut(pages))).toStrictEqual([]);
  });

  it("says nothing about a header, which is drawn in the margin by instruction", () => {
    const header = [paragraphOf([lineAt(36, 72)])];

    expect(readingOf({ header, body: [paragraphOf([lineAt(72, 72)])] })).toStrictEqual([]);
  });

  it("says nothing about a text box, which may be anchored anywhere at all", () => {
    const floats = [floatOf(boxHolding([paragraphOf([lineAt(20, 72)])]))];

    expect(readingOf({ floats, body: [paragraphOf([lineAt(72, 72)])] })).toStrictEqual([]);
  });
});

describe("text off the sheet", () => {
  // The sheet is 792pt tall, so a line whose own top is at 800 has no part of itself
  // on the page at all.
  it("names a line drawn past the foot of the sheet", () => {
    const body = [paragraphOf([lineAt(800, 72)])];

    expect(readingOf({ body })).toStrictEqual([
      { kind: "text-off-the-sheet", page: 0, lines: 1, worstPt: 8 },
    ]);
  });

  it("names a line drawn past the right edge of the sheet", () => {
    const body = [paragraphOf([lineAt(72, 700, 100)])];

    expect(readingOf({ body })).toStrictEqual([
      { kind: "text-off-the-sheet", page: 0, lines: 1, worstPt: 88 },
    ]);
  });

  // A table wider than its page is ordinary, Word draws it over the right edge as we
  // do, and the pdf cuts both off in the same place.
  it("says nothing about a line hanging over an edge", () => {
    const body = [paragraphOf([lineAt(72, 560, 100)])];

    expect(readingOf({ body })).toStrictEqual([]);
  });

  it("names a line off the left edge and above the top once each", () => {
    const body = [paragraphOf([lineAt(-30, -200)])];

    expect(readingOf({ body }).map((each) => each.kind)).toStrictEqual([
      "text-off-the-sheet",
      "text-above-the-body",
    ]);
  });
});

describe("text over other text", () => {
  it("names both lines and says how far into each other they are", () => {
    const body = [paragraphOf([lineAt(100, 72)]), paragraphOf([lineAt(106, 72)], 1)];

    expect(readingOf({ body })).toStrictEqual([
      { kind: "text-over-text", page: 0, lines: 2, worstPt: 6 },
    ]);
  });

  it("says nothing about two lines that merely stand one under the other", () => {
    const body = [paragraphOf([lineAt(100, 72)]), paragraphOf([lineAt(112, 72)], 1)];

    expect(readingOf({ body })).toStrictEqual([]);
  });

  // A descender reaching into the line below is how type is set, not a page coming
  // out wrong: three points of twelve is under the share a pair has to share.
  it("says nothing where the two barely reach into each other", () => {
    const body = [paragraphOf([lineAt(100, 72)]), paragraphOf([lineAt(109, 72)], 1)];

    expect(readingOf({ body })).toStrictEqual([]);
  });

  it("says nothing about two lines side by side, which is what a table row is", () => {
    const body = [paragraphOf([lineAt(100, 72, 100)]), paragraphOf([lineAt(100, 200, 100)], 1)];

    expect(readingOf({ body })).toStrictEqual([]);
  });

  // A paragraph told exactly how tall its lines are draws them into each other, and
  // Word draws them into each other too: that is the document being honoured.
  it("says nothing about two lines of one paragraph", () => {
    const body = [paragraphOf([lineAt(100, 72), lineAt(104, 72)])];

    expect(readingOf({ body })).toStrictEqual([]);
  });

  it("counts a page of text drawn over another page of it as every line of both", () => {
    const first = paragraphOf([lineAt(100, 72), lineAt(112, 72), lineAt(124, 72)], 0);
    const over = paragraphOf([lineAt(100, 72), lineAt(112, 72), lineAt(124, 72)], 1);

    expect(readingOf({ body: [first, over] })).toStrictEqual([
      { kind: "text-over-text", page: 0, lines: 6, worstPt: 12 },
    ]);
  });
});
