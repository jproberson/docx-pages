import { describe, expect, it } from "vitest";

import {
  bestEffortMetrics,
  layOutDocument,
  openDocx,
  type FaceDefaults,
  type LaidOutDocument,
  type PlacedLine,
} from "@docx-pages/core";
import { buildDocx, buildFace, WORDPROCESSING_NS } from "@docx-pages/core/testing";

import { agreementWith, type PageShape } from "./agreement.js";
import type { TextPlacement } from "./text.js";

// What a page's lines say when they are read together, over a document laid out here
// and a drawing made out of that layout: the drawing agrees with it by construction,
// and each case moves the drawing the one way the shape it is meant to be called by
// says. Word is not asked anything, because none of this is about Word's rules. It is
// about telling a page that is moved from a page that is broken, which is a reading
// of two sets of positions and nothing else.

const METRICS = { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 };
const TWIN = buildFace({ name: "Twin Sans", metrics: METRICS, sansSerif: true });

const DEFAULTS: FaceDefaults = {
  faces: [TWIN],
  twins: {},
  sansSerif: "Twin Sans",
  serif: "Twin Sans",
  monospace: "Twin Sans",
  lastResort: "Twin Sans",
};

const TOLERANCE_PT = 1;

// Each paragraph says something no other paragraph says, so a line can only be
// matched to the item drawn for it: the matcher spells a line out of the items in
// front of it, and two lines reading alike would let it spell one out of the other's.
const WORDS = [
  "alpha",
  "bravo",
  "charlie",
  "delta",
  "echo",
  "foxtrot",
  "golf",
  "hotel",
  "india",
  "juliet",
  "kilo",
  "lima",
  "mike",
  "november",
  "oscar",
  "papa",
  "quebec",
  "romeo",
  "sierra",
  "tango",
];

function laidOut(paragraphs: readonly string[]): LaidOutDocument {
  const body = paragraphs.map((word) => `<w:p><w:r><w:t>${word}</w:t></w:r></w:p>`).join("");
  const bytes = buildDocx({
    "word/document.xml":
      `<?xml version="1.0"?><w:document xmlns:w="${WORDPROCESSING_NS}">` +
      `<w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"` +
      ` w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  });
  const laid = layOutDocument(openDocx(bytes), bestEffortMetrics([], DEFAULTS));
  if (laid.kind !== "laid-out") throw new Error(`blocked: ${laid.blocker.kind}`);
  return laid;
}

const textOf = (line: PlacedLine): string =>
  line.line.segments.map((segment) => (segment.kind === "text" ? segment.text : "")).join("");

// The drawing our own layout would make of itself, which is the drawing that agrees
// with it about everything. A case moves this rather than the layout, so whatever the
// layout did the difference between the two is only what the case put there.
type Moved = (line: PlacedLine, pageIndex: number) => { leftPt: number; baselinePt: number };

const asDrawn = (laid: LaidOutDocument, moved: Moved): readonly TextPlacement[] =>
  laid.pages.flatMap((page) =>
    page.body.flatMap((box) =>
      box.lines.flatMap((line) => {
        const text = textOf(line);
        if (text.trim() === "") return [];
        const { leftPt, baselinePt } = moved(line, page.index);
        return [
          {
            kind: "text" as const,
            text,
            fontName: "Twin Sans",
            pageIndex: page.index,
            leftPt,
            baselinePt,
            // The same advance our own line put down, so that the two sides carry the
            // same ink and no case here is about content one of them drew alone.
            widthPt: line.line.widthPt,
            fontSizePt: 11,
          },
        ];
      }),
    ),
  );

const where = (line: PlacedLine) => ({ leftPt: line.leftPt, baselinePt: line.baselinePt });

const shapesOf = (laid: LaidOutDocument, drawn: readonly TextPlacement[]): readonly PageShape[] =>
  agreementWith(laid, drawn, TOLERANCE_PT).pages.map((page) => page.shape);

describe("what a page's lines say read together", () => {
  const laid = laidOut(WORDS);

  it("agrees where every line landed where it was drawn", () => {
    expect(shapesOf(laid, asDrawn(laid, where))).toStrictEqual(["agrees"]);
  });

  it("agrees where every line is out by less than anybody could see", () => {
    const drawn = asDrawn(laid, (line) => ({
      leftPt: line.leftPt + 0.4,
      baselinePt: line.baselinePt - 0.6,
    }));

    expect(shapesOf(laid, drawn)).toStrictEqual(["agrees"]);
  });

  // **The wolf-cry the raster caught.** Displacements this small neither cluster nor
  // drift, and the first cut of this called such a page deformed: over the clean
  // corpus, ten of the seventeen pages it named on nothing further out than 3pt are
  // drawn cell for cell as Word drew them. A page is not deformed by a scatter nobody
  // can see, however unexplained the scatter is.
  it("agrees where the lines are out by a scatter under a quarter of a line", () => {
    const scattered = [
      0, 2.4, -1.9, 2.8, 0.4, -2.7, 2.9, 1.1, -2.2, 1.6, 2.5, -1.4, 1.9, -2.3, 0.5, 2.7, -0.6, 1.2,
      2.6, 0.2,
    ];
    const drawn = asDrawn(laid, (line) => ({
      leftPt: line.leftPt,
      baselinePt: line.baselinePt + (scattered[WORDS.indexOf(textOf(line))] ?? 0),
    }));

    expect(shapesOf(laid, drawn)).toStrictEqual(["agrees"]);
  });

  // A page break in the wrong place: the whole story moves down the page and nothing
  // moves across it, which is the one signature worth reading off this at a glance.
  it("calls a page shifted where one offset explains every line", () => {
    const drawn = asDrawn(laid, (line) => ({
      leftPt: line.leftPt,
      baselinePt: line.baselinePt + 13.8,
    }));
    const page = agreementWith(laid, drawn, TOLERANCE_PT).pages[0];

    expect(page?.shape).toBe("shifted");
    expect(page?.offsetPt?.downPt).toBeCloseTo(13.8, 6);
    expect(page?.offsetPt?.leftPt).toBeCloseTo(0, 6);
    expect(page?.explained).toBe(1);
  });

  // Nineteen of the twenty agree about 13.8, which is the share a shift has to
  // account for exactly: the page is still one fault plus a line of its own.
  it("keeps calling a page shifted where one line in twenty dissents", () => {
    const drawn = asDrawn(laid, (line) => ({
      leftPt: line.leftPt,
      baselinePt: line.baselinePt + (textOf(line) === "golf" ? 40 : 13.8),
    }));
    const page = agreementWith(laid, drawn, TOLERANCE_PT).pages[0];

    expect(page?.shape).toBe("shifted");
    expect(page?.offsetPt?.downPt).toBeCloseTo(13.8, 6);
  });

  // **A short page is judged more harshly than a long one and that is not a fault in
  // the reading.** A dissenter is a twelfth of twelve lines and a fortieth of forty,
  // so the same single line takes a short page out of every cluster it could have
  // been in. A page with two faults on it is deformed however few lines it holds.
  it("calls a short page deformed where the same one line dissents", () => {
    const short = laidOut(WORDS.slice(0, 12));
    const drawn = asDrawn(short, (line) => ({
      leftPt: line.leftPt,
      baselinePt: line.baselinePt + (textOf(line) === "golf" ? 40 : 13.8),
    }));

    expect(shapesOf(short, drawn)).toStrictEqual(["deformed"]);
  });

  // A height that is wrong by a fraction is paid once a line, so it reaches the foot
  // of the page as a slope rather than as an offset.
  it("calls a page drifting where the offset grows with the line's own baseline", () => {
    const first = laid.pages[0]?.body[0]?.lines[0]?.baselinePt ?? 0;
    const drawn = asDrawn(laid, (line) => ({
      leftPt: line.leftPt,
      baselinePt: line.baselinePt + (line.baselinePt - first) * 0.1,
    }));
    const page = agreementWith(laid, drawn, TOLERANCE_PT).pages[0];

    expect(page?.shape).toBe("drifting");
    expect(page?.driftPerPt).toBeCloseTo(0.1, 6);
  });

  // **The reading that called a readable page unreadable.** Page 7 of one corpus
  // document is a table of 216 cells whose rows are each 0.95pt shorter than Word's,
  // reaching 24pt by the foot of the page: drawn in the right order and perfectly
  // readable. A straight line fitted through the page called it deformed, because the
  // squeeze starts at the table rather than at the top of the sheet and no one slope
  // fits both the lines above it and the rows below.
  it("calls a page drifting where the squeeze starts partway down it", () => {
    const drawn = asDrawn(laid, (line) => {
      const at = WORDS.indexOf(textOf(line));
      return {
        leftPt: line.leftPt,
        baselinePt: line.baselinePt - Math.max(0, at - 5) * 0.95,
      };
    });
    const page = agreementWith(laid, drawn, TOLERANCE_PT).pages[0];

    expect(page?.shape).toBe("drifting");
    expect(page?.worstPt).toBeGreaterThan(12);
  });

  // The line that jumped past another. Nothing that moves a page as a whole can do it,
  // and it is what text drawn over text looks like from here.
  it("calls a page deformed where one line crossed another", () => {
    const swapped: Readonly<Record<string, number>> = { golf: 60, november: -60 };
    const drawn = asDrawn(laid, (line) => ({
      leftPt: line.leftPt,
      baselinePt: line.baselinePt + (swapped[textOf(line)] ?? 0),
    }));

    expect(shapesOf(laid, drawn)).toStrictEqual(["deformed"]);
  });

  // The same total error spread about at random rather than accumulating, which is
  // the page a preview cannot show: nothing one fault could have done put the lines
  // where they are.
  it("calls a page deformed where the offsets neither cluster nor drift", () => {
    const scattered = [0, 9, -4, 22, 3, -17, 30, 6, -11, 14, 25, -8, 19, -13, 5, 27, -6, 11, 33, 2];
    const drawn = asDrawn(laid, (line) => ({
      leftPt: line.leftPt,
      baselinePt: line.baselinePt + (scattered[WORDS.indexOf(textOf(line))] ?? 0),
    }));

    expect(shapesOf(laid, drawn)).toStrictEqual(["deformed"]);
  });

  it("calls a page deformed where the lines are out across it as well as down", () => {
    const across = [0, 14, -6, 25, 4, -19, 33, 7, -12, 16, 28, -9, 21, -15, 6, 30, -7, 12, 36, 3];
    const drawn = asDrawn(laid, (line) => ({
      leftPt: line.leftPt + (across[WORDS.indexOf(textOf(line))] ?? 0),
      baselinePt: line.baselinePt,
    }));

    expect(shapesOf(laid, drawn)).toStrictEqual(["deformed"]);
  });

  // **The second wolf-cry, and the reason `missing` is decided by ink.** A line Word
  // broke into pieces this cannot spell out of them goes unmatched with nothing wrong
  // on the page at all: over the clean corpus nine pages called missing on unmatched
  // lines alone are drawn cell for cell as Word drew them, and three quarters of a
  // page can go unmatched on one that is perfect. Here nothing matches and both sides
  // carry the same ink, which is a page nobody drew wrong.
  it("says nothing is missing where nothing matched and both sides drew the same ink", () => {
    const drawn = asDrawn(laid, where).map((item) => ({ ...item, text: `${item.text}!` }));
    const page = agreementWith(laid, drawn, TOLERANCE_PT).pages[0];

    expect(page?.matched).toBe(0);
    expect(page?.oursAlone).toBe(20);
    expect(page?.theirsAlone).toBe(20);
    expect(page?.shape).toBe("agrees");
  });

  // Text we drew that Word drew nowhere is a real fault and it is not this one: our own
  // line carries the trailing spaces Word draws as items of their own and this reading
  // throws away, so our side reads heavier than Word's on a page with nothing wrong
  // (2.7% heavier on one page of the eight known to be right). Nobody has measured what
  // that noise costs, so nothing is claimed and the counts say where to look.
  it("leaves a page agreed where we drew more than Word, and says so in the counts", () => {
    const drawn = asDrawn(laid, where).filter((item) => !WORDS.slice(4).includes(item.text));
    const page = agreementWith(laid, drawn, TOLERANCE_PT).pages[0];

    expect(page?.shape).toBe("agrees");
    expect(page?.oursAlone).toBe(16);
    expect(page?.matched).toBe(4);
  });

  it("calls a page missing where Word drew visibly more ink than we did", () => {
    const drawn = [
      ...asDrawn(laid, where),
      ...WORDS.map((word) => ({
        kind: "text" as const,
        text: `${word} again`,
        fontName: "Twin Sans",
        pageIndex: 0,
        leftPt: 400,
        baselinePt: 100,
        widthPt: 20,
        fontSizePt: 11,
      })),
    ];
    const page = agreementWith(laid, drawn, TOLERANCE_PT).pages[0];

    expect(page?.theirs).toBe(40);
    expect(page?.theirsAlone).toBe(20);
    expect(page?.shape).toBe("missing");
  });

  it("says how many pages the drawing holds, so a page we never made is not silent", () => {
    const drawn = [
      ...asDrawn(laid, where),
      {
        kind: "text" as const,
        text: "mike",
        fontName: "Twin Sans",
        pageIndex: 3,
        leftPt: 72,
        baselinePt: 100,
        widthPt: 20,
        fontSizePt: 11,
      },
    ];
    const agreed = agreementWith(laid, drawn, TOLERANCE_PT);

    expect(agreed.pages).toHaveLength(1);
    expect(agreed.pagesDrawn).toBe(4);
  });
});

// A page whose lines Word drew on the page before is a break in the wrong place, and
// the text is all there: it must not read as content nobody drew.
describe("a line the drawing put on another page", () => {
  it("counts it against the page it stands on rather than as missing", () => {
    const laid = laidOut(WORDS);
    const drawn = asDrawn(laid, where).map((item) =>
      item.text === "lima" ? { ...item, pageIndex: 1 } : item,
    );
    const page = agreementWith(laid, drawn, TOLERANCE_PT).pages[0];

    expect(page?.onAnotherPage).toBe(1);
    expect(page?.oursAlone).toBe(0);
    expect(page?.shape).toBe("agrees");
  });
});

describe("a page with almost nothing on it", () => {
  // One line agrees with itself about whatever offset it is out by, so a page holding
  // one is called shifted by that offset and `matched` beside it is the whole of the
  // evidence there is for saying so. A ranking read off this weighs it.
  it("reports the offset of a page holding one line, and says it is one line", () => {
    const laid = laidOut(["alpha"]);
    const drawn = asDrawn(laid, (line) => ({
      leftPt: line.leftPt,
      baselinePt: line.baselinePt + 40,
    }));
    const page = agreementWith(laid, drawn, TOLERANCE_PT).pages[0];

    expect(page?.matched).toBe(1);
    expect(page?.shape).toBe("shifted");
    expect(page?.offsetPt?.downPt).toBeCloseTo(40, 6);
  });

  // The one place the other direction is not a question of noise: Word drew no text on
  // this page at all, so there is nothing to be noisy about and whatever we put there
  // Word put nowhere.
  it("is called missing where Word drew no text on it at all", () => {
    const laid = laidOut(["alpha", "bravo"]);
    const page = agreementWith(laid, [], TOLERANCE_PT).pages[0];

    expect(page?.shape).toBe("missing");
    expect(page?.oursAlone).toBe(2);
  });
});
