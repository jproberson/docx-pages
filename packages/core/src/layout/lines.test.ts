import { describe, expect, it } from "vitest";

import type { ParagraphMark } from "../docx/styles.js";
import type { RunPiece, TextRun } from "../docx/runs.js";
import { buildSfnt } from "../testing/build-font.js";
import { readFontFile } from "./font-file.js";
import {
  NO_ADVANCES,
  type FaceElsewhere,
  type MetricsLookup,
  type SuppliedFace,
} from "./font-metrics.js";
import { beginLines, breakLines, justifyLine, type TextLine } from "./lines.js";
import type { TabStopPt } from "./tab-stops.js";

// Every glyph is half an em wide, so a 10pt run measures exactly 5pt a character
// and the expected break points can be counted rather than computed.
const HALF_EM = 500;
const NO_BREAK_SPACE = "\u00a0";
// The other three Word will not break at, so that each measures half an em here
// like everything else and a width can be counted rather than computed.
const NO_BREAK_SPACES = `${NO_BREAK_SPACE}\u2007\u202f\ufeff`;
const CHARACTERS = `abcdefghijklmnopqrstuvwxyz0123456789. -${NO_BREAK_SPACES}`;

const FIXTURE = {
  unitsPerEm: 1000,
  ascender: 800,
  descender: -200,
  lineGap: 0,
  advances: Object.fromEntries(Array.from(CHARACTERS, (character) => [character, HALF_EM])),
};

const METRICS = {
  unitsPerEm: FIXTURE.unitsPerEm,
  ascender: FIXTURE.ascender,
  descender: FIXTURE.descender,
  lineGap: FIXTURE.lineGap,
};

const EVEN: SuppliedFace = {
  name: "Even Sans",
  bold: false,
  italic: false,
  metrics: METRICS,
  advances: readFontFile(buildSfnt(FIXTURE)).advances,
};

const metricsFor =
  (faces: readonly SuppliedFace[] = [EVEN]) =>
  (request: { readonly name: string }): MetricsLookup => {
    const face = faces.find((each) => each.name === request.name);
    return face === undefined
      ? { kind: "missing", fontName: request.name }
      : { kind: "found", source: "supplied", metrics: face.metrics, advances: face.advances };
  };

const mark = (fontSizePt = 10, name = "Even Sans"): ParagraphMark => ({
  font: { kind: "named", name },
  fontSizePt,
  bold: false,
  italic: false,
  underline: false,
  raisePt: 0,
  lineSizePt: fontSizePt,
  lineRaisePt: 0,
  color: null,
  characterSpacingPt: 0,
  characterScale: 1,
});

const runOf = (text: string, at: ParagraphMark = mark()): TextRun => ({
  mark: at,
  pieces: [{ kind: "text", text }],
});

const piecesRun = (pieces: readonly RunPiece[], at: ParagraphMark = mark()): TextRun => ({
  mark: at,
  pieces,
});

function linesOf(
  runs: readonly TextRun[],
  widthPt: number,
  faces?: readonly SuppliedFace[],
): readonly TextLine[] {
  const result = breakLines({ runs, widthPt, metricsFor: metricsFor(faces) });
  if (result.kind !== "lines") throw new Error(result.failure.kind);
  return result.lines;
}

const leftStops = (...positionsPt: readonly number[]): readonly TabStopPt[] =>
  positionsPt.map((positionPt) => ({ positionPt, alignment: "left" as const }));

// Where the text after a tab starts, which is the whole question a stop that lines
// its text up answers.
function afterTabPt(pieces: readonly RunPiece[], stop: TabStopPt): number {
  const result = breakLines({
    runs: [piecesRun(pieces)],
    widthPt: 500,
    metricsFor: metricsFor(),
    tabs: { stopsPt: [stop], originPt: 0 },
  });
  if (result.kind !== "lines") throw new Error(result.failure.kind);
  const line = result.lines[0] ?? never();
  return line.segments[1]?.offsetPt ?? Number.NaN;
}

const textOf = (line: TextLine): string =>
  line.segments.map((segment) => (segment.kind === "text" ? segment.text : "")).join("");

const sizesOf = (line: TextLine): readonly number[] =>
  line.segments.flatMap((segment) => (segment.kind === "text" ? [segment.mark.fontSizePt] : []));

describe("breakLines", () => {
  // A paragraph can ask for indents wider than the frame it stands in, which is
  // what a real document did inside a narrow cell: 2457 twips either side of a cell
  // a fraction of that wide. The width left over is then negative, and the line
  // breaker took nothing, moved nothing on and asked again for ever. It laid out no
  // page and never returned, so the whole document hung.
  it("breaks a paragraph given less room than none rather than never finishing", () => {
    for (const roomPt of [-1, -50, -245]) {
      const lines = linesOf([runOf("abc def")], roomPt);
      expect(lines.length).toBeGreaterThan(0);
      expect(lines.map(textOf).join("")).toBe("abcdef");
    }
  });

  // Room below nothing is the same as none, so the two answer alike.
  it("gives a line with no room at all what a line with less than none gets", () => {
    expect(linesOf([runOf("abc def")], -20).map(textOf)).toStrictEqual(
      linesOf([runOf("abc def")], 0).map(textOf),
    );
  });

  it("keeps a line that fits as one line", () => {
    const lines = linesOf([runOf("abc def")], 100);

    expect(lines).toHaveLength(1);
    expect(textOf(lines[0] ?? never())).toBe("abc def");
    expect(lines[0]?.widthPt).toBeCloseTo(35, 9);
  });

  it("breaks at the last space that fits", () => {
    // "abc def" is 35pt, "abc" alone is 15pt, so a 20pt line takes only the first word.
    const lines = linesOf([runOf("abc def")], 20);

    expect(lines.map(textOf)).toStrictEqual(["abc", "def"]);
  });

  it("breaks after a hyphen inside a word, leaving the hyphen on the line it ends", () => {
    // "ab-cd" is 25pt and "ab-" alone is 15pt, so a 20pt line ends on the hyphen.
    expect(linesOf([runOf("ab-cd")], 20).map(textOf)).toStrictEqual(["ab-", "cd"]);
  });

  it("keeps a hyphenated word whole when the line has room for it", () => {
    const lines = linesOf([runOf("ab-cd ef")], 100);

    expect(lines.map(textOf)).toStrictEqual(["ab-cd ef"]);
    expect(lines[0]?.widthPt).toBeCloseTo(40, 9);
  });

  it("drops the space a wrap happens at rather than starting a line with it", () => {
    const lines = linesOf([runOf("abc def")], 20);
    expect(lines[1]?.widthPt).toBeCloseTo(15, 9);
  });

  it("does not count trailing spaces towards the width, so they never force a wrap", () => {
    const lines = linesOf([piecesRun([{ kind: "text", text: "abcd    " }])], 20);

    expect(lines).toHaveLength(1);
    expect(lines[0]?.widthPt).toBeCloseTo(20, 9);
  });

  it("carries a word joined by a no-break space to the next line whole", () => {
    // "abc ab" fills a 30pt line exactly, so a line breaking at the no-break space
    // would end there and start the next one at "cd".
    const lines = linesOf([runOf(`abc ab${NO_BREAK_SPACE}cd`)], 30);

    expect(lines.map(textOf)).toStrictEqual(["abc", `ab${NO_BREAK_SPACE}cd`]);
  });

  it("counts a no-break space towards the width, being a character like any other", () => {
    const lines = linesOf([runOf(`ab${NO_BREAK_SPACE}cd`)], 100);

    expect(lines[0]?.widthPt).toBeCloseTo(25, 9);
  });

  // Word's own answers, measured on 2026-08-13 by a document putting each of these
  // between the second and third of three words in a column holding two of them.
  // The narrow no-break space is the one the corpus turns on: 20 documents and 1169
  // characters, against none at all for the other two.
  for (const [name, character] of [
    ["narrow no-break space", "\u202f"],
    ["figure space", "\u2007"],
    ["zero width no-break space", "\ufeff"],
  ] as const) {
    it(`carries a word joined by a ${name} to the next line whole`, () => {
      const lines = linesOf([runOf(`abc ab${character}cd`)], 30);

      expect(lines.map(textOf)).toStrictEqual(["abc", `ab${character}cd`]);
    });
  }

  // The other half of the same rule, and the half a real document turned on: an
  // ordinary space at the end of a line hangs past the margin and costs nothing,
  // where a no-break space belongs to the word in front of it and has to fit.
  it("holds a trailing narrow no-break space to the width the word must fit", () => {
    const lines = linesOf([runOf(`abc abc\u202f\u202f`)], 30);

    expect(lines.map(textOf)).toStrictEqual(["abc", "abc\u202f\u202f"]);
  });

  it("breaks a word that cannot fit a line of its own at the character that overflows", () => {
    const lines = linesOf([runOf("abcdefgh")], 20);

    expect(lines.map(textOf)).toStrictEqual(["abcd", "efgh"]);
  });

  it("always places at least one character, however narrow the line", () => {
    expect(linesOf([runOf("ab")], 1).map(textOf)).toStrictEqual(["a", "b"]);
  });

  // Word cuts a word too long for its line in the same place whether it was written
  // as one run or as several: measured on 2026-08-08 by the authored
  // `insignificant-space` document, where twenty four characters written as two runs
  // of twelve were cut at the fifteenth exactly as the same characters written as
  // one. This project used to send a whole run to the next line as soon as it did
  // not fit, which put the cut at the boundary between two runs.
  it("cuts a word written in more than one run where the character overflows", () => {
    const lines = linesOf([runOf("abcd"), runOf("efgh")], 20);

    expect(lines.map(textOf)).toStrictEqual(["abcd", "efgh"]);
  });

  it("cuts inside a later run rather than at the boundary in front of it", () => {
    const lines = linesOf([runOf("abcd"), runOf("efgh")], 30);

    expect(lines.map(textOf)).toStrictEqual(["abcdef", "gh"]);
  });

  // Which of a paragraph's lines a page break put at the head of a page, and
  // whether the paragraph ran out on one. The lines alone cannot say: a break the
  // paragraph ends on draws nothing, so what it asks of the page is on the flow.
  function pagesOf(pieces: readonly RunPiece[]): {
    readonly texts: readonly string[];
    readonly starts: readonly boolean[];
    readonly endsPage: boolean;
  } {
    const started = beginLines({ runs: [piecesRun(pieces)], metricsFor: metricsFor() });
    if (started.kind !== "flow") throw new Error(started.failure.kind);

    const texts: string[] = [];
    const starts: boolean[] = [];
    let flow = started.flow;
    for (;;) {
      const taken = flow.next(100);
      if (taken === null) return { texts, starts, endsPage: flow.startsPage };
      texts.push(textOf(taken.line));
      starts.push(flow.startsPage);
      flow = taken.rest;
    }
  }

  const PAGE: RunPiece = { kind: "break", endsPage: true, endsColumn: false };
  const NEW_LINE: RunPiece = { kind: "break", endsPage: false, endsColumn: false };

  it("carries the text after a page break onto a page of its own", () => {
    expect(
      pagesOf([{ kind: "text", text: "ab" }, PAGE, { kind: "text", text: "cd" }]),
    ).toStrictEqual({ texts: ["ab", "cd"], starts: [false, true], endsPage: false });
  });

  it("keeps the line a page break ends even where nothing stood on it", () => {
    expect(pagesOf([PAGE])).toStrictEqual({ texts: [""], starts: [false], endsPage: true });
  });

  // **Every break opens a line under it, and that line stands whether anything is
  // written on it or not.** Measured on 2026-08-08 by the authored
  // `breaks-in-a-paragraph` document: over eight cases written out three times, a
  // paragraph came out one line taller for every break in it, wherever the breaks
  // stood, and one holding a break and nothing else came out two lines tall.
  //
  // Two corpus documents of one converted template turn on it, at 7 of 45 lines
  // placed and 6 of 43: one writes two breaks in a row in the middle of a paragraph,
  // the other ends a paragraph with one, and everything below either was a line too
  // high.
  it("opens a line under a line break with nothing on it", () => {
    expect(pagesOf([NEW_LINE])).toStrictEqual({
      texts: ["", ""],
      starts: [false, false],
      endsPage: false,
    });
  });

  it("opens a line under a break with nothing after it", () => {
    expect(pagesOf([{ kind: "text", text: "ab" }, NEW_LINE])).toStrictEqual({
      texts: ["ab", ""],
      starts: [false, false],
      endsPage: false,
    });
  });

  it("leaves a line with nothing on it between two breaks together", () => {
    expect(
      pagesOf([{ kind: "text", text: "ab" }, NEW_LINE, NEW_LINE, { kind: "text", text: "cd" }]),
    ).toStrictEqual({ texts: ["ab", "", "cd"], starts: [false, false, false], endsPage: false });
  });

  // Nothing follows the break to be put on the next page, so the paragraph itself
  // is what has to carry the ask.
  it("says the paragraph ended on a page break when nothing came after it", () => {
    expect(pagesOf([{ kind: "text", text: "ab" }, PAGE])).toStrictEqual({
      texts: ["ab"],
      starts: [false],
      endsPage: true,
    });
  });

  it("leaves a page with nothing on it between two breaks together", () => {
    expect(pagesOf([PAGE, PAGE])).toStrictEqual({
      texts: ["", ""],
      starts: [false, true],
      endsPage: true,
    });
  });

  it("ends a line where the run asks for a break", () => {
    const run = piecesRun([
      { kind: "text", text: "ab" },
      { kind: "break", endsPage: false, endsColumn: false },
      { kind: "text", text: "cd" },
    ]);

    expect(linesOf([run], 100).map(textOf)).toStrictEqual(["ab", "cd"]);
  });

  it("carries each run's own mark onto the line it lands on", () => {
    const big = mark(20);
    const lines = linesOf([runOf("ab"), runOf("cd", big)], 100);

    expect(sizesOf(lines[0] ?? never())).toStrictEqual([10, 20]);
  });

  it("makes a line as tall as its tallest run", () => {
    const lines = linesOf([runOf("ab"), runOf("cd", mark(20))], 100);

    expect(lines[0]?.heightPt).toBeCloseTo(20, 9);
    expect(lines[0]?.ascentPt).toBeCloseTo(16, 9);
  });

  // Measured on 2026-08-07 by the authored `raised-text` document: a 12pt run
  // raised six points beside a plain one took its line from 14.64pt to 20.64, all
  // of it above the baseline, and a run lowered six took it to the same height
  // below. The face here is 8 above the baseline and 2 below at 10pt.
  it("carries a raised run's whole line with it", () => {
    const raised = { ...mark(), lineRaisePt: 3, raisePt: 3 };
    const line = linesOf([runOf("ab"), runOf("cd", raised)], 100)[0] ?? never();

    expect(line.ascentPt).toBeCloseTo(11, 9);
    expect(line.heightPt).toBeCloseTo(13, 9);
  });

  it("carries a lowered run's line down instead", () => {
    const lowered = { ...mark(), lineRaisePt: -3, raisePt: -3 };
    const line = linesOf([runOf("ab"), runOf("cd", lowered)], 100)[0] ?? never();

    expect(line.ascentPt).toBeCloseTo(8, 9);
    expect(line.heightPt).toBeCloseTo(13, 9);
  });

  // What a raised run no longer reaches below the baseline counts for nothing
  // rather than pulling the next line up: raised six points and alone on its line,
  // a 12pt run left the line 17.52pt tall against the 14.64 it makes flat.
  it("lets a raise past the descent leave the line no shallower", () => {
    const raised = { ...mark(), lineRaisePt: 5, raisePt: 5 };
    const line = linesOf([runOf("ab", raised)], 100)[0] ?? never();

    expect(line.ascentPt).toBeCloseTo(13, 9);
    expect(line.heightPt).toBeCloseTo(13, 9);
  });

  it("does the same at the top for a run lowered past the ascent", () => {
    const lowered = { ...mark(), lineRaisePt: -12, raisePt: -12 };
    const line = linesOf([runOf("ab", lowered)], 100)[0] ?? never();

    expect(line.ascentPt).toBeCloseTo(0, 9);
    expect(line.heightPt).toBeCloseTo(14, 9);
  });

  // A multiple line rule is taken of the line the faces make and the raise is added
  // to it, so the two are kept apart: 12pt text raised six under a line and a half
  // came out 27.96pt, which is 20.64 and half of 14.64 rather than half again of
  // 20.64.
  it("leaves the line its faces make where a raise grew the line itself", () => {
    const raised = { ...mark(), lineRaisePt: 3, raisePt: 3 };
    const line = linesOf([runOf("ab"), runOf("cd", raised)], 100)[0] ?? never();

    expect(line.fontHeightPt).toBeCloseTo(10, 9);
  });

  it("gives a hanging first line the extra room it starts with", () => {
    const result = breakLines({
      runs: [runOf("aaa bbb")],
      widthPt: 20,
      firstLineWidthPt: 40,
      metricsFor: metricsFor(),
    });
    if (result.kind !== "lines") throw new Error(result.failure.kind);

    expect(result.lines.map(textOf)).toStrictEqual(["aaa bbb"]);
  });

  // The stops belong to the text area, but a hanging first line begins outside it,
  // so a tab there reaches the stop the lines below it start from.
  it("measures a tab on a hanging first line from where that line starts", () => {
    const result = breakLines({
      runs: [piecesRun([{ kind: "tab" }, { kind: "text", text: "ab" }])],
      widthPt: 200,
      metricsFor: metricsFor(),
      tabs: { stopsPt: leftStops(10, 60), originPt: 0, firstLineOriginPt: -8 },
    });
    if (result.kind !== "lines") throw new Error(result.failure.kind);

    expect(result.lines[0]?.widthPt).toBeCloseTo(18 + 10, 9);
  });

  it("advances a tab to the next default stop", () => {
    const run = piecesRun([
      { kind: "text", text: "ab" },
      { kind: "tab" },
      { kind: "text", text: "cd" },
    ]);
    const lines = linesOf([run], 200);

    expect(lines[0]?.widthPt).toBeCloseTo(36 + 10, 9);
  });

  // Measured against Word: against a stop at 144pt with 39.5pt of text after it,
  // Word starts that text at 124 for a centre stop, 104 for a right one and 120
  // for a decimal one, which puts the point itself on 144.
  it("centres the text after a tab on a stop that asks for it", () => {
    const pieces: readonly RunPiece[] = [
      { kind: "text", text: "ab" },
      { kind: "tab" },
      { kind: "text", text: "cdef" },
    ];

    expect(afterTabPt(pieces, { positionPt: 100, alignment: "center" })).toBeCloseTo(90, 9);
  });

  it("ends the text after a tab on a right stop", () => {
    const pieces: readonly RunPiece[] = [
      { kind: "text", text: "ab" },
      { kind: "tab" },
      { kind: "text", text: "cdef" },
    ];

    expect(afterTabPt(pieces, { positionPt: 100, alignment: "right" })).toBeCloseTo(80, 9);
  });

  it("puts the first decimal point of the text on a decimal stop", () => {
    const pieces: readonly RunPiece[] = [
      { kind: "text", text: "ab" },
      { kind: "tab" },
      { kind: "text", text: "12.3.4" },
    ];

    expect(afterTabPt(pieces, { positionPt: 100, alignment: "decimal" })).toBeCloseTo(90, 9);
  });

  it("ends text with no decimal point in it on a decimal stop, as a right one would", () => {
    const pieces: readonly RunPiece[] = [
      { kind: "text", text: "ab" },
      { kind: "tab" },
      { kind: "text", text: "cdef" },
    ];

    expect(afterTabPt(pieces, { positionPt: 100, alignment: "decimal" })).toBeCloseTo(80, 9);
  });

  it("reaches only as far as the tab after it, and leaves the space it ends on out", () => {
    const pieces: readonly RunPiece[] = [
      { kind: "text", text: "ab" },
      { kind: "tab" },
      { kind: "text", text: "cd " },
      { kind: "tab" },
      { kind: "text", text: "ef" },
    ];

    expect(afterTabPt(pieces, { positionPt: 100, alignment: "right" })).toBeCloseTo(90, 9);
  });

  it("opens no room at all where the text it lines up will not fit in front of the stop", () => {
    const pieces: readonly RunPiece[] = [
      { kind: "text", text: "ab" },
      { kind: "tab" },
      { kind: "text", text: "cdefghijkl" },
    ];

    expect(afterTabPt(pieces, { positionPt: 20, alignment: "right" })).toBeCloseTo(10, 9);
  });

  it("holds a line open as far as a trailing tab reached, though it draws nothing", () => {
    const lines = linesOf([piecesRun([{ kind: "tab" }])], 200);

    expect(lines).toHaveLength(1);
    expect(lines[0]?.widthPt).toBeCloseTo(36, 9);
    expect(lines[0]?.segments).toStrictEqual([]);
  });

  it("counts a tab after the text as part of the line's width", () => {
    const run = piecesRun([{ kind: "text", text: "ab" }, { kind: "tab" }]);

    expect(linesOf([run], 200)[0]?.widthPt).toBeCloseTo(36, 9);
  });

  it("says where along the line each run starts", () => {
    const lines = linesOf([runOf("ab"), runOf("cd", mark(20))], 200);
    const offsets = (lines[0] ?? never()).segments.map((segment) => segment.offsetPt);

    expect(offsets).toStrictEqual([0, 10]);
  });

  // Nothing else on the line can find the gap a tab opened by adding up widths,
  // which is why every run carries its own place.
  it("starts the run after a tab at the stop the tab reached", () => {
    const run = piecesRun([
      { kind: "text", text: "ab" },
      { kind: "tab" },
      { kind: "text", text: "cd" },
    ]);
    const segments = (linesOf([run], 200)[0] ?? never()).segments;

    expect(segments.map((segment) => segment.offsetPt)).toStrictEqual([0, 36]);
  });

  it("gives a drawing the width it takes on the line", () => {
    const run = piecesRun([
      { kind: "drawing", widthEmu: 914400, heightEmu: 457200, turnDegrees: 0 },
    ]);
    const lines = linesOf([run], 200);

    expect(lines[0]?.widthPt).toBeCloseTo(72, 9);
    expect(lines[0]?.heightPt).toBeCloseTo(36, 9);
  });

  // Word rounds the turn to the nearest quarter and keeps the extent that way
  // round, so a picture turned a quarter is held in a box as wide as it was tall.
  // Measured off Word's own pdf of the authored `rotated-drawings` document.
  it("gives a drawing turned a quarter the room its turn lays it in", () => {
    const run = piecesRun([
      { kind: "drawing", widthEmu: 914400, heightEmu: 457200, turnDegrees: 90 },
    ]);
    const lines = linesOf([run], 200);

    expect(lines[0]?.widthPt).toBeCloseTo(36, 9);
    expect(lines[0]?.heightPt).toBeCloseTo(72, 9);
  });

  // The same picture turned by 30 degrees is drawn 74.35 x 56.78 and is held in
  // the box it was stored in regardless, which is what Word does.
  it("gives a drawing turned less than an eighth the room it was stored at", () => {
    const run = piecesRun([
      { kind: "drawing", widthEmu: 914400, heightEmu: 457200, turnDegrees: 30 },
    ]);
    const lines = linesOf([run], 200);

    expect(lines[0]?.widthPt).toBeCloseTo(72, 9);
    expect(lines[0]?.heightPt).toBeCloseTo(36, 9);
  });

  it("produces no lines for runs that carry no text", () => {
    expect(linesOf([], 100)).toStrictEqual([]);
  });

  it("reports a face whose advances were never supplied rather than guessing widths", () => {
    const unmeasurable = [{ ...EVEN, advances: NO_ADVANCES }];
    const result = breakLines({
      runs: [runOf("abc")],
      widthPt: 100,
      metricsFor: metricsFor(unmeasurable),
    });

    expect(result).toStrictEqual({
      kind: "unmeasurable",
      failure: { kind: "unmeasurable-text", fontName: "Even Sans", reason: "unsupplied" },
    });
  });

  it("reports a character the face does not map", () => {
    const result = breakLines({
      runs: [runOf("aé")],
      widthPt: 100,
      metricsFor: metricsFor(),
    });

    expect(result).toStrictEqual({
      kind: "unmeasurable",
      failure: { kind: "unmapped-character", fontName: "Even Sans", codePoint: 0x00e9 },
    });
  });

  it("reports a font the caller could not resolve", () => {
    const result = breakLines({
      runs: [runOf("abc", mark(10, "Absent Sans"))],
      widthPt: 100,
      metricsFor: metricsFor(),
    });

    expect(result).toStrictEqual({
      kind: "unmeasurable",
      failure: { kind: "unknown-font-metrics", fontName: "Absent Sans" },
    });
  });

  it("reports a run whose font the style cascade never named", () => {
    const unnamed: ParagraphMark = {
      font: { kind: "unresolved" },
      fontSizePt: 10,
      bold: false,
      italic: false,
      underline: false,
      raisePt: 0,
      lineSizePt: 10,
      lineRaisePt: 0,
      color: null,
      characterSpacingPt: 0,
      characterScale: 1,
    };
    const result = breakLines({
      runs: [runOf("abc", unnamed)],
      widthPt: 100,
      metricsFor: metricsFor(),
    });

    expect(result).toStrictEqual({ kind: "unmeasurable", failure: { kind: "unresolved-font" } });
  });
});

// Word measures a line over the face it borrowed a character from as well as over
// the one the run states. Measured on 2026-08-06 off Word's pdf of the authored
// `unmapped-characters` document: the line holding the borrowed bullet came out
// 13.80pt tall at 12pt against the 13.32pt of the symbol face beside it.
// **How wide a run the file scaled is measured**, which decides where its line
// breaks and therefore where every page under it falls. Measured against Word on
// 2026-08-14: the scale multiplies the glyph's own advance, and the letter spacing
// this project already reads is added after it rather than scaled with it. A run
// scaled to 150 with a point of spacing came out 253.63pt against 225.66 scaled
// alone and 150.54 plain, where spacing first and scaling after would have given
// 267.8.
describe("a run the file scaled", () => {
  const scaled = (scale: number, spacingPt = 0): number => {
    const at: ParagraphMark = { ...mark(), characterScale: scale, characterSpacingPt: spacingPt };
    const broken = breakLines({
      runs: [runOf("aaaa", at)],
      widthPt: 1000,
      metricsFor: metricsFor(),
    });
    if (broken.kind !== "lines") throw new Error(broken.kind);
    return broken.lines[0]?.widthPt ?? 0;
  };

  it("multiplies every advance by the scale the run states", () => {
    const plain = scaled(1);
    expect(scaled(1.5)).toBeCloseTo(plain * 1.5, 9);
    expect(scaled(0.5)).toBeCloseTo(plain * 0.5, 9);
    expect(scaled(1.03)).toBeCloseTo(plain * 1.03, 9);
  });

  it("adds the letter spacing after the scale rather than scaling it too", () => {
    const plain = scaled(1);
    // Four characters, each carrying a point of spacing after it.
    expect(scaled(1.5, 1)).toBeCloseTo(plain * 1.5 + 4, 9);
  });
});

describe("a character drawn out of another face", () => {
  const BULLET = "\u2022";

  // Taller and deeper than `EVEN`, so what it lends the line shows in both.
  const TALL = readFontFile(
    buildSfnt({
      unitsPerEm: 1000,
      ascender: 900,
      descender: -300,
      lineGap: 0,
      advances: { [BULLET]: 700 },
    }),
  );
  const advances = TALL.advances;
  const lent: FaceElsewhere | null =
    advances.kind === "advances"
      ? (codePoint) => {
          const advance = advances.advanceFor(codePoint);
          return advance === null ? null : { metrics: TALL.metrics, advance };
        }
      : null;

  const lending = (request: { readonly name: string }): MetricsLookup => {
    const found = metricsFor()(request);
    return found.kind === "found" && lent !== null ? { ...found, elsewhere: lent } : found;
  };

  const lineOf = (text: string, metrics = lending): TextLine => {
    const result = breakLines({ runs: [runOf(text)], widthPt: 1000, metricsFor: metrics });
    if (result.kind !== "lines") throw new Error(result.failure.kind);
    const [line] = result.lines;
    if (line === undefined) throw new Error("no line");
    return line;
  };

  it("takes the width out of the face that drew it", () => {
    expect(lineOf(`ab${BULLET}`).widthPt).toBeCloseTo(5 + 5 + 7, 9);
  });

  it("stands as tall as that face, and seats the line under its ascent", () => {
    const line = lineOf(`ab${BULLET}`);

    expect(line.ascentPt).toBeCloseTo(9, 9);
    expect(line.heightPt).toBeCloseTo(12, 9);
  });

  it("leaves a line of characters the face itself draws alone", () => {
    const line = lineOf("ab");

    expect(line.ascentPt).toBeCloseTo(8, 9);
    expect(line.heightPt).toBeCloseTo(10, 9);
  });
});

// **What a justified line may be squeezed by, measured on 2026-08-10 off Word's own
// pdf**: at most a quarter of the spaces it holds, and at most a third of the
// advance of the word it is being asked to take, counting the space in front of it,
// plus 0.2307 of that advance over the line's spaces less a half.
//
// Every glyph of the face here is half an em, so a 10pt run measures 5pt a character
// and each allowance can be counted out by hand.
describe("breakLines squeezing a justified line", () => {
  const justifiedWords = (text: string, widthPt: number): number => {
    const result = breakLines({
      runs: [runOf(text)],
      widthPt,
      metricsFor: metricsFor(),
      justified: true,
    });
    if (result.kind !== "lines") throw new Error(result.failure.kind);
    return textOf(result.lines[0] ?? never())
      .trim()
      .split(/\s+/).length;
  };

  // "aa aa aa aa b" is 65pt over four spaces. The word it ends on advances 10pt with
  // its space, so the line may be squeezed by 10 * (1/3 + 0.2307/3.5) = 3.992pt,
  // which is under the quarter of 20pt of spaces the same line offers.
  it("takes a word overflowing by a third of its advance and a share over the spaces", () => {
    expect(justifiedWords("aa aa aa aa b", 61.1)).toBe(5);
    expect(justifiedWords("aa aa aa aa b", 61)).toBe(4);
  });

  // The same word on a line of eight spaces: 10 * (1/3 + 0.2307/7.5) = 3.641pt, so
  // the more spaces a line holds the less it may be squeezed to take one more word.
  it("allows less the more spaces the line holds", () => {
    expect(justifiedWords("aa aa aa aa aa aa aa aa b", 121.4)).toBe(9);
    expect(justifiedWords("aa aa aa aa aa aa aa aa b", 121.3)).toBe(8);
  });

  // "abcd abcd abcd abcd abcd" is 120pt over four spaces. A third of the last word's
  // 25pt advance is more than the spaces can give up, so the quarter is what holds.
  it("never squeezes a space by more than a quarter of itself", () => {
    expect(justifiedWords("abcd abcd abcd abcd abcd", 115.1)).toBe(5);
    expect(justifiedWords("abcd abcd abcd abcd abcd", 114.9)).toBe(4);
  });

  it("squeezes nothing at all where the paragraph is not justified", () => {
    const lines = linesOf([runOf("aa aa aa aa b")], 61.1);

    expect(textOf(lines[0] ?? never()).trim()).toBe("aa aa aa aa");
  });
});

// Measured against Word itself: every space character on the line takes an equal
// share of the room the line did not fill, and nothing else takes any.
describe("justifyLine", () => {
  const offsetsOf = (line: TextLine): readonly number[] =>
    line.segments.map((segment) => segment.offsetPt);

  const only = (runs: readonly TextRun[], widthPt = 500): TextLine =>
    linesOf(runs, widthPt)[0] ?? never();

  it("hands each space an equal share of what the line did not fill", () => {
    // "ab cd ef" is 40pt over two spaces, so a 60pt line gives each one 10pt more.
    const line = justifyLine(only([runOf("ab cd ef")]), 60);

    expect(line.widthPt).toBe(60);
    expect(offsetsOf(line)).toStrictEqual([0, 10, 25, 35, 50]);
  });

  it("counts each space character, so a double space grows twice as far", () => {
    // "ab  cd ef" is 45pt: three space characters, 15pt of slack, 5pt each. The
    // double space is one segment, and takes two shares of it.
    const line = justifyLine(only([runOf("ab  cd ef")]), 60);

    expect(offsetsOf(line)).toStrictEqual([0, 10, 30, 40, 50]);
  });

  it("gives a space the same share whatever size it is set in", () => {
    // The 20pt run's space is twice as wide as the 10pt one, and grows as much.
    const line = justifyLine(only([runOf("ab "), runOf("cd ", mark(20)), runOf("ef")]), 90);
    const widths = line.segments.map((segment) => (segment.kind === "text" ? segment.widthPt : 0));

    expect(widths).toStrictEqual([10, 22.5, 20, 27.5, 10]);
  });

  it("leaves a no-break space out of the sharing, as part of the word around it", () => {
    const line = justifyLine(only([runOf(`ab${NO_BREAK_SPACE}cd ef`)]), 60);

    expect(offsetsOf(line)).toStrictEqual([0, 25, 50]);
  });

  it("leaves a line that already fills its room alone", () => {
    const line = only([runOf("ab cd")]);

    expect(justifyLine(line, line.widthPt)).toBe(line);
  });

  it("leaves a line with nowhere to grow alone", () => {
    const line = only([runOf("abcd")]);

    expect(justifyLine(line, 100)).toBe(line);
  });
});

function never(): never {
  throw new Error("expected a line");
}
