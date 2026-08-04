import { describe, expect, it } from "vitest";

import type { ParagraphMark } from "../docx/styles.js";
import type { RunPiece, TextRun } from "../docx/runs.js";
import { buildSfnt } from "../testing/build-font.js";
import { readFontFile } from "./font-file.js";
import { NO_ADVANCES, type MetricsLookup, type SuppliedFace } from "./font-metrics.js";
import { breakLines, type TextLine } from "./lines.js";

// Every glyph is half an em wide, so a 10pt run measures exactly 5pt a character
// and the expected break points can be counted rather than computed.
const HALF_EM = 500;
const CHARACTERS = "abcdefghijklmnopqrstuvwxyz -";

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
  color: null,
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

const textOf = (line: TextLine): string =>
  line.segments.map((segment) => (segment.kind === "text" ? segment.text : "")).join("");

const sizesOf = (line: TextLine): readonly number[] =>
  line.segments.flatMap((segment) => (segment.kind === "text" ? [segment.mark.fontSizePt] : []));

describe("breakLines", () => {
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

  it("breaks a word that cannot fit a line of its own at the character that overflows", () => {
    const lines = linesOf([runOf("abcdefgh")], 20);

    expect(lines.map(textOf)).toStrictEqual(["abcd", "efgh"]);
  });

  it("always places at least one character, however narrow the line", () => {
    expect(linesOf([runOf("ab")], 1).map(textOf)).toStrictEqual(["a", "b"]);
  });

  it("ends a line where the run asks for a break", () => {
    const run = piecesRun([
      { kind: "text", text: "ab" },
      { kind: "break" },
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

  it("advances a tab to the next default stop", () => {
    const run = piecesRun([
      { kind: "text", text: "ab" },
      { kind: "tab" },
      { kind: "text", text: "cd" },
    ]);
    const lines = linesOf([run], 200);

    expect(lines[0]?.widthPt).toBeCloseTo(36 + 10, 9);
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
    const run = piecesRun([{ kind: "drawing", widthEmu: 914400, heightEmu: 457200 }]);
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
      color: null,
    };
    const result = breakLines({
      runs: [runOf("abc", unnamed)],
      widthPt: 100,
      metricsFor: metricsFor(),
    });

    expect(result).toStrictEqual({ kind: "unmeasurable", failure: { kind: "unresolved-font" } });
  });
});

function never(): never {
  throw new Error("expected a line");
}
