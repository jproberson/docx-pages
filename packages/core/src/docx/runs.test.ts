import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { readParagraphs } from "./blocks.js";
import { openDocx } from "./package.js";
import { readRuns, type TextRun } from "./runs.js";
import { readStyleTable } from "./styles.js";

const NORMAL = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr></w:style></w:styles>`;

function runsOf(
  body: string,
  stylesXml: string = NORMAL,
  documentXml: string = wordDocument(body),
): readonly TextRun[] {
  const pkg = openDocx(
    buildDocx({ "word/document.xml": documentXml, "word/styles.xml": stylesXml }),
  );
  const paragraph = readParagraphs(pkg)[0];
  if (paragraph === undefined) throw new Error("expected a paragraph");
  return readRuns(paragraph, readStyleTable(pkg));
}

const NO_FACE = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:rPr><w:sz w:val="24"/></w:rPr></w:style></w:styles>`;

const textOf = (runs: readonly TextRun[]): string =>
  runs
    .flatMap((run) => run.pieces)
    .map((piece) => (piece.kind === "text" ? piece.text : ""))
    .join("");

describe("readRuns", () => {
  it("reads a paragraph's runs in the order they are written", () => {
    const first = `<w:r><w:t xml:space="preserve">one </w:t></w:r>`;
    const runs = runsOf(`<w:p>${first}<w:r><w:t>two</w:t></w:r></w:p>`);

    expect(runs).toHaveLength(2);
    expect(textOf(runs)).toBe("one two");
  });

  it("keeps whitespace a run asks to preserve", () => {
    const runs = runsOf(`<w:p><w:r><w:t xml:space="preserve"> a </w:t></w:r></w:p>`);
    expect(textOf(runs)).toBe(" a ");
  });

  it("drops edge whitespace a run does not ask to preserve", () => {
    expect(textOf(runsOf(`<w:p><w:r><w:t>  a  </w:t></w:r></w:p>`))).toBe("a");
  });

  // A no-break space is significant wherever it sits, and Word leaves it to a run
  // that never asks for it: the runs either side of an emphasised phrase carry the
  // spaces around it that way.
  it("keeps a no-break space at a run's edge, whatever the run asks for", () => {
    const body = `<w:p><w:r><w:t>  \u00a0a\u00a0  </w:t></w:r></w:p>`;

    expect(textOf(runsOf(body))).toBe("\u00a0a\u00a0");
  });

  // **`xml:space` is inherited, and Word reads it from wherever it is stated.** The
  // worst-placed document in the corpus states `preserve` on `w:document` itself and
  // on no `w:t` anywhere, and its bare spaces survive: Word ends a heading's line at
  // one of them, and draws that line whole again when the attribute is taken off the
  // root and nothing else is touched.
  const statingSpace = (body: string, space: string): string =>
    wordDocument(body).replace("<w:document ", `<w:document xml:space="${space}" `);

  it("keeps whitespace a document preserves for the whole of itself", () => {
    const body = `<w:p><w:r><w:t> a </w:t></w:r></w:p>`;

    expect(textOf(runsOf(body, NORMAL, statingSpace(body, "preserve")))).toBe(" a ");
  });

  it("drops whitespace where the nearer statement is the default", () => {
    const body = `<w:p><w:r><w:t xml:space="default"> a </w:t></w:r></w:p>`;

    expect(textOf(runsOf(body, NORMAL, statingSpace(body, "preserve")))).toBe("a");
  });

  it("joins several text pieces inside one run", () => {
    const body = `<w:p><w:r><w:t xml:space="preserve">a </w:t><w:t>b</w:t></w:r></w:p>`;
    const runs = runsOf(body);

    expect(runs).toHaveLength(1);
    expect(textOf(runs)).toBe("a b");
  });

  it("reads a tab as its own piece, since its width is not a glyph's", () => {
    const runs = runsOf(`<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>`);
    expect(runs[0]?.pieces.map((piece) => piece.kind)).toStrictEqual(["text", "tab", "text"]);
  });

  it("reads an explicit line break", () => {
    const runs = runsOf(`<w:p><w:r><w:t>a</w:t><w:br/><w:t>b</w:t></w:r></w:p>`);
    expect(runs[0]?.pieces.map((piece) => piece.kind)).toStrictEqual(["text", "break", "text"]);
  });

  // A break says which of the two it is, and everything but a page break ends the
  // line without ending the page.
  it("tells a break that ends the page from one that ends the line", () => {
    const breaks = `<w:br/><w:br w:type="textWrapping"/><w:br w:type="page"/><w:br w:type="column"/>`;
    const runs = runsOf(`<w:p><w:r>${breaks}</w:r></w:p>`);

    expect(
      runs[0]?.pieces.map((piece) => (piece.kind === "break" ? piece.endsPage : piece.kind)),
    ).toStrictEqual([false, false, true, false]);
  });

  // A carriage return is the line ending a `w:br` of no type is, spelled the way the
  // producers that are not Word spell it. It ends neither a page nor a column.
  it("reads a carriage return as the line break it is", () => {
    const runs = runsOf(`<w:p><w:r><w:t>a</w:t><w:cr/><w:t>b</w:t></w:r></w:p>`);
    expect(runs[0]?.pieces).toStrictEqual([
      { kind: "text", text: "a" },
      { kind: "break", endsPage: false, endsColumn: false },
      { kind: "text", text: "b" },
    ]);
  });

  // A run holding nothing but a carriage return is dropped before it is read unless
  // the paragraph counts one as content, which runs the text either side together.
  it("keeps a carriage return that is a run of its own", () => {
    const runs = runsOf(
      `<w:p><w:r><w:t>a</w:t></w:r><w:r><w:cr/></w:r><w:r><w:t>b</w:t></w:r></w:p>`,
    );

    expect(runs.map((run) => run.pieces.map((piece) => piece.kind))).toStrictEqual([
      ["text"],
      ["break"],
      ["text"],
    ]);
  });

  // A run of its own is where Word puts a break the author typed between two runs
  // of text, and a run carrying nothing else still ends the line it sits on.
  it("keeps a break that is a run of its own", () => {
    const runs = runsOf(
      `<w:p><w:r><w:t>a</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>b</w:t></w:r></w:p>`,
    );

    expect(runs.map((run) => run.pieces.map((piece) => piece.kind))).toStrictEqual([
      ["text"],
      ["break"],
      ["text"],
    ]);
  });

  it("carries the size a drawing takes on the line", () => {
    const drawing = `<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
      <wp:inline><wp:extent cx="914400" cy="457200"/></wp:inline></w:drawing>`;
    const runs = runsOf(`<w:p><w:r>${drawing}</w:r></w:p>`);

    expect(runs[0]?.pieces).toStrictEqual([
      { kind: "drawing", widthEmu: 914400, heightEmu: 457200, turnDegrees: 0 },
    ]);
  });

  // The extent is the drawing the right way up whatever the turn says, so the turn
  // is carried beside it rather than worked into it.
  it("carries how far round a drawing was turned after it was drawn", () => {
    const drawing = `<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <wp:inline><wp:extent cx="914400" cy="457200"/>
        <a:xfrm rot="5400000"><a:ext cx="914400" cy="457200"/></a:xfrm></wp:inline></w:drawing>`;
    const runs = runsOf(`<w:p><w:r>${drawing}</w:r></w:p>`);

    expect(runs[0]?.pieces).toStrictEqual([
      { kind: "drawing", widthEmu: 914400, heightEmu: 457200, turnDegrees: 90 },
    ]);
  });

  it("resolves each run's own font and size through the cascade", () => {
    const body = `<w:p><w:r><w:t>a</w:t></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Georgia"/><w:sz w:val="48"/></w:rPr><w:t>b</w:t></w:r></w:p>`;
    const runs = runsOf(body);

    expect(runs[0]?.mark).toStrictEqual({
      font: { kind: "named", name: "Arial" },
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
      highlight: null,
      capitals: "none",
    });
    expect(runs[1]?.mark).toStrictEqual({
      font: { kind: "named", name: "Georgia" },
      fontSizePt: 24,
      bold: false,
      italic: false,
      underline: false,
      raisePt: 0,
      lineSizePt: 24,
      lineRaisePt: 0,
      color: null,
      characterSpacingPt: 0,
      characterScale: 1,
      kernFromHalfPoints: null,
      highlight: null,
      capitals: "none",
    });
  });

  it("marks a bold run bold, so it is measured with the bold face", () => {
    const body = `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r></w:p>`;
    expect(runs0(body).bold).toBe(true);
  });

  it("lets a run turn an inherited bold back off", () => {
    const styles = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/><w:b/></w:rPr></w:style></w:styles>`;
    const body = `<w:p><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>a</w:t></w:r></w:p>`;

    expect(runs0(`<w:p><w:r><w:t>a</w:t></w:r></w:p>`, styles).bold).toBe(true);
    expect(runs0(body, styles).bold).toBe(false);
  });

  it("skips a run that places nothing on the line", () => {
    const anchor = `<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
      <wp:anchor><wp:extent cx="914400" cy="457200"/></wp:anchor></w:drawing>`;
    expect(runsOf(`<w:p><w:r>${anchor}</w:r><w:r><w:t>a</w:t></w:r></w:p>`)).toHaveLength(1);
  });

  it("reads no runs from an empty spacer paragraph", () => {
    expect(runsOf(`<w:p/>`)).toStrictEqual([]);
  });

  // **An equation of runs alone is text on the line like any other**, in the math
  // font and spelled in the Mathematical Alphanumeric block. Anything stacked is
  // refused whole and draws nothing until there is geometry for it.
  const MATH = `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"`;

  it("reads an equation's runs as text on the line", () => {
    const equation = `<m:oMath ${MATH}><m:r><m:t>bandril</m:t></m:r></m:oMath>`;
    expect(textOf(runsOf(`<w:p>${equation}</w:p>`))).toBe(
      "\u{1D44F}\u{1D44E}\u{1D45B}\u{1D451}\u{1D45F}\u{1D456}\u{1D459}",
    );
  });

  it("keeps an equation's runs where the paragraph holds them", () => {
    const equation = `<m:oMath ${MATH}><m:r><m:t>x</m:t></m:r></m:oMath>`;
    const body =
      `<w:p><w:r><w:t xml:space="preserve">one </w:t></w:r>${equation}` +
      `<w:r><w:t xml:space="preserve"> two</w:t></w:r></w:p>`;
    expect(textOf(runsOf(body))).toBe("one \u{1D465} two");
  });

  it("spells a run stating m:nor as it is written", () => {
    const equation = `<m:oMath ${MATH}><m:r><m:rPr><m:nor/></m:rPr><m:t>bandril</m:t></m:r></m:oMath>`;
    expect(textOf(runsOf(`<w:p>${equation}</w:p>`))).toBe("bandril");
  });

  it("sets a math run naming no face in the document's math font", () => {
    const equation = `<m:oMath ${MATH}><m:r><m:t>x</m:t></m:r></m:oMath>`;
    const mark = runsOf(`<w:p>${equation}</w:p>`, NO_FACE)[0]?.mark;
    expect(mark?.font).toStrictEqual({ kind: "named", name: "Cambria Math" });
  });

  // Word draws the letters themselves slanted rather than asking for a slanted face,
  // which is what the pdf of every style says: nothing is bold or italic anywhere.
  it("leaves a math run upright and unbolded whatever its style states", () => {
    const equation = `<m:oMath ${MATH}><m:r><m:rPr><m:sty m:val="bi"/></m:rPr><m:t>x</m:t></m:r></m:oMath>`;
    const mark = runsOf(`<w:p>${equation}</w:p>`)[0]?.mark;
    expect(mark?.bold).toBe(false);
    expect(mark?.italic).toBe(false);
  });

  // **A fraction comes back as one piece, never as its halves.** The runs standing
  // inside it are handed out by `paragraphRuns` so that the cascade marks each of
  // them, and they are gathered back into the shape the reader found before anything
  // reaches a line: a line that saw a half would lay the two side by side.
  const FRACTION =
    `<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num>` + `<m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>`;

  it("gathers a fraction's halves back into one piece", () => {
    const runs = runsOf(`<w:p><m:oMath ${MATH}>${FRACTION}</m:oMath></w:p>`);

    expect(runs).toHaveLength(1);
    expect(runs[0]?.pieces.map((each) => each.kind)).toStrictEqual(["equation"]);
  });

  it("keeps the shape the reader found, and the mark of every run in it", () => {
    const piece = runsOf(`<w:p><m:oMath ${MATH}>${FRACTION}</m:oMath></w:p>`, NO_FACE)[0]
      ?.pieces[0];
    if (piece?.kind !== "equation") throw new Error("not an equation");
    const [fraction] = piece.content;
    if (fraction?.kind !== "fraction") throw new Error("not a fraction");

    expect(fraction.numerator.map((each) => (each.kind === "run" ? each.text : ""))).toStrictEqual([
      "\u{1D44E}",
    ]);
    expect(
      fraction.denominator.map((each) => (each.kind === "run" ? each.text : "")),
    ).toStrictEqual(["\u{1D44F}"]);
    const [top] = fraction.numerator;
    if (top?.kind !== "run") throw new Error("no numerator run");
    expect(top.mark.font).toStrictEqual({ kind: "named", name: "Cambria Math" });
  });

  // **A structure takes the mark of the first run it holds at whatever depth that
  // stands.** A delimiter round a fraction holds no run of its own, and looking only
  // at what stands directly inside it left it with no mark and dropped the whole
  // delimiter: the paragraph came out empty and Word's own page holds a fraction in
  // brackets.
  it("marks a delimiter that holds no run of its own", () => {
    const round = `<m:d><m:e>${FRACTION}</m:e></m:d>`;
    const piece = runsOf(`<w:p><m:oMath ${MATH}>${round}</m:oMath></w:p>`, NO_FACE)[0]?.pieces[0];
    if (piece?.kind !== "equation") throw new Error("not an equation");
    const [delimiter] = piece.content;
    if (delimiter?.kind !== "delimiter") throw new Error("not a delimiter");

    expect(delimiter.mark.font).toStrictEqual({ kind: "named", name: "Cambria Math" });
    expect(delimiter.content.map((each) => each.kind)).toStrictEqual(["fraction"]);
  });

  /**
   * **A break inside an equation ends the line, and what follows it is an equation of
   * its own.** Measured on 2026-08-14 by the authored `equation-break-probe` document,
   * nine cases three times each against Word's own pdf: a display fraction alone is
   * 27.36pt, the same with a break after it 42.00, with a break before it 40.08, with a
   * run after the break 40.08, and with a second fraction after it 54.48, which is what
   * the same two equations with an ordinary `w:r` break between them come to as well.
   */
  const BROKEN = `${FRACTION}<m:r><w:br/></m:r>`;

  it("ends the line at a break inside an equation", () => {
    const runs = runsOf(`<w:p><m:oMath ${MATH}>${BROKEN}</m:oMath></w:p>`);

    expect(runs.flatMap((run) => run.pieces.map((each) => each.kind))).toStrictEqual([
      "equation",
      "break",
    ]);
  });

  it("hands out what follows the break as an equation of its own", () => {
    const runs = runsOf(`<w:p><m:oMath ${MATH}>${BROKEN}${FRACTION}</m:oMath></w:p>`);

    expect(runs.flatMap((run) => run.pieces.map((each) => each.kind))).toStrictEqual([
      "equation",
      "break",
      "equation",
    ]);
  });

  // A break before anything else opens with the break rather than with an empty
  // equation, which would be a piece holding nothing to set.
  it("hands out no equation in front of a break that opens one", () => {
    const runs = runsOf(`<w:p><m:oMath ${MATH}><m:r><w:br/></m:r>${FRACTION}</m:oMath></w:p>`);

    expect(runs.flatMap((run) => run.pieces.map((each) => each.kind))).toStrictEqual([
      "break",
      "equation",
    ]);
  });

  // **The break is handed out under the mark of the run carrying it**, since that is
  // what the line it ends is measured from. See `heldOpenPt` in `layout/lines.ts`.
  it("carries the break under its own run's mark", () => {
    const sized = `${FRACTION}<m:r><w:rPr><w:sz w:val="72"/></w:rPr><w:br/></m:r>`;
    const runs = runsOf(`<w:p><m:oMath ${MATH}>${sized}</m:oMath></w:p>`);
    const broke = runs.find((run) => run.pieces.some((each) => each.kind === "break"));

    expect(broke?.mark.fontSizePt).toBe(36);
  });

  // Neither of the two a break can be: it ends a line and nothing else, since the
  // reader refuses an equation holding one that starts a page or a column.
  it("ends no page and no column", () => {
    const runs = runsOf(`<w:p><m:oMath ${MATH}>${BROKEN}</m:oMath></w:p>`);
    const piece = runs.flatMap((run) => run.pieces).find((each) => each.kind === "break");
    if (piece?.kind !== "break") throw new Error("no break");

    expect(piece.endsPage).toBe(false);
    expect(piece.endsColumn).toBe(false);
  });

  // The text either side of an equation is untouched by any of this.
  it("leaves the runs standing beside a fraction where they were", () => {
    const runs = runsOf(
      `<w:p><w:r><w:t>one</w:t></w:r><m:oMath ${MATH}>${FRACTION}</m:oMath>` +
        `<w:r><w:t>two</w:t></w:r></w:p>`,
    );

    expect(runs.flatMap((run) => run.pieces.map((each) => each.kind))).toStrictEqual([
      "text",
      "equation",
      "text",
    ]);
  });

  // An equation of runs alone is drawn where a paragraph's own runs are, so it is
  // left flat and never becomes a piece of its own.
  it("leaves an equation of runs alone flat", () => {
    const runs = runsOf(`<w:p><m:oMath ${MATH}><m:r><m:t>x</m:t></m:r></m:oMath></w:p>`);

    expect(runs.flatMap((run) => run.pieces.map((each) => each.kind))).toStrictEqual(["text"]);
  });

  // **Capitals, measured against Word's own pdf on 2026-08-13.** `w:caps` draws every
  // letter as a capital at the run's own size; `w:smallCaps` draws a letter that was
  // not one already at four fifths of that size and leaves digits and marks alone.
  const CAPS = `<w:caps/>`;
  const SMALL = `<w:smallCaps/>`;

  it("draws every letter of a w:caps run as a capital", () => {
    const body = `<w:p><w:r><w:rPr>${CAPS}</w:rPr><w:t>Bandril 12 (a-b)</w:t></w:r></w:p>`;
    expect(textOf(runsOf(body))).toBe("BANDRIL 12 (A-B)");
    expect(runsOf(body)).toHaveLength(1);
  });

  // Word drew this one as itself under both, since its capital is two letters.
  it("leaves a letter whose capital is two letters as it is", () => {
    const body = `<w:p><w:r><w:rPr>${CAPS}</w:rPr><w:t>maß</w:t></w:r></w:p>`;
    expect(textOf(runsOf(body))).toBe("MAß");
  });

  it("sets a small capital at four fifths of the run's size", () => {
    const body = `<w:p><w:r><w:rPr>${SMALL}</w:rPr><w:t>Bandril</w:t></w:r></w:p>`;
    const runs = runsOf(body);
    expect(textOf(runs)).toBe("BANDRIL");
    expect(runs.map((run) => run.mark.fontSizePt)).toStrictEqual([12, 9.5]);
  });

  it("leaves a digit and a mark at the run's own size", () => {
    const body = `<w:p><w:r><w:rPr>${SMALL}</w:rPr><w:t>a1(b)</w:t></w:r></w:p>`;
    const runs = runsOf(body);
    expect(textOf(runs)).toBe("A1(B)");
    expect(runs.map((run) => run.mark.fontSizePt)).toStrictEqual([9.5, 12, 9.5, 12]);
  });

  // The 20pt small capitals stood on the 20pt line, so what a small capital is drawn
  // at is not what its line is measured at.
  it("measures a small capital's line at the size the run states", () => {
    const body = `<w:p><w:r><w:rPr>${SMALL}</w:rPr><w:t>ab</w:t></w:r></w:p>`;
    expect(runsOf(body)[0]?.mark.lineSizePt).toBe(12);
  });

  // Four fifths of eleven is 8.8 and of thirteen is 10.4, and Word set them at 9 and
  // 10.5: the nearest half point, not the one below.
  it("rounds a small capital's size to the nearest half point", () => {
    const at = (halfPoints: number): number | undefined => {
      const body =
        `<w:p><w:r><w:rPr>${SMALL}<w:sz w:val="${String(halfPoints)}"/></w:rPr>` +
        `<w:t>a</w:t></w:r></w:p>`;
      return runsOf(body)[0]?.mark.fontSizePt;
    };
    expect([at(22), at(26), at(28), at(42)]).toStrictEqual([9, 10.5, 11, 17]);
  });

  // Every space came out at the small size, the ones after a digit and after a
  // bracket included, and no other character with no capital of its own did.
  it("sets every space in a small capital run small", () => {
    const body = `<w:p><w:r><w:rPr>${SMALL}</w:rPr><w:t xml:space="preserve">1 a 2</w:t></w:r></w:p>`;
    const runs = runsOf(body);
    expect(runs.map((run) => [run.mark.fontSizePt, textOf([run])])).toStrictEqual([
      [12, "1"],
      [9.5, " A "],
      [12, "2"],
    ]);
  });

  it("draws a run stating both as w:caps", () => {
    const body = `<w:p><w:r><w:rPr>${CAPS}${SMALL}</w:rPr><w:t>ab</w:t></w:r></w:p>`;
    const runs = runsOf(body);
    expect(textOf(runs)).toBe("AB");
    expect(runs.map((run) => run.mark.fontSizePt)).toStrictEqual([12]);
  });

  it("lets a run turn an inherited w:caps off", () => {
    const styles = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/><w:caps/></w:rPr></w:style></w:styles>`;
    const body = `<w:p><w:r><w:rPr><w:caps w:val="0"/></w:rPr><w:t>ab</w:t></w:r></w:p>`;
    expect(textOf(runsOf(`<w:p><w:r><w:t>ab</w:t></w:r></w:p>`, styles))).toBe("AB");
    expect(textOf(runsOf(body, styles))).toBe("ab");
  });
});

function runs0(body: string, stylesXml: string = NORMAL): TextRun["mark"] {
  const mark = runsOf(body, stylesXml)[0]?.mark;
  if (mark === undefined) throw new Error("expected a run");
  return mark;
}
