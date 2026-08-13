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
});

function runs0(body: string, stylesXml: string = NORMAL): TextRun["mark"] {
  const mark = runsOf(body, stylesXml)[0]?.mark;
  if (mark === undefined) throw new Error("expected a run");
  return mark;
}
