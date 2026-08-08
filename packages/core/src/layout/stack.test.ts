import { describe, expect, it } from "vitest";

import { readBlocks } from "../docx/blocks.js";
import { openDocx } from "../docx/package.js";
import { DEFAULT_SETTINGS, type DocumentSettings } from "../docx/settings.js";
import { readStyleTable } from "../docx/styles.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { buildFace } from "../testing/build-font.js";
import { lookupFontMetrics } from "./font-metrics.js";
import { measureStack, type BandResolver, type ParagraphBox } from "./stack.js";
import type { WrapBand } from "./wrapping.js";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

const NORMAL = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr></w:style></w:styles>`;

const measure = (
  body: string,
  stylesXml: string = NORMAL,
  widthPt = 468,
  settings: DocumentSettings = DEFAULT_SETTINGS,
) => {
  const pkg = openDocx(
    buildDocx({ "word/document.xml": wordDocument(body), "word/styles.xml": stylesXml }),
  );
  return measureStack({
    blocks: readBlocks(pkg),
    styles: readStyleTable(pkg),
    metricsFor: (request) => lookupFontMetrics(request, [ARIAL]),
    part: "word/document.xml",
    originPt: 36,
    leftPt: 72,
    widthPt,
    settings,
  });
};

function boxesOf(
  body: string,
  stylesXml: string = NORMAL,
  widthPt = 468,
  settings: DocumentSettings = DEFAULT_SETTINGS,
) {
  const result = measure(body, stylesXml, widthPt, settings);
  if (result.kind !== "measured") throw new Error(result.blocker.kind);
  return result.boxes;
}

function firstBox(body: string, stylesXml: string = NORMAL, widthPt = 468) {
  const box = boxesOf(body, stylesXml, widthPt)[0];
  if (box === undefined) throw new Error("expected a paragraph");
  return box;
}

const paragraph = (properties: string, text = "aaaa bbbb") =>
  `<w:p><w:pPr>${properties}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

const textOf = (box: ParagraphBox): readonly string[] =>
  box.lines.map((placed) =>
    placed.line.segments.map((segment) => (segment.kind === "text" ? segment.text : "")).join(""),
  );

// Arial's own metrics, so the heights below stay Word's, with an invented set of
// widths so text can be measured at all.
const ARIAL = buildFace({
  name: "Arial",
  metrics: { unitsPerEm: 2048, ascender: 1854, descender: -434, lineGap: 67 },
});

const ARIAL_12 = 13.798828125;

describe("measureStack", () => {
  it("stacks empty spacer paragraphs from the origin down", () => {
    const result = measure(`<w:p/><w:p/><w:p/>`);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);

    expect(result.boxes.map((box) => box.topPt)).toStrictEqual([
      36,
      36 + ARIAL_12,
      36 + ARIAL_12 * 2,
    ]);
    expect(result.heightPt).toBeCloseTo(ARIAL_12 * 3, 9);
  });

  it("scales a paragraph's height with its own mark size", () => {
    const body = `<w:p><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr></w:p>`;
    const result = measure(body);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.boxes[0]?.heightPt).toBeCloseTo(ARIAL_12 * 2, 9);
  });

  it("takes the tallest run, not the paragraph mark, when a run is bigger", () => {
    const body = `<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr>
      <w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t>Heading</w:t></w:r></w:p>`;
    const result = measure(body);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.boxes[0]?.heightPt).toBeCloseTo(22 * 1.14990234375, 9);
  });

  it("uses an inline drawing's stored height when it exceeds the text line", () => {
    const body = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="${WP_NS}">
      <wp:extent cx="2857500" cy="1828800"/></wp:inline></w:drawing></w:r></w:p>`;
    const result = measure(body);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.boxes[0]?.heightPt).toBeCloseTo(144, 6);
  });

  it("blocks on a font whose metrics are unknown instead of guessing", () => {
    const styles = NORMAL.replace('w:ascii="Arial"', 'w:ascii="Meridian Sans"');
    const result = measure(`<w:p/>`, styles);
    if (result.kind !== "blocked") throw new Error("expected to be blocked");
    expect(result.blocker).toStrictEqual({
      kind: "unknown-font-metrics",
      part: "word/document.xml",
      paragraphIndex: 0,
      fontName: "Meridian Sans",
    });
  });

  it("blocks when the cascade resolves no font at all", () => {
    const empty = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;
    const result = measure(`<w:p/>`, empty);
    if (result.kind !== "blocked") throw new Error("expected to be blocked");
    expect(result.blocker.kind).toBe("unresolved-font");
  });

  it("measures an empty document as zero height", () => {
    const result = measure(``);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result).toStrictEqual({
      kind: "measured",
      boxes: [],
      cells: [],
      untornRows: [],
      anchoredObjects: [],
      heightPt: 0,
    });
  });
});

// Every glyph in the test face is half an em, so 12pt text is 6pt a character.
const PER_CHARACTER = 6;
const ARIAL_ASCENT_12 = (12 * (1854 + 67)) / 2048;

describe("measureStack over text", () => {
  it("carries the lines a paragraph breaks into", () => {
    const box = firstBox(paragraph(``), NORMAL, 30);

    expect(textOf(box)).toStrictEqual(["aaaa", "bbbb"]);
    expect(box.heightPt).toBeCloseTo(ARIAL_12 * 2, 9);
  });

  it("stacks each line below the one before it", () => {
    const box = firstBox(paragraph(``), NORMAL, 30);

    expect(box.lines.map((placed) => placed.topPt)).toStrictEqual([36, 36 + ARIAL_12]);
  });

  it("seats a line's baseline its own ascent below its top", () => {
    const box = firstBox(paragraph(``));
    expect(box.lines[0]?.baselinePt).toBeCloseTo(36 + ARIAL_ASCENT_12, 9);
  });

  it("starts a left-aligned line at the frame's left edge", () => {
    expect(firstBox(paragraph(``)).lines[0]?.leftPt).toBe(72);
  });

  it("centres a centred line inside the frame", () => {
    const box = firstBox(paragraph(`<w:jc w:val="center"/>`, "aaaa"));
    expect(box.lines[0]?.leftPt).toBeCloseTo(72 + (468 - 4 * PER_CHARACTER) / 2, 9);
  });

  it("ends a right-aligned line at the frame's right edge", () => {
    const box = firstBox(paragraph(`<w:jc w:val="right"/>`, "aaaa"));
    expect(box.lines[0]?.leftPt).toBeCloseTo(72 + 468 - 4 * PER_CHARACTER, 9);
  });

  // Measured against Word: every line of a justified paragraph but its last fills
  // the width, and one that ended at a manual break fills it too.
  it("fills the width of every justified line but the paragraph's last", () => {
    const box = firstBox(paragraph(`<w:jc w:val="both"/>`, "aa bb cccccc"), NORMAL, 40);
    const widths = box.lines.map((placed) => placed.line.widthPt);

    expect(textOf(box)).toStrictEqual(["aa bb", "cccccc"]);
    expect(widths[0]).toBeCloseTo(40, 9);
    expect(widths[1]).toBeCloseTo(6 * PER_CHARACTER, 9);
  });

  // The space a line broke at hangs past its edge rather than sitting on it, so a
  // line with no space of its own has nowhere to put the room it did not fill.
  it("leaves a justified line that broke at its only space where it fell", () => {
    const box = firstBox(paragraph(`<w:jc w:val="both"/>`), NORMAL, 30);

    expect(textOf(box)).toStrictEqual(["aaaa", "bbbb"]);
    expect(box.lines[0]?.line.widthPt).toBeCloseTo(4 * PER_CHARACTER, 9);
  });

  it("stretches a justified line at its spaces, leaving its first word alone", () => {
    const box = firstBox(paragraph(`<w:jc w:val="both"/>`, "aa bb cccccc"), NORMAL, 40);
    const [first] = box.lines;

    // "aa bb" is five characters of the forty the line has room for, and its one
    // space takes all ten points of what is left.
    expect(first?.leftPt).toBe(72);
    expect(first?.line.segments.map((segment) => segment.offsetPt)).toStrictEqual([
      0,
      2 * PER_CHARACTER,
      3 * PER_CHARACTER + 10,
    ]);
  });

  it("stretches the line a manual break ended, which is not the paragraph's last", () => {
    const body =
      `<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>aa bb</w:t>` +
      `<w:br/><w:t>cc dd</w:t></w:r></w:p>`;
    const box = firstBox(body, NORMAL, 100);

    expect(box.lines.map((placed) => placed.line.widthPt)).toStrictEqual([100, 5 * PER_CHARACTER]);
  });

  it("leaves a justified line alone when nothing on it can grow", () => {
    const box = firstBox(paragraph(`<w:jc w:val="both"/>`, "aaaa"), NORMAL, 100);
    expect(box.lines[0]?.line.widthPt).toBeCloseTo(4 * PER_CHARACTER, 9);
  });

  it("moves an indented line in and breaks it at the narrower width", () => {
    const box = firstBox(paragraph(`<w:ind w:left="720" w:right="720"/>`, "aaaa"), NORMAL, 100);
    expect(box.lines[0]?.leftPt).toBeCloseTo(72 + 36, 9);
  });

  it("indents only the first line when the paragraph asks for a first-line indent", () => {
    const box = firstBox(paragraph(`<w:ind w:firstLine="360"/>`), NORMAL, 30);

    expect(box.lines[0]?.leftPt).toBeCloseTo(72 + 18, 9);
    expect(box.lines[1]?.leftPt).toBeCloseTo(72, 9);
  });

  it("pulls a hanging indent's first line back out of the left indent", () => {
    const box = firstBox(paragraph(`<w:ind w:left="720" w:hanging="360"/>`, "aaaa"));

    expect(box.lines[0]?.leftPt).toBeCloseTo(72 + 36 - 18, 9);
  });

  it("adds the space a paragraph asks for before and after itself", () => {
    const box = firstBox(paragraph(`<w:spacing w:before="240" w:after="120"/>`, "aaaa"));

    expect(box.topPt).toBe(36);
    expect(box.lines[0]?.topPt).toBeCloseTo(36 + 12, 9);
    expect(box.heightPt).toBeCloseTo(12 + ARIAL_12 + 6, 9);
  });

  it("multiplies a line's height when the rule is auto", () => {
    const box = firstBox(paragraph(`<w:spacing w:line="276" w:lineRule="auto"/>`, "aaaa"));
    expect(box.heightPt).toBeCloseTo((ARIAL_12 * 276) / 240, 9);
  });

  // A multiple is taken of the line the faces make and what a raise added stands on
  // top of it: measured on 2026-08-07 by the authored `raised-text` document, where
  // 12pt text raised six points under a line and a half came out 27.96pt rather
  // than the 30.96 half again of the grown line would have made.
  it("takes a multiple of the line the faces make and adds the raise to it", () => {
    const body =
      `<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>` +
      `<w:r><w:t>aaaa</w:t></w:r>` +
      `<w:r><w:rPr><w:position w:val="12"/></w:rPr><w:t>bbbb</w:t></w:r></w:p>`;

    expect(firstBox(body).heightPt).toBeCloseTo(ARIAL_12 + 6 + ARIAL_12 / 2, 9);
  });

  // A line told exactly how tall to be has nowhere to grow, and the run is still
  // drawn its six points off the baseline.
  it("leaves a line told its own height alone when a run on it is raised", () => {
    const body =
      `<w:p><w:pPr><w:spacing w:line="480" w:lineRule="exact"/></w:pPr>` +
      `<w:r><w:t>aaaa</w:t></w:r>` +
      `<w:r><w:rPr><w:position w:val="12"/></w:rPr><w:t>bbbb</w:t></w:r></w:p>`;

    expect(firstBox(body).heightPt).toBeCloseTo(24, 9);
  });

  it("replaces a line's height when the rule is exact", () => {
    const box = firstBox(paragraph(`<w:spacing w:line="400" w:lineRule="exact"/>`, "aaaa"));
    expect(box.heightPt).toBeCloseTo(20, 9);
  });

  it("takes a line's own height when the rule only sets a floor under it", () => {
    const box = firstBox(paragraph(`<w:spacing w:line="120" w:lineRule="atLeast"/>`, "aaaa"));
    expect(box.heightPt).toBeCloseTo(ARIAL_12, 9);
  });

  // Word leaves the text at the top of the taller line and lets the leading fall
  // below it, which is where the first line of a spaced-out paragraph sits.
  it("puts extra leading below the text, where Word puts it", () => {
    const box = firstBox(paragraph(`<w:spacing w:line="480" w:lineRule="auto"/>`, "aaaa"));
    expect(box.lines[0]?.baselinePt).toBeCloseTo(36 + ARIAL_ASCENT_12, 9);
  });

  it("leaves the last line alone when the paragraph mark is bigger than its runs", () => {
    const body = `<w:p><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr>
      <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>aaaa</w:t></w:r></w:p>`;

    expect(firstBox(body).heightPt).toBeCloseTo(ARIAL_12, 9);
  });

  // Nothing a line carries but white raises it: what holds such a line open is the
  // paragraph's own mark, whatever the run the white is written in asks for.
  it("stands a line held open by a tab alone at its paragraph mark's height", () => {
    const body = `<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>
      <w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:tab/></w:r></w:p>`;

    expect(firstBox(body).heightPt).toBeCloseTo(ARIAL_12, 9);
  });

  it("stands a line holding one space at its paragraph mark's height", () => {
    const body = `<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>
      <w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r></w:p>`;

    expect(firstBox(body).heightPt).toBeCloseTo(ARIAL_12, 9);
  });

  it("leaves a space between two words out of the height of the line it sits on", () => {
    const wide = `<w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>`;
    const body = `<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>aa</w:t></w:r>
      ${wide}<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>bb</w:t></w:r></w:p>`;

    expect(firstBox(body).heightPt).toBeCloseTo(ARIAL_12, 9);
  });

  it("keeps an empty paragraph at its mark's height", () => {
    expect(firstBox(`<w:p/>`).lines).toStrictEqual([]);
  });

  // What the break itself acts on: the line a break of the paragraph's own put at
  // the head of a page, the ask a paragraph carries, and the break it ended on.
  it("carries a page break through to the line it starts", () => {
    const body = `<w:p><w:r><w:t>aa</w:t><w:br w:type="page"/><w:t>bb</w:t></w:r></w:p>`;
    expect(firstBox(body).lines.map((line) => line.startsPage)).toStrictEqual([false, true]);
  });

  it("carries a paragraph's ask for a page of its own", () => {
    const body = `<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>aa</w:t></w:r></w:p>`;
    expect(firstBox(body).startsPage).toBe(true);
  });

  it("says a paragraph ended on a page break, which draws no line to say it", () => {
    const body = `<w:p><w:r><w:t>aa</w:t><w:br w:type="page"/></w:r></w:p>`;
    const box = firstBox(body);

    expect(box.lines).toHaveLength(1);
    expect(box.endsPage).toBe(true);
  });

  // A break with nothing on the line it ended still holds the room that line takes
  // on the page it is leaving, which is the mark's own height.
  it("gives the line a page break ends on its own the height of the mark", () => {
    const body = `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
    const box = firstBox(body);

    expect(box.lines).toHaveLength(1);
    expect(box.heightPt).toBeCloseTo(ARIAL_12, 9);
  });
});

// The stops a paragraph falls back on are the document's, not a number this
// package chose: a tab in a paragraph declaring none lands on the first of them.
describe("measureStack and the stops the document falls back on", () => {
  const tabbed = (settings: DocumentSettings | undefined) => {
    const pkg = openDocx(
      buildDocx({
        "word/document.xml": wordDocument(`<w:p><w:r><w:tab/><w:t>aaaa</w:t></w:r></w:p>`),
        "word/styles.xml": NORMAL,
      }),
    );
    const result = measureStack({
      blocks: readBlocks(pkg),
      styles: readStyleTable(pkg),
      metricsFor: (request) => lookupFontMetrics(request, [ARIAL]),
      part: "word/document.xml",
      originPt: 36,
      leftPt: 72,
      widthPt: 468,
      ...(settings === undefined ? {} : { settings }),
    });
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    const segment = result.boxes[0]?.lines[0]?.line.segments.at(-1);
    if (segment === undefined) throw new Error("expected a segment after the tab");
    return segment.offsetPt;
  };

  it("starts the text after a tab at the stop the settings space out", () => {
    expect(tabbed({ ...DEFAULT_SETTINGS, defaultTabStopTwips: 567 })).toBeCloseTo(28.35, 6);
  });

  it("falls back to Word's own spacing where nothing has read the settings", () => {
    expect(tabbed(undefined)).toBeCloseTo(36, 6);
  });
});

// Which Word a document was written for decides where a table's indent is
// measured to, so every table below says which one it stands in.
const MODERN: DocumentSettings = { ...DEFAULT_SETTINGS, compatibilityMode: 15 };
const LEGACY: DocumentSettings = { ...DEFAULT_SETTINGS, compatibilityMode: null };

describe("measureStack over tables", () => {
  it("gives a row the height of its tallest cell, not the sum of them", () => {
    const result = measure(table(cell(`<w:p/>`), cell(`<w:p/><w:p/><w:p/>`)));
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.heightPt).toBeCloseTo(ARIAL_12 * 3, 9);
  });

  it("starts every cell of a row at the row's top", () => {
    const result = measure(table(cell(`<w:p/>`), cell(`<w:p/><w:p/>`)));
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.boxes.map((box) => box.topPt)).toStrictEqual([36, 36, 36 + ARIAL_12]);
  });

  it("continues the stack below the table rather than below its tallest cell alone", () => {
    const result = measure(`${table(cell(`<w:p/><w:p/>`))}<w:p/>`);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.boxes[2]?.topPt).toBeCloseTo(36 + ARIAL_12 * 2, 9);
  });

  it("centres a short cell in a taller row when the cell asks for it", () => {
    const centred = `<w:vAlign w:val="center"/>`;
    const result = measure(table(cell(`<w:p/>`, centred), cell(`<w:p/><w:p/><w:p/>`)));
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.boxes[0]?.topPt).toBeCloseTo(36 + ARIAL_12, 9);
  });

  it("seats a bottom-aligned cell against the row's baseline edge", () => {
    const result = measure(
      table(cell(`<w:p/>`, `<w:vAlign w:val="bottom"/>`), cell(`<w:p/><w:p/><w:p/>`)),
    );
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.boxes[0]?.topPt).toBeCloseTo(36 + ARIAL_12 * 2, 9);
  });

  it("stacks rows one below the next", () => {
    const body = `<w:tbl><w:tr>${cell(`<w:p/><w:p/>`)}</w:tr><w:tr>${cell(`<w:p/>`)}</w:tr></w:tbl>`;
    const result = measure(body);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.boxes[2]?.topPt).toBeCloseTo(36 + ARIAL_12 * 2, 9);
    expect(result.heightPt).toBeCloseTo(ARIAL_12 * 3, 9);
  });

  it("holds a cell's text off its edge by the margin Word leaves there", () => {
    const boxes = boxesOf(
      table(cell(`<w:p><w:r><w:t>aaaa</w:t></w:r></w:p>`)),
      NORMAL,
      468,
      MODERN,
    );
    // The frame starts at 72pt and Word's own cell margin is an eighth of an inch.
    expect(boxes[0]?.lines[0]?.leftPt).toBeCloseTo(72 + 5.4, 9);
  });

  it("indents a table by what it asks for, and takes the margin it asks for", () => {
    const properties =
      `<w:tblPr><w:tblInd w:w="-100" w:type="dxa"/>` +
      `<w:tblCellMar><w:left w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>`;
    const body = `<w:tbl>${properties}<w:tr>${cell(`<w:p><w:r><w:t>aaaa</w:t></w:r></w:p>`)}</w:tr></w:tbl>`;
    expect(boxesOf(body, NORMAL, 468, MODERN)[0]?.lines[0]?.leftPt).toBeCloseTo(72 - 5, 9);
  });

  // An old document measures the same indent to the text rather than to the
  // table, so the cell's margin stands outside the indent instead of inside it.
  // Word draws the table's own edge left of the page margin to put it there.
  const indented = (indentTwips: number, marginTwips: number, settings: DocumentSettings) => {
    const properties =
      `<w:tblPr><w:tblInd w:w="${String(indentTwips)}" w:type="dxa"/>` +
      `<w:tblCellMar><w:left w:w="${String(marginTwips)}" w:type="dxa"/></w:tblCellMar></w:tblPr>`;
    const body = `<w:tbl>${properties}<w:tr>${cell(`<w:p><w:r><w:t>aaaa</w:t></w:r></w:p>`)}</w:tr></w:tbl>`;
    return boxesOf(body, NORMAL, 468, settings)[0]?.lines[0]?.leftPt;
  };

  it("measures an old document's indent to the text in its first cell", () => {
    expect(indented(216, 108, LEGACY)).toBeCloseTo(72 + 10.8, 9);
    expect(indented(0, 108, LEGACY)).toBeCloseTo(72, 9);
  });

  it("measures a modern document's indent to the table's own edge", () => {
    expect(indented(216, 108, MODERN)).toBeCloseTo(72 + 10.8 + 5.4, 9);
    expect(indented(0, 108, MODERN)).toBeCloseTo(72 + 5.4, 9);
  });

  it("leaves an indent where it is in either where the cell asks for no margin", () => {
    expect(indented(216, 0, LEGACY)).toBeCloseTo(72 + 10.8, 9);
    expect(indented(216, 0, MODERN)).toBeCloseTo(72 + 10.8, 9);
  });

  it("measures an old document's indent to the first row's cell, not to each row's", () => {
    const properties =
      `<w:tblPr><w:tblInd w:w="216" w:type="dxa"/>` +
      `<w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>`;
    const wide = `<w:tcMar><w:left w:w="288" w:type="dxa"/></w:tcMar>`;
    const rows =
      `<w:tr>${cell(`<w:p><w:r><w:t>aaaa</w:t></w:r></w:p>`)}</w:tr>` +
      `<w:tr>${cell(`<w:p><w:r><w:t>aaaa</w:t></w:r></w:p>`, wide)}</w:tr>`;
    const boxes = boxesOf(`<w:tbl>${properties}${rows}</w:tbl>`, NORMAL, 468, LEGACY);
    // The table's edge is 5.4pt left of the indent, and the second row stands its
    // own 14.4pt off that same edge.
    expect(boxes[0]?.lines[0]?.leftPt).toBeCloseTo(72 + 10.8, 9);
    expect(boxes[1]?.lines[0]?.leftPt).toBeCloseTo(72 + 10.8 - 5.4 + 14.4, 9);
  });

  it("holds every cell of a row off the top wall by the largest margin any of them asks for", () => {
    const held = `<w:tcMar><w:top w:w="288" w:type="dxa"/></w:tcMar>`;
    const result = measure(table(cell(`<w:p/>`), cell(`<w:p/>`, held)));
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.boxes.map((box) => box.topPt)).toStrictEqual([36 + 14.4, 36 + 14.4]);
  });

  it("adds the largest bottom margin in a row to the row without moving its text", () => {
    const held = `<w:tcMar><w:bottom w:w="288" w:type="dxa"/></w:tcMar>`;
    const result = measure(table(cell(`<w:p/>`, held), cell(`<w:p/>`)));
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.boxes[0]?.topPt).toBe(36);
    expect(result.heightPt).toBeCloseTo(ARIAL_12 + 14.4, 9);
  });

  it("holds a cell's text off the side by its own margin, leaving its neighbour the table's", () => {
    const held = `<w:tcMar><w:left w:w="288" w:type="dxa"/></w:tcMar>`;
    const boxes = boxesOf(
      table(cell(`<w:p><w:r><w:t>aaaa</w:t></w:r></w:p>`, held), cell(`<w:p/>`)),
      NORMAL,
      468,
      MODERN,
    );
    expect(boxes[0]?.lines[0]?.leftPt).toBeCloseTo(72 + 14.4, 9);
  });

  it("takes a stated row height as a floor under the row", () => {
    const asked = `<w:trPr><w:trHeight w:val="1440"/></w:trPr>`;
    const body = `<w:tbl><w:tr>${asked}${cell(`<w:p/>`)}</w:tr></w:tbl>`;
    const result = measure(body);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.boxes[0]?.topPt).toBe(36);
    expect(result.heightPt).toBeCloseTo(72, 9);
  });

  it("takes an exact row height as the whole of the row, whatever its cells hold", () => {
    const asked = `<w:trPr><w:trHeight w:val="288" w:hRule="exact"/></w:trPr>`;
    const body = `<w:tbl><w:tr>${asked}${cell(`<w:p/><w:p/>`)}</w:tr></w:tbl>`;
    const result = measure(body);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.heightPt).toBeCloseTo(14.4, 9);
  });

  // Which rows a page break may not be torn through, measured on 2026-08-07 by the
  // authored `tearing` document. A row is torn at a line like anything else unless
  // it says it may not be, or it stands taller than its own text: the empty foot a
  // stated height opens under the last line has no line in it to be torn at.
  const rowsThatRefuse = (properties: string, cells: string): readonly number[] => {
    const result = measure(`<w:tbl><w:tr>${properties}${cells}</w:tr></w:tbl>`);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    return result.untornRows.map((row) => row.bottomPt - row.topPt);
  };

  it("speaks for no ordinary row", () => {
    expect(rowsThatRefuse("", cell(`<w:p/><w:p/>`))).toStrictEqual([]);
  });

  it("speaks for a row saying it may not be split", () => {
    const rows = rowsThatRefuse(`<w:trPr><w:cantSplit/></w:trPr>`, cell(`<w:p/><w:p/>`));
    expect(rows).toHaveLength(1);
    expect(rows[0]).toBeCloseTo(27.6, 1);
  });

  it("speaks for a row asking to stand taller than its own text", () => {
    const asked = `<w:trPr><w:trHeight w:val="1440"/></w:trPr>`;
    expect(rowsThatRefuse(asked, cell(`<w:p/>`))).toStrictEqual([72]);
  });

  it("leaves a row whose text overflows the height it asked for to be torn", () => {
    const asked = `<w:trPr><w:trHeight w:val="144"/></w:trPr>`;
    expect(rowsThatRefuse(asked, cell(`<w:p/><w:p/>`))).toStrictEqual([]);
  });

  // The grid is what a table is drawn on, and a cell that states no width of its
  // own stands in the column it falls in rather than filling the frame.
  const acrossTheGrid = (grid: string, cells: string): readonly number[] => {
    const body = `<w:tbl>${grid}<w:tr>${cells}</w:tr></w:tbl>`;
    const result = measure(body);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    return result.cells.map((each) => each.widthPt);
  };

  const GRID = `<w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="2880"/></w:tblGrid>`;

  it("stands a cell that states no width in the column the grid gives it", () => {
    expect(acrossTheGrid(GRID, cell(`<w:p/>`) + cell(`<w:p/>`))).toStrictEqual([72, 144]);
  });

  it("keeps a cell's own stated width in front of the grid's", () => {
    const own = `<w:tcW w:w="4320" w:type="dxa"/>`;
    expect(acrossTheGrid(GRID, cell(`<w:p/>`, own) + cell(`<w:p/>`))).toStrictEqual([216, 144]);
  });

  it("gives a cell the whole frame where the table declares no grid at all", () => {
    expect(acrossTheGrid("", cell(`<w:p/>`) + cell(`<w:p/>`))).toStrictEqual([468, 468]);
  });

  it("names the paragraph a row opens with, which is where the break is decided", () => {
    const asked = `<w:trPr><w:trHeight w:val="1440"/></w:trPr>`;
    const body = `<w:p/><w:tbl><w:tr>${asked}${cell(`<w:p/>`)}${cell(`<w:p/>`)}</w:tr></w:tbl>`;
    const result = measure(body);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.untornRows[0]?.opensAt).toBe(1);
  });

  // Word ignores a break inside a cell outright: not a page, and not even the line
  // an ordinary break would have ended.
  it("passes over a page break inside a cell", () => {
    const broken = `<w:p><w:r><w:t>aa</w:t><w:br w:type="page"/><w:t>bb</w:t></w:r></w:p>`;
    const boxes = boxesOf(table(cell(broken)));

    expect(boxes[0]?.lines).toHaveLength(1);
    expect(boxes[0]?.endsPage).toBe(false);
  });

  it("leaves a paragraph in a cell no page of its own to ask for", () => {
    const asking = `<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>aa</w:t></w:r></w:p>`;
    expect(boxesOf(table(cell(asking)))[0]?.startsPage).toBe(false);
  });

  it("reports the cell paragraph that blocks measurement", () => {
    const styles = NORMAL.replace('w:ascii="Arial"', 'w:ascii="Meridian Sans"');
    const result = measure(table(cell(`<w:p/>`)), styles);
    if (result.kind !== "blocked") throw new Error("expected to be blocked");
    expect(result.blocker).toStrictEqual({
      kind: "unknown-font-metrics",
      part: "word/document.xml",
      paragraphIndex: 0,
      fontName: "Meridian Sans",
    });
  });
});

const numbering = (levels: string) => `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">${levels}</w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

const decimalLevel = (extra = "") => `<w:lvl w:ilvl="0">
  <w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>${extra}
  <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>`;

const LISTS = numbering(decimalLevel());

const numberedBoxes = (body: string, numberingXml = LISTS, widthPt = 468) => {
  const pkg = openDocx(
    buildDocx({
      "word/document.xml": wordDocument(body),
      "word/styles.xml": NORMAL,
      "word/numbering.xml": numberingXml,
    }),
  );
  const result = measureStack({
    blocks: readBlocks(pkg),
    styles: readStyleTable(pkg),
    metricsFor: (request) => lookupFontMetrics(request, [ARIAL, SYMBOLS, TALL_MARKS, DEEP_MARKS]),
    part: "word/document.xml",
    originPt: 36,
    leftPt: 72,
    widthPt,
  });
  if (result.kind !== "measured") throw new Error(result.blocker.kind);
  return result.boxes;
};

const numberedFirst = (body: string, numberingXml = LISTS, widthPt = 468) => {
  const box = numberedBoxes(body, numberingXml, widthPt)[0];
  if (box === undefined) throw new Error("expected a paragraph");
  return box;
};

const listItem = (text = "aaaa bbbb") =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
     <w:r><w:t>${text}</w:t></w:r></w:p>`;

// A second face, so a level that names its own font is measured in that font's
// widths rather than the paragraph's.
const SYMBOLS = buildFace({
  name: "Symbols",
  metrics: { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 },
  advance: 250,
  characters: "1.",
});

// One face reaching above the text's ascender and one reaching below its
// descender, which are the two halves of what a number can do to a line.
const TALL_MARKS = buildFace({
  name: "Tall Marks",
  metrics: { unitsPerEm: 1000, ascender: 1000, descender: -200, lineGap: 0 },
  advance: 250,
  characters: "1.",
});

const DEEP_MARKS = buildFace({
  name: "Deep Marks",
  metrics: { unitsPerEm: 1000, ascender: 700, descender: -500, lineGap: 0 },
  advance: 250,
  characters: "1.",
});

const levelInFace = (name: string) =>
  numbering(decimalLevel(`<w:rPr><w:rFonts w:ascii="${name}" w:hAnsi="${name}"/></w:rPr>`));

describe("measureStack over a numbered paragraph", () => {
  it("draws the number at the position the hanging indent pulls back to", () => {
    const marker = numberedFirst(listItem()).marker;
    expect(marker?.text).toBe("1.");
    expect(marker?.leftPt).toBeCloseTo(72 + 36 - 18, 9);
    expect(marker?.widthPt).toBeCloseTo(PER_CHARACTER * 2, 9);
  });

  it("starts the first line at the stop the number tabs across to", () => {
    expect(numberedFirst(listItem()).lines[0]?.leftPt).toBeCloseTo(72 + 36, 9);
  });

  it("leaves the lines after the first at the left indent", () => {
    const box = numberedFirst(listItem(), LISTS, 36 + 30);
    expect(box.lines.map((line) => line.leftPt)).toStrictEqual([108, 108]);
  });

  it("puts the number on the first line's baseline", () => {
    const box = numberedFirst(listItem());
    expect(box.marker?.baselinePt).toBe(box.lines[0]?.baselinePt);
  });

  it("still numbers a paragraph with nothing on its line", () => {
    const body = `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:p>`;
    expect(numberedFirst(body).marker?.text).toBe("1.");
  });

  it("counts on down the paragraphs of a list", () => {
    const texts = numberedBoxes(listItem().repeat(3)).map((box) => box.marker?.text);
    expect(texts).toStrictEqual(["1.", "2.", "3."]);
  });

  it("draws no number for a paragraph in no list", () => {
    expect(numberedFirst(`<w:p><w:r><w:t>plain</w:t></w:r></w:p>`).marker).toBeNull();
  });

  it("measures the number in the face its level names", () => {
    const lists = numbering(
      decimalLevel(`<w:rPr><w:rFonts w:ascii="Symbols" w:hAnsi="Symbols"/></w:rPr>`),
    );
    const marker = numberedFirst(listItem(), lists).marker;
    expect(marker?.mark.font).toStrictEqual({ kind: "named", name: "Symbols" });
    expect(marker?.widthPt).toBeCloseTo((12 * 250 * 2) / 1000, 9);
  });

  it("lifts the top of the line by how far its number reaches over the text", () => {
    const box = numberedFirst(listItem("aaaa"), levelInFace("Tall Marks"));
    expect(box.heightPt).toBeCloseTo(ARIAL_12 + (12 - ARIAL_ASCENT_12), 9);
    expect(box.lines[0]?.baselinePt).toBeCloseTo(36 + 12, 9);
  });

  it("leaves the line alone for a number that only reaches below the baseline", () => {
    const box = numberedFirst(listItem("aaaa"), levelInFace("Deep Marks"));
    expect(box.heightPt).toBeCloseTo(ARIAL_12, 9);
    expect(box.lines[0]?.baselinePt).toBeCloseTo(36 + ARIAL_ASCENT_12, 9);
  });

  it("lifts only the line the number sits on", () => {
    const box = numberedFirst(listItem(), levelInFace("Tall Marks"), 36 + 30);
    const [first, second] = box.lines;
    expect((second?.topPt ?? 0) - (first?.topPt ?? 0)).toBeCloseTo(
      ARIAL_12 + (12 - ARIAL_ASCENT_12),
      9,
    );
    expect((second?.baselinePt ?? 0) - (second?.topPt ?? 0)).toBeCloseTo(ARIAL_ASCENT_12, 9);
  });

  it("leaves the text against the number when the level asks for no suffix", () => {
    const lists = numbering(decimalLevel(`<w:suff w:val="nothing"/>`));
    expect(numberedFirst(listItem(), lists).lines[0]?.leftPt).toBeCloseTo(90 + 12, 9);
  });

  it("blocks on a level it cannot count in rather than numbering it wrong", () => {
    const lists = numbering(
      `<w:lvl w:ilvl="0"><w:numFmt w:val="ideographDigital"/><w:lvlText w:val="%1"/></w:lvl>`,
    );
    const pkg = openDocx(
      buildDocx({
        "word/document.xml": wordDocument(listItem()),
        "word/styles.xml": NORMAL,
        "word/numbering.xml": lists,
      }),
    );
    const result = measureStack({
      blocks: readBlocks(pkg),
      styles: readStyleTable(pkg),
      metricsFor: (request) => lookupFontMetrics(request, [ARIAL]),
      part: "word/document.xml",
      originPt: 36,
      leftPt: 72,
      widthPt: 468,
    });
    if (result.kind !== "blocked") throw new Error("expected to be blocked");
    expect(result.blocker).toStrictEqual({
      kind: "unsupported-number-format",
      part: "word/document.xml",
      paragraphIndex: 0,
      numId: "1",
      ilvl: 0,
    });
  });
});

// Every glyph is half an em, so a 12pt line of n characters is 6n wide; the frame
// below runs from 72 to 540.
const wrapped = (body: string, bandsFor: BandResolver) => {
  const pkg = openDocx(
    buildDocx({ "word/document.xml": wordDocument(body), "word/styles.xml": NORMAL }),
  );
  const result = measureStack({
    blocks: readBlocks(pkg),
    styles: readStyleTable(pkg),
    metricsFor: (request) => lookupFontMetrics(request, [ARIAL]),
    part: "word/document.xml",
    originPt: 36,
    leftPt: 72,
    widthPt: 468,
    bandsFor,
  });
  if (result.kind !== "measured") throw new Error(result.blocker.kind);
  return result.boxes;
};

const bandOn =
  (index: number, band: WrapBand): BandResolver =>
  (paragraph) =>
    paragraph.index === index ? [band] : [];

describe("measureStack around wrapping objects", () => {
  it("moves a line's start past an object standing over the frame's left edge", () => {
    const boxes = wrapped(
      paragraph(``, "aaaa"),
      bandOn(0, { leftPt: 0, rightPt: 120, topPt: 0, bottomPt: 100 }),
    );

    expect(boxes[0]?.lines[0]?.leftPt).toBe(120);
    expect(boxes[0]?.lines[0]?.topPt).toBe(36);
  });

  // A line falls to the bottom edge of the object that blocked it, landing on it
  // exactly rather than on a step of its own height.
  it("drops a line to the foot of an object it is too wide to sit beside", () => {
    const boxes = wrapped(
      paragraph(``, "aaaa"),
      bandOn(0, { leftPt: 0, rightPt: 530, topPt: 0, bottomPt: 100 }),
    );

    expect(boxes[0]?.lines[0]?.topPt).toBe(100);
  });

  it("keeps the paragraph's own top where the flow left it, so its floats stay put", () => {
    const boxes = wrapped(
      paragraph(``, "aaaa"),
      bandOn(0, { leftPt: 0, rightPt: 530, topPt: 0, bottomPt: 100 }),
    );

    expect(boxes[0]?.topPt).toBe(36);
    expect(boxes[0]?.heightPt).toBeCloseTo(100 + ARIAL_12 - 36, 9);
  });

  it("starts the next paragraph below the line that fell, not where it would have sat", () => {
    const boxes = wrapped(
      paragraph(``, "aaaa") + paragraph(``, "bbbb"),
      bandOn(0, { leftPt: 0, rightPt: 530, topPt: 0, bottomPt: 100 }),
    );

    expect(boxes[1]?.topPt).toBeCloseTo(100 + ARIAL_12, 9);
  });

  it("leaves a line the object no longer reaches alone", () => {
    const boxes = wrapped(
      paragraph(``, "aaaa") + paragraph(``, "bbbb"),
      bandOn(0, { leftPt: 0, rightPt: 530, topPt: 0, bottomPt: 45 }),
    );

    expect(boxes[1]?.lines[0]?.leftPt).toBe(72);
    expect(boxes[1]?.lines[0]?.topPt).toBeCloseTo(45 + ARIAL_12, 9);
  });

  // Word moves an empty paragraph out of an object's way like any other line, and
  // the paragraphs under it follow, which is what decides where a page breaks.
  it("moves an empty paragraph out of an object's way, though it draws nothing", () => {
    const boxes = wrapped(
      `<w:p/>` + paragraph(``, "aaaa"),
      bandOn(0, { leftPt: 0, rightPt: 530, topPt: 0, bottomPt: 100 }),
    );

    expect(boxes[0]?.lines).toStrictEqual([]);
    expect(boxes[0]?.heightPt).toBeCloseTo(100 + ARIAL_12 - 36, 9);
    expect(boxes[1]?.topPt).toBeCloseTo(100 + ARIAL_12, 9);
  });

  it("keeps an object out of the paragraphs ahead of the one it is anchored to", () => {
    const boxes = wrapped(
      paragraph(``, "aaaa") + paragraph(``, "bbbb"),
      bandOn(1, { leftPt: 0, rightPt: 120, topPt: 0, bottomPt: 100 }),
    );

    expect(boxes[0]?.lines[0]?.leftPt).toBe(72);
    expect(boxes[1]?.lines[0]?.leftPt).toBe(120);
  });

  it("wraps every line of a paragraph, not only its first", () => {
    const boxes = wrapped(
      paragraph(``, `${"a".repeat(40)} ${"b".repeat(40)}`),
      bandOn(0, { leftPt: 0, rightPt: 120, topPt: 0, bottomPt: 100 }),
    );

    expect(boxes[0]?.lines.map((line) => line.leftPt)).toStrictEqual([120, 120]);
  });

  // Six ten-letter words fit on one line of the frame and three on a line of what
  // the object leaves, so a line that is not broken again gives itself away.
  it("breaks a line again at the width an object beside it left it", () => {
    const boxes = wrapped(
      paragraph(``, Array.from({ length: 6 }, () => "a".repeat(10)).join(" ")),
      bandOn(0, { leftPt: 300, rightPt: 540, topPt: 0, bottomPt: 100 }),
    );

    expect(boxes[0]?.lines.map((line) => line.topPt)).toStrictEqual([36, 36 + ARIAL_12]);
    expect(boxes[0]?.lines.map((line) => line.line.widthPt)).toStrictEqual([192, 192]);
  });

  // The word is 60pt and the run of space left over is 40, which no breaking makes
  // it fit into.
  it("falls past an object that leaves less room than the line's first word", () => {
    const boxes = wrapped(
      paragraph(``, `${"a".repeat(10)} ${"b".repeat(10)}`),
      bandOn(0, { leftPt: 0, rightPt: 500, topPt: 0, bottomPt: 100 }),
    );

    expect(boxes[0]?.lines.map((line) => line.topPt)).toStrictEqual([100]);
    expect(boxes[0]?.lines[0]?.leftPt).toBe(72);
  });

  // A tab holds its line open as far as the stop it reached, and nothing about it
  // can be broken, so the line asks for that whole width even though it draws
  // nothing. A document's invisible line falls past its objects on this alone.
  it("asks for the room a tab held a line open to, which no breaking gives back", () => {
    const boxes = wrapped(
      `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="7200"/></w:tabs></w:pPr><w:r><w:tab/></w:r></w:p>`,
      bandOn(0, { leftPt: 300, rightPt: 540, topPt: 0, bottomPt: 100 }),
    );

    expect(boxes[0]?.lines[0]?.topPt).toBe(100);
  });

  it("leaves a cell's text alone, since a cell is measured from its own origin", () => {
    const boxes = wrapped(table(cell(paragraph(``, "aaaa"))), () => [
      { leftPt: 0, rightPt: 530, topPt: 0, bottomPt: 100 },
    ]);

    expect(boxes[0]?.lines[0]?.topPt).toBe(36);
  });
});

const SPACED = `<w:spacing w:before="240" w:after="240"/><w:contextualSpacing/>`;

const OTHER_STYLE = NORMAL.replace(
  "</w:styles>",
  `<w:style w:type="paragraph" w:styleId="Other">
     <w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr></w:style></w:styles>`,
);

describe("measureStack where a paragraph asks for contextual spacing", () => {
  it("drops the space between two paragraphs of the same style", () => {
    const boxes = boxesOf(paragraph(SPACED, "aaaa") + paragraph(SPACED, "bbbb"));

    expect(boxes[0]?.heightPt).toBeCloseTo(12 + ARIAL_12, 9);
    expect(boxes[1]?.topPt).toBeCloseTo(36 + 12 + ARIAL_12, 9);
    expect(boxes[1]?.lines[0]?.topPt).toBeCloseTo(36 + 12 + ARIAL_12, 9);
    expect(boxes[1]?.heightPt).toBeCloseTo(ARIAL_12 + 12, 9);
  });

  it("keeps the space against a paragraph set in another style", () => {
    const body =
      paragraph(SPACED, "aaaa") + paragraph(`<w:pStyle w:val="Other"/>${SPACED}`, "bbbb");
    const boxes = boxesOf(body, OTHER_STYLE);

    // The room between the two is kept rather than dropped, and it is the larger
    // of what either asks for rather than both of them.
    expect(boxes[0]?.heightPt).toBeCloseTo(12 + ARIAL_12 + 12, 9);
    expect(boxes[1]?.lines[0]?.topPt).toBeCloseTo(36 + 12 + ARIAL_12 + 12, 9);
  });

  it("keeps the space where the paragraph beside it is in another cell", () => {
    const spaced = paragraph(SPACED, "aaaa");
    const boxes = boxesOf(
      `<w:tbl><w:tr>${cell(spaced)}</w:tr><w:tr>${cell(spaced)}</w:tr></w:tbl>`,
    );

    expect(boxes[0]?.heightPt).toBeCloseTo(12 + ARIAL_12 + 12, 9);
    expect(boxes[1]?.heightPt).toBeCloseTo(12 + ARIAL_12 + 12, 9);
  });
});

const cell = (inner: string, properties = "") =>
  `<w:tc><w:tcPr>${properties}</w:tcPr>${inner}</w:tc>`;
const table = (...cells: readonly string[]) => `<w:tbl><w:tr>${cells.join("")}</w:tr></w:tbl>`;

// A line drawn round a cell, and the margins the table holds its text off its
// walls by, which are what a border has to be told from.
const lined = (eighths: number) =>
  `<w:tcBorders><w:top w:val="single" w:sz="${String(eighths)}" w:color="FF0000"/>` +
  `<w:bottom w:val="single" w:sz="${String(eighths)}" w:color="FF0000"/></w:tcBorders>`;

const walls = (twips: number) =>
  `<w:tblPr><w:tblCellMar><w:top w:w="${String(twips)}" w:type="dxa"/>` +
  `<w:bottom w:w="${String(twips)}" w:type="dxa"/></w:tblCellMar></w:tblPr>`;

const linedTable = (eighths: number, marginTwips = 0) =>
  `<w:tbl>${walls(marginTwips)}<w:tr>${cell(`<w:p/>`, lined(eighths))}</w:tr></w:tbl>`;

// Word's own answers, measured by the authored `positioned-table` document. The
// frame here runs from 72 to 540 and the stack starts at 36, so a table an inch off
// the column stands at 144 and one the same inch off the sheet at 72.
describe("measureStack over a table taken out of the flow", () => {
  const positioned = (properties: string, ...cells: readonly string[]) =>
    `<w:tbl><w:tblPr><w:tblpPr ${properties}/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>` +
    `<w:tr>${cells.join("")}</w:tr></w:tbl>`;

  const across = `w:vertAnchor="text" w:tblpX="1440"`;

  it("leaves the flow no room where the table stood", () => {
    const result = measure(`<w:p/>${positioned(across, cell(`<w:p/>`))}<w:p/>`);
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    const after = result.boxes[result.boxes.length - 1];
    expect(after?.topPt).toBeCloseTo(36 + ARIAL_12, 9);
    expect(result.heightPt).toBeCloseTo(ARIAL_12 * 2, 9);
  });

  it("measures the table from where the flow stood", () => {
    const boxes = boxesOf(`<w:p/>${positioned(across, cell(`<w:p/>`))}<w:p/>`);
    expect(boxes[1]?.topPt).toBeCloseTo(36 + ARIAL_12, 9);
  });

  it("drops it by the offset it asks for down the page", () => {
    const boxes = boxesOf(
      `<w:p/>${positioned(`w:vertAnchor="text" w:tblpX="1440" w:tblpY="-360"`, cell(`<w:p/>`))}<w:p/>`,
    );
    expect(boxes[1]?.topPt).toBeCloseTo(36 + ARIAL_12 - 18, 9);
  });

  it("stands it off the column where the anchor names the column", () => {
    const boxes = boxesOf(positioned(across, cell(`<w:p><w:r><w:t>aaaa</w:t></w:r></w:p>`)));
    expect(boxes[0]?.lines[0]?.leftPt).toBeCloseTo(72 + 72, 9);
  });

  it("stands it off the sheet where the anchor names the page", () => {
    const boxes = boxesOf(
      positioned(
        `w:vertAnchor="text" w:horzAnchor="page" w:tblpX="1440"`,
        cell(`<w:p><w:r><w:t>aaaa</w:t></w:r></w:p>`),
      ),
    );
    expect(boxes[0]?.lines[0]?.leftPt).toBeCloseTo(72, 9);
  });

  it("puts its right edge on the frame's where it asks for the right", () => {
    const boxes = boxesOf(
      positioned(
        `w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="right"`,
        cell(`<w:p><w:r><w:t>aaaa</w:t></w:r></w:p>`),
      ),
    );
    expect(boxes[0]?.lines[0]?.leftPt).toBeCloseTo(72 + 468 - 144, 9);
  });

  it("keeps the text it left clear of it by the distance it asks for", () => {
    const body =
      positioned(
        `w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:tblpX="0"`,
        cell(`<w:p/>`),
      ) + paragraph("", "aaaa");
    const boxes = boxesOf(body);
    expect(boxes[boxes.length - 1]?.lines[0]?.leftPt).toBeCloseTo(72 + 144 + 9, 9);
  });
});

describe("measureStack over a table's own lines", () => {
  it("leaves each row half of every line drawn along its edges", () => {
    const result = measure(linedTable(48));
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    // Six points of line, half of it inside the row at each edge and half outside.
    expect(result.boxes[0]?.topPt).toBeCloseTo(36 + 6, 9);
    expect(result.heightPt).toBeCloseTo(6 + ARIAL_12 + 6, 9);
  });

  // The margin holds the text off the wall and the line stands inside that, so a
  // row lined either side is its margins, its text and the whole of both lines.
  it("adds a line to the margin the cell already asks for rather than the larger of the two", () => {
    const result = measure(linedTable(2, 144));
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.heightPt).toBeCloseTo(7.2 + ARIAL_12 + 7.2 + 0.25 + 0.25, 9);
  });

  it("hands the page a rectangle for every cell, whatever its paragraphs did", () => {
    const result = measure(table(cell(`<w:p/><w:p/>`, lined(8)), cell(`<w:p/>`)));
    if (result.kind !== "measured") throw new Error(result.blocker.kind);
    expect(result.cells.map((each) => [each.topPt, each.heightPt])).toStrictEqual([
      [36.5, ARIAL_12 * 2 + 1],
      [36.5, ARIAL_12 * 2 + 1],
    ]);
    expect(result.cells[0]?.borders.top?.widthPt).toBe(1);
    expect(result.cells[1]?.borders.top).toBeNull();
  });
});

describe("measureStack over a paragraph's own lines", () => {
  const boxed = (properties: string, space = 0) =>
    paragraph(
      `<w:pBdr><w:top w:val="single" w:sz="12" w:space="${String(space)}" w:color="FF0000"/>` +
        `<w:bottom w:val="single" w:sz="12" w:space="${String(space)}" w:color="FF0000"/></w:pBdr>${properties}`,
    );

  it("takes the room the lines round it need out of the flow", () => {
    const boxes = boxesOf(boxed("", 6));
    expect(boxes[0]?.heightPt).toBeCloseTo(7.5 + ARIAL_12 + 7.5, 9);
    expect(boxes[0]?.lines[0]?.topPt).toBeCloseTo(36 + 7.5, 9);
  });

  it("joins a run of paragraphs asking for the same box into one", () => {
    const boxes = boxesOf(`${boxed("")}${boxed("")}`);
    expect(boxes[0]?.paint?.borders.bottom).toBeNull();
    expect(boxes[1]?.paint?.borders.top).toBeNull();
    expect(boxes[0]?.paint?.borders.top?.widthPt).toBe(1.5);
    // The line between them is not drawn, so neither is any room left for it.
    expect(boxes[1]?.topPt).toBeCloseTo(36 + 1.5 + ARIAL_12, 9);
  });

  it("reaches across the text area rather than across what the text filled", () => {
    const boxes = boxesOf(boxed(`<w:ind w:left="720" w:right="1440"/>`));
    expect([boxes[0]?.paint?.leftPt, boxes[0]?.paint?.rightPt]).toStrictEqual([72 + 36, 540 - 72]);
  });

  it("leaves a paragraph that asks for neither line nor colour nothing to draw", () => {
    expect(boxesOf(paragraph(""))[0]?.paint).toBeNull();
  });
});
