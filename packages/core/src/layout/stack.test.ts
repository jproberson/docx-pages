import { describe, expect, it } from "vitest";

import { readBlocks } from "../docx/blocks.js";
import { openDocx } from "../docx/package.js";
import { readStyleTable } from "../docx/styles.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { buildFace } from "../testing/build-font.js";
import { lookupFontMetrics } from "./font-metrics.js";
import { measureStack, type ParagraphBox } from "./stack.js";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

const NORMAL = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr></w:style></w:styles>`;

const measure = (body: string, stylesXml: string = NORMAL, widthPt = 468) => {
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
  });
};

function boxesOf(body: string, stylesXml: string = NORMAL, widthPt = 468) {
  const result = measure(body, stylesXml, widthPt);
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
    expect(result).toStrictEqual({ kind: "measured", boxes: [], heightPt: 0 });
  });
});

// Every glyph in the test face is half an em, so 12pt text is 6pt a character.
const PER_CHARACTER = 6;
const ARIAL_ASCENT_12 = (12 * 1854) / 2048;

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

  it("replaces a line's height when the rule is exact", () => {
    const box = firstBox(paragraph(`<w:spacing w:line="400" w:lineRule="exact"/>`, "aaaa"));
    expect(box.heightPt).toBeCloseTo(20, 9);
  });

  it("takes a line's own height when the rule only sets a floor under it", () => {
    const box = firstBox(paragraph(`<w:spacing w:line="120" w:lineRule="atLeast"/>`, "aaaa"));
    expect(box.heightPt).toBeCloseTo(ARIAL_12, 9);
  });

  it("puts extra leading above the text, where Word puts it", () => {
    const box = firstBox(paragraph(`<w:spacing w:line="480" w:lineRule="auto"/>`, "aaaa"));
    expect(box.lines[0]?.baselinePt).toBeCloseTo(36 + ARIAL_12 + ARIAL_ASCENT_12, 9);
  });

  it("lets the paragraph mark raise the last line, since it sits on it", () => {
    const body = `<w:p><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr>
      <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>aaaa</w:t></w:r></w:p>`;

    expect(firstBox(body).heightPt).toBeCloseTo(ARIAL_12 * 2, 9);
  });

  it("keeps an empty paragraph at its mark's height", () => {
    expect(firstBox(`<w:p/>`).lines).toStrictEqual([]);
  });
});

const cell = (inner: string, properties = "") =>
  `<w:tc><w:tcPr>${properties}</w:tcPr>${inner}</w:tc>`;
const table = (...cells: readonly string[]) => `<w:tbl><w:tr>${cells.join("")}</w:tr></w:tbl>`;

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
