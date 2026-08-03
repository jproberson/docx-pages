import { describe, expect, it } from "vitest";

import { readBlocks } from "../docx/blocks.js";
import { openDocx } from "../docx/package.js";
import { readStyleTable } from "../docx/styles.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { lookupFontMetrics } from "./font-metrics.js";
import { measureStack } from "./stack.js";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

const NORMAL = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr></w:style></w:styles>`;

const measure = (body: string, stylesXml: string = NORMAL) => {
  const pkg = openDocx(
    buildDocx({ "word/document.xml": wordDocument(body), "word/styles.xml": stylesXml }),
  );
  return measureStack({
    blocks: readBlocks(pkg),
    styles: readStyleTable(pkg),
    metricsFor: (name) => lookupFontMetrics(name),
    part: "word/document.xml",
    originPt: 36,
  });
};

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
