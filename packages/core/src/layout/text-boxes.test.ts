import { describe, expect, it } from "vitest";

import { readBlocks } from "../docx/blocks.js";
import { DEFAULT_TEXT_INSETS, type TextBoxAnchor, type TextBoxBody } from "../docx/drawing.js";
import { openDocx } from "../docx/package.js";
import { readStyleTable, type StyleTable } from "../docx/styles.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { buildFace } from "../testing/build-font.js";
import { lookupFontMetrics } from "./font-metrics.js";
import { layOutTextBox, type PlacedTextBox, type TextBoxRect } from "./text-boxes.js";

const NORMAL = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr></w:style></w:styles>`;

const ARIAL = buildFace({
  name: "Arial",
  metrics: { unitsPerEm: 2048, ascender: 1854, descender: -434, lineGap: 67 },
});

const ARIAL_12 = 13.798828125;

const RECT: TextBoxRect = { leftPt: 100, topPt: 200, widthPt: 180, heightPt: 90 };

const NO_INSETS = { leftEmu: 0, topEmu: 0, rightEmu: 0, bottomEmu: 0 };

type BodyOptions = {
  readonly anchor?: TextBoxAnchor;
  readonly wraps?: boolean;
  readonly insets?: TextBoxBody["insets"];
};

function bodyOf(inner: string, options: BodyOptions = {}): [TextBoxBody, StyleTable] {
  const pkg = openDocx(
    buildDocx({ "word/document.xml": wordDocument(inner), "word/styles.xml": NORMAL }),
  );
  return [
    {
      blocks: readBlocks(pkg),
      insets: options.insets ?? NO_INSETS,
      anchor: options.anchor ?? "top",
      wraps: options.wraps ?? true,
    },
    readStyleTable(pkg),
  ];
}

function place(inner: string, options: BodyOptions = {}, rect: TextBoxRect = RECT): PlacedTextBox {
  const [body, styles] = bodyOf(inner, options);
  const result = layOutTextBox({
    body,
    rect,
    styles,
    metricsFor: (request) => lookupFontMetrics(request, [ARIAL]),
    part: "word/document.xml",
  });
  if (result.kind !== "laid-out") throw new Error(result.blocker.kind);
  return result.text;
}

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

describe("layOutTextBox", () => {
  it("starts the text at the box's own top left corner", () => {
    const text = place(paragraph("aa"));

    expect(text.boxes[0]?.topPt).toBe(200);
    expect(text.boxes[0]?.lines[0]?.leftPt).toBe(100);
  });

  it("holds the text inside the insets the shape states", () => {
    const text = place(paragraph("aa"), { insets: DEFAULT_TEXT_INSETS });

    expect(text.boxes[0]?.lines[0]?.leftPt).toBeCloseTo(107.2, 9);
    expect(text.boxes[0]?.topPt).toBeCloseTo(203.6, 9);
  });

  it("breaks the text at the box's width", () => {
    const text = place(paragraph("aaaa bbbb"), {}, { ...RECT, widthPt: 30 });
    expect(text.boxes[0]?.lines).toHaveLength(2);
  });

  it("lets a box that refuses to wrap run its text past its own edge", () => {
    const text = place(paragraph("aaaa bbbb"), { wraps: false }, { ...RECT, widthPt: 30 });
    expect(text.boxes[0]?.lines).toHaveLength(1);
  });

  it("reports how tall the text came out, whatever the frame's height", () => {
    expect(place(`${paragraph("aa")}${paragraph("bb")}`).contentHeightPt).toBeCloseTo(
      ARIAL_12 * 2,
      9,
    );
  });

  it("centres the text in a box taller than it when the shape asks", () => {
    const text = place(paragraph("aa"), { anchor: "center" });
    expect(text.boxes[0]?.topPt).toBeCloseTo(200 + (90 - ARIAL_12) / 2, 9);
  });

  it("seats the text against the bottom when the shape asks", () => {
    const text = place(paragraph("aa"), { anchor: "bottom" });
    expect(text.boxes[0]?.topPt).toBeCloseTo(200 + 90 - ARIAL_12, 9);
  });

  it("moves each line with the text it belongs to", () => {
    const text = place(paragraph("aa"), { anchor: "bottom" });
    const line = text.boxes[0]?.lines[0];

    expect(line?.topPt).toBeCloseTo(200 + 90 - ARIAL_12, 9);
    expect(line?.baselinePt).toBeCloseTo(200 + 90 - ARIAL_12 + (12 * 1854) / 2048, 9);
  });

  it("lets text overflow a box too short for it rather than clipping it away", () => {
    const short = { ...RECT, heightPt: 5 };
    const text = place(`${paragraph("aa")}${paragraph("bb")}`, { anchor: "bottom" }, short);

    expect(text.boxes[0]?.topPt).toBeLessThan(200);
  });

  it("reports the paragraph it could not measure rather than laying out around it", () => {
    const [body, styles] = bodyOf(paragraph("aa"));
    const result = layOutTextBox({
      body,
      rect: RECT,
      styles,
      metricsFor: (request) => lookupFontMetrics(request),
      part: "word/document.xml",
    });

    expect(result.kind === "blocked" && result.blocker.kind).toBe("unmeasurable-text");
  });

  it("lays out a box holding nothing as empty", () => {
    expect(place(``).boxes).toStrictEqual([]);
  });
});
