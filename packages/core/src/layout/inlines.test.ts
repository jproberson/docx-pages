import { describe, expect, it } from "vitest";

import { readParagraphs } from "../docx/blocks.js";
import { readInlines } from "../docx/inlines.js";
import { openDocx } from "../docx/package.js";
import type { SectionGeometry } from "../docx/section.js";
import { NO_THEME } from "../docx/theme.js";
import { readStyleTable, resolveParagraphFrame } from "../docx/styles.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { placeInlines } from "./inlines.js";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

const LETTER: SectionGeometry = {
  widthTwips: 12240,
  heightTwips: 15840,
  margin: {
    topTwips: 720,
    rightTwips: 720,
    bottomTwips: 0,
    leftTwips: 720,
    headerTwips: 432,
    footerTwips: 144,
  },
};

// 180pt by 90pt.
const image = (cx = 2286000, cy = 1143000) =>
  `<w:r><w:drawing><wp:inline xmlns:wp="${WP_NS}">
     <wp:extent cx="${String(cx)}" cy="${String(cy)}"/>
     <wp:docPr id="1" name="Logo"/></wp:inline></w:drawing></w:r>`;

const paragraph = (properties: string, runs: string) =>
  `<w:p>${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}${runs}</w:p>`;

const place = (body: string, paragraphTopPt = 100) => {
  const pkg = openDocx(buildDocx({ "word/document.xml": wordDocument(body) }));
  const found = readParagraphs(pkg)[0];
  if (found === undefined) throw new Error("expected a paragraph");
  return placeInlines({
    drawings: readInlines(found),
    page: LETTER,
    frame: resolveParagraphFrame(found, readStyleTable(pkg)),
    paragraphTopPt,
    resolvePart: () => null,
    theme: NO_THEME,
  });
};

describe("placeInlines", () => {
  it("starts an unaligned drawing at the left margin", () => {
    const [placed] = place(paragraph("", image()));
    expect(placed?.leftPt).toBeCloseTo(36, 6);
    expect(placed?.widthPt).toBeCloseTo(180, 6);
  });

  it("seats the drawing at its paragraph's top", () => {
    const [placed] = place(paragraph("", image()), 253.87);
    expect(placed?.topPt).toBeCloseTo(253.87, 6);
    expect(placed?.heightPt).toBeCloseTo(90, 6);
  });

  it("ends a right-aligned drawing at the right margin", () => {
    const [placed] = place(paragraph(`<w:jc w:val="right"/>`, image()));
    expect(placed?.leftPt).toBeCloseTo(612 - 36 - 180, 6);
  });

  it("centres between the margins", () => {
    const [placed] = place(paragraph(`<w:jc w:val="center"/>`, image()));
    expect(placed?.leftPt).toBeCloseTo(36 + (540 - 180) / 2, 6);
  });

  it("treats a justified paragraph as starting at the left, like a lone drawing does", () => {
    const [placed] = place(paragraph(`<w:jc w:val="both"/>`, image()));
    expect(placed?.leftPt).toBeCloseTo(36, 6);
  });

  it("narrows the line by the paragraph's indents", () => {
    const indented = `<w:jc w:val="right"/><w:ind w:left="720" w:right="1440"/>`;
    const [placed] = place(paragraph(indented, image()));
    expect(placed?.leftPt).toBeCloseTo(612 - 36 - 72 - 180, 6);
  });

  it("runs several drawings along the line rather than stacking them", () => {
    const placed = place(paragraph("", `${image()}${image(1143000, 1143000)}`));
    expect(placed.map((each) => each.leftPt)).toStrictEqual([36, 216]);
  });

  it("aligns a run of drawings by their combined width", () => {
    const placed = place(paragraph(`<w:jc w:val="right"/>`, `${image()}${image()}`));
    expect(placed[0]?.leftPt).toBeCloseTo(612 - 36 - 360, 6);
  });

  it("takes the alignment a paragraph style sets when the paragraph sets none", () => {
    const styles = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Figure"><w:pPr><w:jc w:val="right"/></w:pPr></w:style>
      </w:styles>`;
    const body = paragraph(`<w:pStyle w:val="Figure"/>`, image());
    const pkg = openDocx(
      buildDocx({ "word/document.xml": wordDocument(body), "word/styles.xml": styles }),
    );
    const found = readParagraphs(pkg)[0];
    if (found === undefined) throw new Error("expected a paragraph");
    expect(resolveParagraphFrame(found, readStyleTable(pkg)).alignment).toBe("right");
  });
});
