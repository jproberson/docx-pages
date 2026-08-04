import { describe, expect, it } from "vitest";

import { readParagraphs } from "../docx/blocks.js";
import { openDocx } from "../docx/package.js";
import { readStyleTable, resolveParagraphFrame, type ParagraphFrame } from "../docx/styles.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { nextTabStopPt, tabStopsPt } from "./tab-stops.js";

const styles = (inner: string) => `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${inner}</w:styles>`;

const frameOf = (body: string, stylesXml = styles("")): ParagraphFrame => {
  const pkg = openDocx(
    buildDocx({ "word/document.xml": wordDocument(body), "word/styles.xml": stylesXml }),
  );
  const paragraph = readParagraphs(pkg)[0];
  if (paragraph === undefined) throw new Error("expected a paragraph");
  return resolveParagraphFrame(paragraph, readStyleTable(pkg));
};

const paragraph = (properties: string) => `<w:p><w:pPr>${properties}</w:pPr></w:p>`;

describe("tabStopsPt", () => {
  it("reads the stops a paragraph declares, in order", () => {
    const body = paragraph(
      `<w:tabs><w:tab w:val="left" w:pos="2880"/><w:tab w:val="left" w:pos="1440"/></w:tabs>`,
    );
    expect(tabStopsPt(frameOf(body))).toStrictEqual([72, 144]);
  });

  it("adds the stops a paragraph declares to the ones its style did", () => {
    const withStops = styles(
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
         <w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"/></w:tabs></w:pPr></w:style>`,
    );
    const body = paragraph(`<w:tabs><w:tab w:val="left" w:pos="2880"/></w:tabs>`);
    expect(tabStopsPt(frameOf(body, withStops))).toStrictEqual([72, 144]);
  });

  it("drops a stop the paragraph clears", () => {
    const withStops = styles(
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
         <w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"/></w:tabs></w:pPr></w:style>`,
    );
    const body = paragraph(`<w:tabs><w:tab w:val="clear" w:pos="1440"/></w:tabs>`);
    expect(tabStopsPt(frameOf(body, withStops))).toStrictEqual([]);
  });

  it("puts an implicit stop at the left indent of a hanging paragraph", () => {
    const body = paragraph(`<w:ind w:left="720" w:hanging="360"/>`);
    expect(tabStopsPt(frameOf(body))).toStrictEqual([36]);
  });

  it("leaves a paragraph with a first line indent to its declared stops", () => {
    const body = paragraph(`<w:ind w:left="720" w:firstLine="360"/>`);
    expect(tabStopsPt(frameOf(body))).toStrictEqual([]);
  });
});

describe("nextTabStopPt", () => {
  it("advances to the first stop past where the text has reached", () => {
    expect(nextTabStopPt(20, [12, 40, 90])).toBe(40);
  });

  it("passes over a stop the text has already reached", () => {
    expect(nextTabStopPt(40, [12, 40, 90])).toBe(90);
  });

  it("falls back to the default stops past the last one declared", () => {
    expect(nextTabStopPt(100, [12, 40, 90])).toBe(108);
  });

  it("uses the default stops when the paragraph declares none", () => {
    expect(nextTabStopPt(0, [])).toBe(36);
    expect(nextTabStopPt(36, [])).toBe(72);
  });
});
