import { describe, expect, it } from "vitest";

import { readParagraphs } from "../docx/blocks.js";
import { openDocx } from "../docx/package.js";
import { readStyleTable, resolveParagraphFrame, type ParagraphFrame } from "../docx/styles.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { nextTabStop, tabStopsPt, type TabStopPt } from "./tab-stops.js";

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

const left = (...positionsPt: readonly number[]): readonly TabStopPt[] =>
  positionsPt.map((positionPt) => ({ positionPt, alignment: "left" as const }));

describe("tabStopsPt", () => {
  it("reads the stops a paragraph declares, in order", () => {
    const body = paragraph(
      `<w:tabs><w:tab w:val="left" w:pos="2880"/><w:tab w:val="left" w:pos="1440"/></w:tabs>`,
    );
    expect(tabStopsPt(frameOf(body))).toStrictEqual(left(72, 144));
  });

  it("adds the stops a paragraph declares to the ones its style did", () => {
    const withStops = styles(
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
         <w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"/></w:tabs></w:pPr></w:style>`,
    );
    const body = paragraph(`<w:tabs><w:tab w:val="left" w:pos="2880"/></w:tabs>`);
    expect(tabStopsPt(frameOf(body, withStops))).toStrictEqual(left(72, 144));
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
    expect(tabStopsPt(frameOf(body))).toStrictEqual(left(36));
  });

  it("leaves a paragraph with a first line indent to its declared stops", () => {
    const body = paragraph(`<w:ind w:left="720" w:firstLine="360"/>`);
    expect(tabStopsPt(frameOf(body))).toStrictEqual([]);
  });
});

describe("nextTabStop", () => {
  const at = (fromPt: number, stops: readonly TabStopPt[]): number =>
    nextTabStop(fromPt, stops).positionPt;

  it("advances to the first stop past where the text has reached", () => {
    expect(at(20, left(12, 40, 90))).toBe(40);
  });

  it("passes over a stop the text has already reached", () => {
    expect(at(40, left(12, 40, 90))).toBe(90);
  });

  it("falls back to the default stops past the last one declared", () => {
    expect(at(100, left(12, 40, 90))).toBe(108);
  });

  it("uses the default stops when the paragraph declares none", () => {
    expect(at(0, [])).toBe(36);
    expect(at(36, [])).toBe(72);
  });

  it("carries what the stop it reached does with the text that follows", () => {
    const stops: readonly TabStopPt[] = [{ positionPt: 40, alignment: "right" }];
    expect(nextTabStop(20, stops)).toStrictEqual({ positionPt: 40, alignment: "right" });
    expect(nextTabStop(50, stops).alignment).toBe("left");
  });
});

describe("tabStopsPt and bars", () => {
  it("leaves a bar out, since no tab ever lands on one", () => {
    const body = paragraph(
      `<w:tabs><w:tab w:val="bar" w:pos="2880"/><w:tab w:val="right" w:pos="1440"/></w:tabs>`,
    );
    expect(tabStopsPt(frameOf(body))).toStrictEqual([{ positionPt: 72, alignment: "right" }]);
  });
});
