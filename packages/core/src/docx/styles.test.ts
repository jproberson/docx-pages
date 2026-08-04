import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { openDocx } from "./package.js";
import { readStyleTable, resolveParagraphMark } from "./styles.js";
import { readParagraphs } from "./blocks.js";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

const theme = (major: string, minor: string) => `<?xml version="1.0"?>
<a:theme xmlns:a="${A_NS}"><a:themeElements><a:fontScheme>
  <a:majorFont><a:latin typeface="${major}"/></a:majorFont>
  <a:minorFont><a:latin typeface="${minor}"/></a:minorFont>
</a:fontScheme></a:themeElements></a:theme>`;

const styles = (inner: string) => `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${inner}</w:styles>`;

const markOf = (body: string, stylesXml?: string, themeXml?: string) => {
  const parts: Record<string, string> = { "word/document.xml": wordDocument(body) };
  if (stylesXml !== undefined) parts["word/styles.xml"] = stylesXml;
  if (themeXml !== undefined) parts["word/theme/theme1.xml"] = themeXml;
  const pkg = openDocx(buildDocx(parts));
  const paragraph = readParagraphs(pkg)[0];
  if (paragraph === undefined) throw new Error("expected a paragraph");
  return resolveParagraphMark(paragraph, readStyleTable(pkg));
};

const NORMAL = `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
  <w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="22"/></w:rPr></w:style>`;

describe("resolveParagraphMark", () => {
  it("falls back to the default paragraph style when the paragraph names none", () => {
    expect(markOf(`<w:p/>`, styles(NORMAL))).toStrictEqual({
      font: { kind: "named", name: "Arial" },
      fontSizePt: 11,
      bold: false,
      italic: false,
      color: null,
    });
  });

  it("lets the paragraph mark override the style it inherits from", () => {
    const body = `<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr></w:p>`;
    expect(markOf(body, styles(NORMAL))).toStrictEqual({
      font: { kind: "named", name: "Arial" },
      fontSizePt: 14,
      bold: false,
      italic: false,
      color: null,
    });
  });

  it("is the case that caught the spacer paragraphs: size overridden, font inherited", () => {
    const body = `<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Meridian Sans"/></w:rPr></w:pPr></w:p>`;
    expect(markOf(body, styles(NORMAL)).font).toStrictEqual({
      kind: "named",
      name: "Meridian Sans",
    });
  });

  it("walks the basedOn chain, with the nearer style winning", () => {
    const chain = styles(
      `${NORMAL}<w:style w:type="paragraph" w:styleId="Child">
         <w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="40"/></w:rPr></w:style>`,
    );
    const body = `<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr></w:p>`;
    expect(markOf(body, chain)).toStrictEqual({
      font: { kind: "named", name: "Arial" },
      fontSizePt: 20,
      bold: false,
      italic: false,
      color: null,
    });
  });

  it("takes docDefaults as the floor beneath every style", () => {
    const withDefaults = styles(
      `<w:docDefaults><w:rPrDefault><w:rPr>
         <w:rFonts w:ascii="Calibri"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>`,
    );
    expect(markOf(`<w:p/>`, withDefaults)).toStrictEqual({
      font: { kind: "named", name: "Calibri" },
      fontSizePt: 10,
      bold: false,
      italic: false,
      color: null,
    });
  });

  it("resolves a minor theme font reference through the theme part", () => {
    const themed = styles(
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
         <w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/><w:sz w:val="24"/></w:rPr></w:style>`,
    );
    expect(markOf(`<w:p/>`, themed, theme("Georgia", "Aptos"))).toStrictEqual({
      font: { kind: "named", name: "Aptos" },
      fontSizePt: 12,
      bold: false,
      italic: false,
      color: null,
    });
  });

  it("resolves a major theme font reference", () => {
    const themed = styles(
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
         <w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/></w:rPr></w:style>`,
    );
    expect(markOf(`<w:p/>`, themed, theme("Georgia", "Aptos")).font).toStrictEqual({
      kind: "named",
      name: "Georgia",
    });
  });

  it("reports an unresolved font rather than guessing one", () => {
    expect(markOf(`<w:p/>`, styles(""))).toStrictEqual({
      font: { kind: "unresolved" },
      fontSizePt: 10,
      bold: false,
      italic: false,
      color: null,
    });
  });

  it("uses hAnsi when ascii is absent", () => {
    const body = `<w:p><w:pPr><w:rPr><w:rFonts w:hAnsi="Verdana"/></w:rPr></w:pPr></w:p>`;
    expect(markOf(body, styles("")).font).toStrictEqual({ kind: "named", name: "Verdana" });
  });
});
