import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { openDocx } from "./package.js";
import {
  readStyleTable,
  resolveParagraphFrame,
  resolveParagraphMark,
  resolveParagraphNumbering,
} from "./styles.js";
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

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const numbering = `<?xml version="1.0"?>
<w:numbering xmlns:w="${W_NS}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="bullet"/><w:lvlText w:val="-"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

const resolved = (body: string, stylesXml: string = styles(NORMAL)) => {
  const pkg = openDocx(
    buildDocx({
      "word/document.xml": wordDocument(body),
      "word/styles.xml": stylesXml,
      "word/numbering.xml": numbering,
    }),
  );
  const paragraph = readParagraphs(pkg)[0];
  if (paragraph === undefined) throw new Error("expected a paragraph");
  const table = readStyleTable(pkg);
  return {
    frame: resolveParagraphFrame(paragraph, table),
    numbering: resolveParagraphNumbering(paragraph, table),
  };
};

const numbered = (properties = "") =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
     ${properties}</w:pPr></w:p>`;

describe("resolveParagraphNumbering", () => {
  it("resolves the level a paragraph's numPr names", () => {
    expect(resolved(numbered()).numbering).toMatchObject({
      numId: "1",
      ilvl: 0,
      level: { format: "bullet", text: "-" },
    });
  });

  it("defaults to the first level when the paragraph gives only a numId", () => {
    const body = `<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p>`;
    expect(resolved(body).numbering?.ilvl).toBe(0);
  });

  it("takes the numId from the style and the level from the paragraph", () => {
    const listStyle = styles(
      `${NORMAL}<w:style w:type="paragraph" w:styleId="List">
         <w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>`,
    );
    const body = `<w:p><w:pPr><w:pStyle w:val="List"/>
      <w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr></w:p>`;
    expect(resolved(body, listStyle).numbering).toMatchObject({
      ilvl: 1,
      level: { format: "decimal" },
    });
  });

  it("reads numbering a paragraph carries only through its style", () => {
    const listStyle = styles(
      `${NORMAL}<w:style w:type="paragraph" w:styleId="List">
         <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>`,
    );
    const body = `<w:p><w:pPr><w:pStyle w:val="List"/></w:pPr></w:p>`;
    expect(resolved(body, listStyle).numbering?.ilvl).toBe(1);
  });

  it("takes numId zero as the paragraph refusing the numbering it inherits", () => {
    const listStyle = styles(
      `${NORMAL}<w:style w:type="paragraph" w:styleId="List">
         <w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>`,
    );
    const body = `<w:p><w:pPr><w:pStyle w:val="List"/>
      <w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:p>`;
    expect(resolved(body, listStyle).numbering).toBeNull();
  });

  it("has no numbering for a paragraph that names none", () => {
    expect(resolved(`<w:p/>`).numbering).toBeNull();
  });

  it("has no numbering when the numId resolves to no level", () => {
    const body = `<w:p><w:pPr><w:numPr><w:numId w:val="9"/></w:numPr></w:pPr></w:p>`;
    expect(resolved(body).numbering).toBeNull();
  });
});

describe("resolveParagraphFrame", () => {
  it("indents a numbered paragraph the way its level does", () => {
    expect(resolved(numbered()).frame).toMatchObject({
      indentLeftTwips: 720,
      indentFirstLineTwips: -360,
    });
  });

  it("lets the paragraph's own indent override the level's", () => {
    const frame = resolved(numbered(`<w:ind w:left="1080"/>`)).frame;
    expect(frame.indentLeftTwips).toBe(1080);
    expect(frame.indentFirstLineTwips).toBe(-360);
  });

  it("lets the level's indent override the style's", () => {
    const indented = styles(
      `${NORMAL}<w:style w:type="paragraph" w:styleId="List">
         <w:pPr><w:ind w:left="2880"/></w:pPr></w:style>`,
    );
    const body = `<w:p><w:pPr><w:pStyle w:val="List"/>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:p>`;
    expect(resolved(body, indented).frame.indentLeftTwips).toBe(720);
  });

  it("leaves an unnumbered paragraph on the margin", () => {
    expect(resolved(`<w:p/>`).frame.indentLeftTwips).toBe(0);
  });
});
