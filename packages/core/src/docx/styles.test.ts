import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { openDocx } from "./package.js";
import {
  readStyleTable,
  resolveParagraphFrame,
  resolveParagraphMark,
  resolveParagraphNumbering,
  resolveRuns,
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

const runMark = (rPr: string) =>
  markOf(`<w:p><w:pPr><w:rPr>${rPr}</w:rPr></w:pPr></w:p>`, styles(NORMAL));

describe("resolveParagraphMark", () => {
  // Word's own answer, measured on 2026-08-13 off the room a footer takes out of the
  // body: a footer of one paragraph stating a size of nought costs exactly what one
  // of a single half-point paragraph costs, and half what a whole point costs. So
  // nought is held to the smallest size the attribute can spell rather than ignored
  // and inherited.
  it("holds a run stating no size at all to half a point rather than inheriting one", () => {
    expect(runMark(`<w:sz w:val="0"/>`).fontSizePt).toBe(0.5);
  });

  it("leaves a run of half a point alone, which is the smallest one may be", () => {
    expect(runMark(`<w:sz w:val="1"/>`).fontSizePt).toBe(0.5);
  });

  it("leaves a run of a point alone", () => {
    expect(runMark(`<w:sz w:val="2"/>`).fontSizePt).toBe(1);
  });

  it("leaves every ordinary size alone", () => {
    expect(runMark(`<w:sz w:val="24"/>`).fontSizePt).toBe(12);
  });

  it("falls back to the default paragraph style when the paragraph names none", () => {
    expect(markOf(`<w:p/>`, styles(NORMAL))).toStrictEqual({
      font: { kind: "named", name: "Arial" },
      fontSizePt: 11,
      bold: false,
      italic: false,
      underline: false,
      raisePt: 0,
      lineSizePt: 11,
      lineRaisePt: 0,
      color: null,
      characterSpacingPt: 0,
      characterScale: 1,
      kernFromHalfPoints: null,
      highlight: null,
    });
  });

  it("lets the paragraph mark override the style it inherits from", () => {
    const body = `<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr></w:p>`;
    expect(markOf(body, styles(NORMAL))).toStrictEqual({
      font: { kind: "named", name: "Arial" },
      fontSizePt: 14,
      bold: false,
      italic: false,
      underline: false,
      raisePt: 0,
      lineSizePt: 14,
      lineRaisePt: 0,
      color: null,
      characterSpacingPt: 0,
      characterScale: 1,
      kernFromHalfPoints: null,
      highlight: null,
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
      underline: false,
      raisePt: 0,
      lineSizePt: 20,
      lineRaisePt: 0,
      color: null,
      characterSpacingPt: 0,
      characterScale: 1,
      kernFromHalfPoints: null,
      highlight: null,
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
      underline: false,
      raisePt: 0,
      lineSizePt: 10,
      lineRaisePt: 0,
      color: null,
      characterSpacingPt: 0,
      characterScale: 1,
      kernFromHalfPoints: null,
      highlight: null,
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
      underline: false,
      raisePt: 0,
      lineSizePt: 12,
      lineRaisePt: 0,
      color: null,
      characterSpacingPt: 0,
      characterScale: 1,
      kernFromHalfPoints: null,
      highlight: null,
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

  // **What `w:w` says, which 73 of the 718 corpus documents state on 3683 runs.**
  // Measured against Word on 2026-08-14 over one line of the same word repeated: a
  // run scaled to 103 came out 102.84% as wide as the plain one, to 90 89.90%, to
  // 150 149.90% and to 50 49.89%, each a tenth of a percent short of its own
  // multiple where Word rounds a scaled advance.
  it("reads the scale a run states, and one where it states none", () => {
    expect(runMark(`<w:w w:val="103"/>`).characterScale).toBeCloseTo(1.03, 9);
    expect(runMark(`<w:w w:val="50"/>`).characterScale).toBe(0.5);
    expect(runMark(``).characterScale).toBe(1);
  });

  // Inside `w:pPr` the same name is the width of a table's own column, and inside
  // `w:sectPr` the width of the page: only a run's own properties state a scale.
  it("takes no scale from a value it cannot read", () => {
    expect(runMark(`<w:w w:val="nought"/>`).characterScale).toBe(1);
    expect(runMark(`<w:w w:val="0"/>`).characterScale).toBe(1);
  });

  // Word sets a superscript smaller than the run it belongs to and lifts it off
  // the baseline; the size the file declares is not the size it draws.
  it("shrinks a superscript and lifts it off the baseline", () => {
    const mark = runMark(`<w:sz w:val="28"/><w:vertAlign w:val="superscript"/>`);

    expect(mark.fontSizePt).toBeCloseTo(14 * 0.65, 9);
    expect(mark.raisePt).toBeCloseTo(14 / 3, 9);
  });

  // Not by as much: measured on 2026-08-07 by the authored `raised-text` document
  // at three sizes, Word lifts a superscript a third of the size and drops a
  // subscript a tenth of it. 12pt, 24pt and 36pt went up 4.08, 7.92 and 12.00 and
  // down 0.96, 2.40 and 3.60.
  it("drops a subscript below the baseline by a tenth of its size", () => {
    const mark = runMark(`<w:sz w:val="28"/><w:vertAlign w:val="subscript"/>`);

    expect(mark.fontSizePt).toBeCloseTo(14 * 0.65, 9);
    expect(mark.raisePt).toBeCloseTo(-14 / 10, 9);
  });

  // Neither script changes the line: it is measured at the size the run was
  // declared at whatever the script shrank it to, which is why a 24pt superscript
  // and a 24pt subscript beside the same 12pt text made the same line.
  it("leaves a script's line the size the run was declared at", () => {
    for (const align of ["superscript", "subscript"]) {
      const mark = runMark(`<w:sz w:val="48"/><w:vertAlign w:val="${align}"/>`);

      expect(mark.lineSizePt).toBe(24);
      expect(mark.lineRaisePt).toBe(0);
    }
  });

  // A raise is stated in half-points and is a distance rather than a share of the
  // size: measured on 2026-08-07 by the authored `raised-text` document, where a
  // 12pt run stating twelve was drawn exactly six points up.
  it("raises a run by half the half-points it states, whatever size it is set in", () => {
    expect(runMark(`<w:sz w:val="24"/><w:position w:val="12"/>`).raisePt).toBeCloseTo(6, 9);
    expect(runMark(`<w:sz w:val="48"/><w:position w:val="12"/>`).raisePt).toBeCloseTo(6, 9);
    expect(runMark(`<w:sz w:val="24"/><w:position w:val="-1"/>`).raisePt).toBeCloseTo(-0.5, 9);
    expect(runMark(`<w:sz w:val="24"/><w:position w:val="0"/>`).raisePt).toBe(0);
  });

  // The two raises add: a 12pt superscript stating twelve was drawn 10.08pt off the
  // baseline, which is Word's own 4.08 for the script and the 6 the run asked for.
  it("adds a stated raise to the one a superscript already has", () => {
    const mark = runMark(
      `<w:sz w:val="24"/><w:vertAlign w:val="superscript"/><w:position w:val="12"/>`,
    );

    expect(mark.fontSizePt).toBeCloseTo(12 * 0.65, 9);
    expect(mark.raisePt).toBeCloseTo(12 / 3 + 6, 9);
  });

  it("leaves a run asked for at the baseline the size it was declared at", () => {
    const mark = runMark(`<w:sz w:val="28"/><w:vertAlign w:val="baseline"/>`);

    expect(mark.fontSizePt).toBe(14);
    expect(mark.raisePt).toBe(0);
  });

  it("reports an unresolved font rather than guessing one", () => {
    expect(markOf(`<w:p/>`, styles(""))).toStrictEqual({
      font: { kind: "unresolved" },
      fontSizePt: 10,
      bold: false,
      italic: false,
      underline: false,
      raisePt: 0,
      lineSizePt: 10,
      lineRaisePt: 0,
      color: null,
      characterSpacingPt: 0,
      characterScale: 1,
      kernFromHalfPoints: null,
      highlight: null,
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
  // **Word centres an equation with a paragraph to itself and ignores the
  // paragraph's own w:jc.** Measured over five cases against Word's own pdf.
  const MATH = `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"`;
  const equation = `<m:oMath ${MATH}><m:r><m:t>x</m:t></m:r></m:oMath>`;

  it("centres a paragraph holding an equation and nothing else", () => {
    expect(resolved(`<w:p>${equation}</w:p>`).frame.alignment).toBe("center");
  });

  it("centres it however the paragraph is aligned", () => {
    const body = `<w:p><w:pPr><w:jc w:val="right"/></w:pPr>${equation}</w:p>`;
    expect(resolved(body).frame.alignment).toBe("center");
  });

  it("lets the equation's own m:jc move it", () => {
    const body =
      `<w:p><m:oMathPara ${MATH}><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>` +
      `<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara></w:p>`;
    expect(resolved(body).frame.alignment).toBe("left");
  });

  it("leaves a paragraph holding text beside an equation as it is aligned", () => {
    const body = `<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>a</w:t></w:r>${equation}</w:p>`;
    expect(resolved(body).frame.alignment).toBe("right");
  });

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

// Word's own answer, measured on 2026-08-13 in Times New Roman, twelve cases three
// times each: baselines 13.8 apart go to 28.1 with an automatic space above and to
// 27.6 with one below, and a paragraph at twice the size gets the same fourteen
// points on top of its taller line. So it is fourteen points, it does not follow
// the face, and it wins over a value stated beside it.
describe("the space a paragraph asks for automatically", () => {
  const spacing = (attributes: string, stylesXml?: string) =>
    resolved(`<w:p><w:pPr><w:spacing ${attributes}/></w:pPr></w:p>`, stylesXml).frame;

  it("gives fourteen points above and below", () => {
    const frame = spacing(`w:beforeAutospacing="1" w:afterAutospacing="1"`);
    expect(frame.spaceBeforeTwips).toBe(280);
    expect(frame.spaceAfterTwips).toBe(280);
  });

  it("takes each side on its own", () => {
    expect(spacing(`w:beforeAutospacing="1"`).spaceAfterTwips).toBe(0);
    expect(spacing(`w:afterAutospacing="1"`).spaceBeforeTwips).toBe(0);
  });

  it("wins over a value stated beside it", () => {
    expect(spacing(`w:before="240" w:beforeAutospacing="1"`).spaceBeforeTwips).toBe(280);
  });

  // Six hundred paragraphs of the corpus turn the automatic space off while the
  // style they name turns it on, so the answer has to be the value under it rather
  // than fourteen points.
  it("leaves a paragraph turning it off with the value the cascade states", () => {
    const web = styles(
      `${NORMAL}<w:style w:type="paragraph" w:styleId="Web">
         <w:pPr><w:spacing w:before="120" w:beforeAutospacing="1"/></w:pPr></w:style>`,
    );
    const styled = (attributes: string) =>
      resolved(`<w:p><w:pPr><w:pStyle w:val="Web"/><w:spacing ${attributes}/></w:pPr></w:p>`, web)
        .frame;
    expect(styled(`w:beforeAutospacing="1"`).spaceBeforeTwips).toBe(280);
    expect(styled(`w:beforeAutospacing="0"`).spaceBeforeTwips).toBe(120);
    expect(styled(`w:before="60" w:beforeAutospacing="0"`).spaceBeforeTwips).toBe(60);
  });

  it("reads the style's own automatic space where the paragraph states none", () => {
    const web = styles(
      `${NORMAL}<w:style w:type="paragraph" w:styleId="Web">
         <w:pPr><w:spacing w:afterAutospacing="1"/></w:pPr></w:style>`,
    );
    expect(
      resolved(`<w:p><w:pPr><w:pStyle w:val="Web"/></w:pPr></w:p>`, web).frame.spaceAfterTwips,
    ).toBe(280);
  });
});

// A link takes its underline from the character style it is given, and the one
// run that turns its own off is what the cascade has to let through.
describe("the underline a run carries", () => {
  const HYPERLINK = `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>
    <w:rPr><w:u w:val="single"/></w:rPr></w:style>`;

  const underlineOf = (rPr: string) => {
    const body = `<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>a link</w:t></w:r></w:p>`;
    const pkg = openDocx(
      buildDocx({
        "word/document.xml": wordDocument(body),
        "word/styles.xml": styles(NORMAL + HYPERLINK),
      }),
    );
    const paragraph = readParagraphs(pkg)[0];
    if (paragraph === undefined) throw new Error("expected a paragraph");
    return resolveRuns(paragraph, readStyleTable(pkg))[0]?.mark.underline;
  };

  it("draws no underline where none is asked for", () => {
    expect(underlineOf(``)).toBe(false);
  });

  it("takes any named kind of underline as one to draw", () => {
    expect(underlineOf(`<w:u w:val="single"/>`)).toBe(true);
    expect(underlineOf(`<w:u w:val="dotted"/>`)).toBe(true);
    expect(underlineOf(`<w:u/>`)).toBe(true);
  });

  it("reads a run that turns its underline off, which no toggle would", () => {
    expect(underlineOf(`<w:rStyle w:val="Hyperlink"/>`)).toBe(true);
    expect(underlineOf(`<w:rStyle w:val="Hyperlink"/><w:u w:val="none"/>`)).toBe(false);
  });
});

// A paragraph inside a table reads the table's own style between the document's
// defaults and its own. Where a document leaves `Normal` empty and states its
// spacing in `docDefaults`, which is what Word writes, the table style is the only
// thing between the two and it decides the height of every row.
describe("a paragraph inside a table", () => {
  const TABLE_STYLE = `<w:style w:type="table" w:styleId="TableGrid">
      <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    </w:style>`;
  const DEFAULTS = `<w:docDefaults><w:pPrDefault><w:pPr>
      <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault></w:docDefaults>`;

  const frameIn = (tableStyleId: string | null) => {
    const pkg = openDocx(
      buildDocx({
        "word/document.xml": wordDocument(`<w:p><w:r><w:t>a</w:t></w:r></w:p>`),
        "word/styles.xml": `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">${DEFAULTS}${TABLE_STYLE}</w:styles>`,
      }),
    );
    const styles = readStyleTable(pkg);
    const paragraph = readParagraphs(pkg)[0];
    if (paragraph === undefined) throw new Error("no paragraph");
    return resolveParagraphFrame(
      paragraph,
      styles,
      tableStyleId === null ? null : { styleId: tableStyleId, at: null },
    );
  };

  it("keeps the document's own defaults outside a table", () => {
    expect(frameIn(null).spaceAfterTwips).toBe(160);
    expect(frameIn(null).lineTwips).toBe(259);
  });

  it("takes the table style's spacing over the document's defaults inside one", () => {
    expect(frameIn("TableGrid").spaceAfterTwips).toBe(0);
    expect(frameIn("TableGrid").lineTwips).toBe(240);
  });
});
