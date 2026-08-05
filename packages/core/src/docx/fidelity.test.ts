import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument, WORDPROCESSING_NS } from "../testing/build-docx.js";
import { readUnhonoured, withSubstitutedFaces, type Unhonoured } from "./fidelity.js";
import { openDocx } from "./package.js";

const SECTION = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const reportOf = (body: string, parts: Readonly<Record<string, string>> = {}) =>
  readUnhonoured(
    openDocx(buildDocx({ "word/document.xml": wordDocument(`${body}${SECTION}`), ...parts })),
  );

const kinds = (report: readonly Unhonoured[]): readonly string[] =>
  report.map((entry) => entry.kind);

describe("readUnhonoured", () => {
  it("says nothing about a document holding only what is read", () => {
    expect(reportOf(`<w:p><w:r><w:t>plain</w:t></w:r></w:p>`)).toStrictEqual([]);
  });

  it("names a run's kerning, its letter spacing and its capitals", () => {
    const marks = `<w:kern w:val="16"/><w:spacing w:val="20"/><w:caps/>`;
    expect(kinds(reportOf(`<w:p><w:r><w:rPr>${marks}</w:rPr><w:t>a</w:t></w:r></w:p>`))).toContain(
      "character-kerning",
    );
    expect(kinds(reportOf(`<w:p><w:r><w:rPr>${marks}</w:rPr><w:t>a</w:t></w:r></w:p>`))).toContain(
      "character-spacing",
    );
    expect(kinds(reportOf(`<w:p><w:r><w:rPr>${marks}</w:rPr><w:t>a</w:t></w:r></w:p>`))).toContain(
      "capitals",
    );
  });

  // A document that turns a feature off is asking for exactly what it gets.
  it("passes over a toggle turned off, and a number that states nothing", () => {
    const off = `<w:caps w:val="false"/><w:kern w:val="0"/><w:spacing w:val="0"/>`;
    expect(reportOf(`<w:p><w:r><w:rPr>${off}</w:rPr><w:t>a</w:t></w:r></w:p>`)).toStrictEqual([]);
  });

  // A paragraph's own w:spacing is the room around it, which is read.
  it("does not take the room around a paragraph for the room between its letters", () => {
    const spacing = `<w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr>`;
    expect(reportOf(`<w:p>${spacing}<w:r><w:t>a</w:t></w:r></w:p>`)).toStrictEqual([]);
  });

  it("names the paragraph a feature was met in, and the part", () => {
    const report = reportOf(
      `<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>gone</w:t></w:r></w:p>`,
    );
    expect(report).toStrictEqual([
      {
        kind: "hidden-text",
        effect: "moves-text",
        places: [{ part: "word/document.xml", paragraphIndex: 1 }],
      },
    ]);
  });

  // What a second section costs is a second page, not the break itself: one that
  // changes only a header or a column count leaves the geometry alone and reading
  // the last section loses nothing.
  it("counts a second section only where it makes a different page", () => {
    const wide = `<w:sectPr><w:pgSz w:w="15840" w:h="12240"/></w:sectPr>`;
    const same = `<w:p><w:pPr>${SECTION}</w:pPr></w:p>`;
    const differing = `<w:p><w:pPr>${wide}</w:pPr></w:p>`;
    expect(kinds(reportOf(differing))).toStrictEqual(["more-than-one-section"]);
    expect(reportOf(same)).toStrictEqual([]);
    expect(reportOf(`<w:p/>`)).toStrictEqual([]);
  });

  it("says nothing about a break that changes only the columns", () => {
    const columns = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:cols w:num="1"/></w:sectPr>`;
    expect(reportOf(`<w:p><w:pPr>${columns}</w:pPr></w:p>`)).toStrictEqual([]);
  });

  it("names columns only where the section asks for more than one", () => {
    const columns = (num: string) =>
      `<w:p><w:pPr><w:sectPr><w:cols w:num="${num}"/></w:sectPr></w:pPr></w:p>`;
    expect(kinds(reportOf(columns("2")))).toContain("text-columns");
    expect(kinds(reportOf(columns("1")))).not.toContain("text-columns");
  });

  it("names a merged cell, wherever the merge is stated", () => {
    const cell = (properties: string) =>
      `<w:tbl><w:tr><w:tc><w:tcPr>${properties}</w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
    expect(kinds(reportOf(cell(`<w:gridSpan w:val="2"/>`)))).toStrictEqual(["merged-cells"]);
    expect(kinds(reportOf(cell(`<w:vMerge w:val="restart"/>`)))).toStrictEqual(["merged-cells"]);
  });

  it("names a bar stop, and no other kind of stop", () => {
    const stop = (val: string) =>
      `<w:p><w:pPr><w:tabs><w:tab w:val="${val}" w:pos="720"/></w:tabs></w:pPr></w:p>`;
    expect(kinds(reportOf(stop("bar")))).toStrictEqual(["bar-tab-stop"]);
    expect(reportOf(stop("right"))).toStrictEqual([]);
  });

  it("names a drawing it can make neither a picture nor a shape of", () => {
    const drawing = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}">
      <wp:extent cx="914400" cy="914400"/>
      <a:graphic><a:graphicData><c:chart xmlns:c="urn:chart"/></a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p>`;
    expect(kinds(reportOf(drawing))).toStrictEqual(["unknown-drawing"]);
  });

  it("names a picture held in a format nothing here decodes, and not one it draws", () => {
    const picture = (target: string) => {
      const rels = `<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId9" Type="${R_NS}/image" Target="${target}"/></Relationships>`;
      const body = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}">
        <wp:extent cx="914400" cy="914400"/>
        <a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic xmlns:pic="${PIC_NS}">
          <pic:blipFill><a:blip r:embed="rId9" xmlns:r="${R_NS}"/></pic:blipFill>
        </pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p>`;
      return kinds(reportOf(body, { "word/_rels/document.xml.rels": rels }));
    };
    expect(picture("media/chart.wmf")).toStrictEqual(["undrawable-picture"]);
    expect(picture("media/logo.png")).toStrictEqual([]);
    expect(picture("media/logo.emf")).toStrictEqual([]);
  });

  it("reads the styles and the settings as well as the flow", () => {
    const styles = `<?xml version="1.0"?><w:styles xmlns:w="${WORDPROCESSING_NS}">
      <w:style w:type="table" w:styleId="Grid"><w:tblStylePr w:type="firstRow"/></w:style></w:styles>`;
    const settings = `<?xml version="1.0"?><w:settings xmlns:w="${WORDPROCESSING_NS}">
      <w:autoHyphenation/></w:settings>`;
    expect(
      kinds(reportOf(`<w:p/>`, { "word/styles.xml": styles, "word/settings.xml": settings })),
    ).toStrictEqual(["automatic-hyphenation", "table-style-conditional-formatting"]);
  });

  it("gathers every place a kind was met into the one entry", () => {
    const hidden = `<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>a</w:t></w:r></w:p>`;
    const [entry] = reportOf(hidden + hidden);
    expect(entry?.places).toStrictEqual([
      { part: "word/document.xml", paragraphIndex: 0 },
      { part: "word/document.xml", paragraphIndex: 1 },
    ]);
  });
});

// The one entry that is not in the document's own words: a face is only known to
// have been stood in for once the layout has asked for it.
describe("withSubstitutedFaces", () => {
  const kerning: Unhonoured = {
    kind: "character-kerning",
    effect: "moves-text",
    places: [{ part: "word/document.xml", paragraphIndex: 0 }],
  };

  it("leaves the list alone where every face the document asked for answered", () => {
    expect(withSubstitutedFaces([kerning], [])).toStrictEqual([kerning]);
  });

  it("puts a face that was stood in for in the same list as everything else", () => {
    const [entry] = withSubstitutedFaces([], [{ requested: { name: "Meridian Sans" } }]);
    expect(entry?.kind).toBe("substituted-face");
    expect(entry?.effect).toBe("moves-text");
    expect(entry?.places).toHaveLength(1);
  });

  it("keeps the list in the order of its kinds", () => {
    const merged = withSubstitutedFaces(
      [kerning, { kind: "text-columns", effect: "moves-text", places: [] }],
      [{ requested: { name: "Meridian Sans" } }],
    );
    expect(merged.map((each) => each.kind)).toStrictEqual([
      "character-kerning",
      "substituted-face",
      "text-columns",
    ]);
  });
});
