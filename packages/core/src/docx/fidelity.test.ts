import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument, WORDPROCESSING_NS } from "../testing/build-docx.js";
import {
  readUnhonoured,
  withFallbackCharacters,
  withSubstitutedFaces,
  type Unhonoured,
} from "./fidelity.js";
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

// The same document with a header of its own, which is what the last section still
// answers for on every page.
const withAHeader = (body: string) =>
  readUnhonoured(
    openDocx(
      buildDocx({
        "word/document.xml": `<?xml version="1.0"?>
          <w:document xmlns:w="${WORDPROCESSING_NS}" xmlns:r="${R_NS}"><w:body>${body}
            <w:sectPr><w:headerReference w:type="default" r:id="rId1"/>
              <w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
          </w:body></w:document>`,
        "word/header1.xml": `<?xml version="1.0"?>
          <w:hdr xmlns:w="${WORDPROCESSING_NS}"><w:p/></w:hdr>`,
        "word/_rels/document.xml.rels": `<?xml version="1.0"?>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="header1.xml"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"/>
          </Relationships>`,
      }),
    ),
  );

describe("readUnhonoured", () => {
  // A highlight is painted now: the run's own advance across, the line's box down.
  it("says nothing about a highlighted run", () => {
    const body =
      `<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr>` + `<w:t>one</w:t></w:r></w:p>`;
    expect(kinds(reportOf(body))).toStrictEqual([]);
  });

  // A section's own first-page header and footer are drawn on the page it opens, so
  // w:titlePg stands in for nothing. The even-page pair is still read and never
  // chosen, which is what is left of this gap.
  it("says nothing about a section stating w:titlePg", () => {
    const body = `<w:p><w:pPr><w:sectPr><w:titlePg/></w:sectPr></w:pPr></w:p>`;
    expect(kinds(reportOf(body))).not.toContain("alternate-first-or-even-page");
  });

  it("names a document asking for a header of its own on even pages", () => {
    const settings = `<?xml version="1.0"?><w:settings xmlns:w="${WORDPROCESSING_NS}">
      <w:evenAndOddHeaders/></w:settings>`;
    const report = reportOf(`<w:p/>`, { "word/settings.xml": settings });
    expect(kinds(report)).toContain("alternate-first-or-even-page");
  });

  it("says nothing about a document holding only what is read", () => {
    expect(reportOf(`<w:p><w:r><w:t>plain</w:t></w:r></w:p>`)).toStrictEqual([]);
  });

  // **Read off the top of the deformed ranking on 2026-08-12.** Two documents of one
  // template lose six of their fourteen pages between them, the worst 61% wrong by the
  // raster, and both stated no gap at all: their fractions are drawn nowhere and
  // measured as an empty paragraph, so every equation costs 16pt and the text of it is
  // missing.
  it("names an equation, which is a language of its own and read by nothing here", () => {
    const fraction =
      `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">` +
      `<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>` +
      `</m:oMath>`;

    expect(kinds(reportOf(`<w:p>${fraction}</w:p>`))).toStrictEqual(["equation"]);
  });

  // The wrapper Word writes round an equation standing alone in its paragraph. Naming
  // both would count one equation twice.
  // The equation here is one this cannot read, since an equation of plain runs is
  // read now and is no longer a gap at all.
  it("names an equation once where a paragraph holds nothing else", () => {
    const oMath =
      `<m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num>` +
      `<m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>`;
    const alone =
      `<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">` +
      `${oMath}</m:oMathPara>`;

    expect(kinds(reportOf(`<w:p>${alone}</w:p>`))).toStrictEqual(["equation"]);
  });

  it("names a run's kerning and its capitals", () => {
    const marks = `<w:kern w:val="16"/><w:caps/>`;
    expect(kinds(reportOf(`<w:p><w:r><w:rPr>${marks}</w:rPr><w:t>a</w:t></w:r></w:p>`))).toContain(
      "character-kerning",
    );
    expect(kinds(reportOf(`<w:p><w:r><w:rPr>${marks}</w:rPr><w:t>a</w:t></w:r></w:p>`))).toContain(
      "capitals",
    );
  });

  // Letter spacing is measured and laid out rather than passed over, so a run
  // asking for it is asking for what it gets and the report says nothing.
  it("says nothing about a run's letter spacing, which is honoured", () => {
    const spaced = `<w:spacing w:val="20"/>`;
    expect(reportOf(`<w:p><w:r><w:rPr>${spaced}</w:rPr><w:t>a</w:t></w:r></w:p>`)).toStrictEqual(
      [],
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

  // Each page is made by the section whose text opened it, so a second page size
  // costs the body nothing. What is still read from the last section alone is the
  // header and the footer, drawn at its margins on every page, so a second page is
  // worth naming only where there is one of those to put in the wrong place.
  it("counts a second page only where a header stands to be drawn on it", () => {
    const wide = `<w:sectPr><w:pgSz w:w="15840" w:h="12240"/></w:sectPr>`;
    const same = `<w:p><w:pPr>${SECTION}</w:pPr></w:p>`;
    const differing = `<w:p><w:pPr>${wide}</w:pPr></w:p>`;
    expect(reportOf(differing)).toStrictEqual([]);
    expect(kinds(withAHeader(differing))).toStrictEqual(["more-than-one-section"]);
    expect(withAHeader(same)).toStrictEqual([]);
    expect(withAHeader(`<w:p/>`)).toStrictEqual([]);
  });

  it("says nothing about a break that changes only the columns", () => {
    const columns = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:cols w:num="1"/></w:sectPr>`;
    expect(reportOf(`<w:p><w:pPr>${columns}</w:pPr></w:p>`)).toStrictEqual([]);
  });

  // Columns were built on 2026-08-08, so a section running its text in more than one
  // of them stands in for nothing on its own.
  it("says nothing about a section that runs its text in more than one column", () => {
    const columns = (num: string) =>
      `<w:p><w:pPr><w:sectPr><w:cols w:num="${num}"/></w:sectPr></w:pPr></w:p>`;
    expect(kinds(reportOf(columns("2")))).not.toContain("text-columns");
  });

  // A column break is honoured where it stands alone in its paragraph or opens one,
  // which is where every one of the 25 in the corpus stands. One with text of its
  // own paragraph in front of it is a place inside a block and not between two.
  it("names a column break only where its paragraph has already drawn something", () => {
    const alone = `<w:p><w:r><w:br w:type="column"/></w:r></w:p>`;
    const opening = `<w:p><w:r><w:br w:type="column"/><w:t>after</w:t></w:r></w:p>`;
    const inside = `<w:p><w:r><w:t>before</w:t></w:r><w:r><w:br w:type="column"/></w:r></w:p>`;
    // A break in a run of its own with the paragraph's text in the run after it,
    // which is where twelve of the corpus put one.
    const ahead = `<w:p><w:r><w:br w:type="column"/></w:r><w:r><w:t>after</w:t></w:r></w:p>`;
    expect(kinds(reportOf(alone))).not.toContain("column-break");
    expect(kinds(reportOf(opening))).not.toContain("column-break");
    expect(kinds(reportOf(ahead))).not.toContain("column-break");
    expect(kinds(reportOf(inside))).toContain("column-break");
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
    // A conditional format holding a `w:pPr` alone is read now, so the one this
    // asks about shades a cell, which is not.
    const styles = `<?xml version="1.0"?><w:styles xmlns:w="${WORDPROCESSING_NS}">
      <w:style w:type="table" w:styleId="Grid"><w:tblStylePr w:type="firstRow">
        <w:tcPr><w:shd w:val="clear" w:fill="FF0000"/></w:tcPr></w:tblStylePr></w:style></w:styles>`;
    const settings = `<?xml version="1.0"?><w:settings xmlns:w="${WORDPROCESSING_NS}">
      <w:autoHyphenation/></w:settings>`;
    const table = `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr>
      <w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`;
    expect(
      kinds(reportOf(table, { "word/styles.xml": styles, "word/settings.xml": settings })),
    ).toStrictEqual(["automatic-hyphenation", "table-style-conditional-formatting"]);
  });

  // What a conditional format says about a paragraph or a run is read, so a format
  // stating only those asks for nothing it does not get.
  it("passes over a conditional format it reads the whole of", () => {
    const styles = `<?xml version="1.0"?><w:styles xmlns:w="${WORDPROCESSING_NS}">
      <w:style w:type="table" w:styleId="Grid"><w:tblStylePr w:type="firstRow">
        <w:pPr><w:jc w:val="center"/></w:pPr></w:tblStylePr></w:style></w:styles>`;
    const table = `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr>
      <w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`;
    expect(kinds(reportOf(table, { "word/styles.xml": styles }))).toStrictEqual([]);
  });

  // 481 of the 718 corpus documents state `w:kern` on a style alone, and the row
  // that led the ranking for days was counting every one of them.
  it("passes over a style nothing in the flow is written in", () => {
    const styles = (used: string) => `<?xml version="1.0"?><w:styles xmlns:w="${WORDPROCESSING_NS}">
      <w:style w:type="paragraph" w:styleId="${used}"><w:rPr><w:kern w:val="16"/></w:rPr></w:style>
      </w:styles>`;
    expect(kinds(reportOf(`<w:p/>`, { "word/styles.xml": styles("Body") }))).toStrictEqual([]);
    expect(
      kinds(
        reportOf(`<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr></w:p>`, {
          "word/styles.xml": styles("Body"),
        }),
      ),
    ).toStrictEqual(["character-kerning"]);
  });

  it("reads the style a paragraph's own style is based on, and the default one", () => {
    const styles = `<?xml version="1.0"?><w:styles xmlns:w="${WORDPROCESSING_NS}">
      <w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:kern w:val="16"/></w:rPr></w:style>
      <w:style w:type="paragraph" w:styleId="Body"><w:basedOn w:val="Base"/></w:style>
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:caps/></w:rPr></w:style>
      </w:styles>`;
    expect(
      kinds(
        reportOf(`<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr></w:p>`, {
          "word/styles.xml": styles,
        }),
      ),
    ).toStrictEqual(["capitals", "character-kerning"]);
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

// The other entry the layout finds rather than the document stating it: which
// characters a face had no glyph for is only known once the layout has asked.
describe("withFallbackCharacters", () => {
  const BULLET = 0x2022;

  it("leaves the list alone where every character its own face drew", () => {
    expect(withFallbackCharacters([], [])).toStrictEqual([]);
  });

  // The character takes the room Word gave it, so nothing moves; only the glyph
  // drawn in that room is anyone's guess.
  it("names a character drawn from another face as paint rather than as movement", () => {
    const [entry] = withFallbackCharacters([], [{ codePoint: BULLET }]);
    expect(entry?.kind).toBe("character-from-another-face");
    expect(entry?.effect).toBe("changes-paint");
  });

  it("keeps one place a character", () => {
    const [entry] = withFallbackCharacters([], [{ codePoint: BULLET }, { codePoint: 0x25a0 }]);
    expect(entry?.places).toHaveLength(2);
  });
});
