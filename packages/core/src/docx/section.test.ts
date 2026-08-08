import { describe, expect, it } from "vitest";

import { isDocxPagesError, DocxPagesError } from "../errors.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { openDocx, partXml } from "./package.js";
import { endsASection, readSectionGeometry, readSections, W_NS } from "./section.js";
import { descendantsNamed, firstNamed, type XmlElement } from "./xml.js";

const LETTER_SECTION = `
  <w:p/>
  <w:sectPr>
    <w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="720" w:right="720" w:bottom="0" w:left="720" w:header="432" w:footer="144" w:gutter="0"/>
  </w:sectPr>`;

const geometryOf = (bodyXml: string, prefix?: string) =>
  readSectionGeometry(openDocx(buildDocx({ "word/document.xml": wordDocument(bodyXml, prefix) })));

describe("readSectionGeometry", () => {
  it("reads page size and margins in twips", () => {
    expect(geometryOf(LETTER_SECTION)).toStrictEqual({
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
    });
  });

  it("does not care which prefix the document binds the wordprocessing namespace to", () => {
    const renamed = LETTER_SECTION.replaceAll("w:", "x:");
    expect(geometryOf(renamed, "x")).toStrictEqual(geometryOf(LETTER_SECTION));
  });

  it("defaults the omitted gutter and footer to zero rather than guessing", () => {
    const sparse = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:left="1440"/></w:sectPr>`;
    expect(geometryOf(sparse).margin).toStrictEqual({
      topTwips: 1440,
      rightTwips: 0,
      bottomTwips: 0,
      leftTwips: 1440,
      headerTwips: 0,
      footerTwips: 0,
    });
  });

  it("reports a located error when the body has no section properties", () => {
    let thrown: unknown;
    try {
      geometryOf("<w:p/>");
    } catch (error: unknown) {
      thrown = error;
    }

    expect(isDocxPagesError(thrown)).toBe(true);
    if (!(thrown instanceof DocxPagesError)) throw new Error("expected a DocxPagesError");
    expect(thrown.code).toBe("docx-malformed");
    expect(thrown.at).toBe("core/docx/section.readSectionGeometry");
    expect(thrown.context["part"]).toBe("word/document.xml");
  });

  it("reports a located error when a dimension is not a number", () => {
    const bad = `<w:sectPr><w:pgSz w:w="wide" w:h="15840"/><w:pgMar w:top="0"/></w:sectPr>`;
    let thrown: unknown;
    try {
      geometryOf(bad);
    } catch (error: unknown) {
      thrown = error;
    }

    if (!(thrown instanceof DocxPagesError)) throw new Error("expected a DocxPagesError");
    expect(thrown.code).toBe("docx-malformed");
    expect(thrown.context["attribute"]).toBe("w");
    expect(thrown.context["value"]).toBe("wide");
  });
});

describe("openDocx", () => {
  it("reports a located error when a required part is missing", () => {
    let thrown: unknown;
    try {
      openDocx(buildDocx({}));
    } catch (error: unknown) {
      thrown = error;
    }

    if (!(thrown instanceof DocxPagesError)) throw new Error("expected a DocxPagesError");
    expect(thrown.code).toBe("docx-malformed");
    expect(thrown.context["part"]).toBe("word/document.xml");
  });

  it("reports a located error when the bytes are not a zip at all", () => {
    let thrown: unknown;
    try {
      openDocx(new Uint8Array([1, 2, 3, 4]));
    } catch (error: unknown) {
      thrown = error;
    }

    if (!(thrown instanceof DocxPagesError)) throw new Error("expected a DocxPagesError");
    expect(thrown.code).toBe("docx-unreadable");
  });
});

// A document is as many sections as it has `w:sectPr` elements: one on a paragraph
// ends the section that paragraph closes, and the body's own governs the text after
// the last of those.
describe("readSections", () => {
  const sectionsOf = (bodyXml: string) =>
    readSections(openDocx(buildDocx({ "word/document.xml": wordDocument(bodyXml) })));

  const ending = (properties: string): string =>
    `<w:p><w:pPr><w:sectPr>${properties}</w:sectPr></w:pPr></w:p>`;

  it("reads the one section of a document that has only the body's own", () => {
    const sections = sectionsOf(LETTER_SECTION);
    expect(sections).toHaveLength(1);
    expect(sections[0]?.geometry.widthTwips).toBe(12240);
  });

  it("reads every section in the order they run, the body's own last", () => {
    const sections = sectionsOf(`${ending(`<w:pgSz w:w="15840" w:h="12240"/>`)}${LETTER_SECTION}`);
    expect(sections.map((each) => each.geometry.widthTwips)).toStrictEqual([15840, 12240]);
  });

  // Word writes no type at all for the commonest break, so one saying nothing is
  // one that starts a page.
  it("reads the break a section stands on, and calls one that says nothing a page", () => {
    const sections = sectionsOf(
      `${ending(`<w:type w:val="continuous"/>`)}${ending("")}${LETTER_SECTION}`,
    );
    expect(sections.map((each) => each.breakKind)).toStrictEqual([
      "continuous",
      "nextPage",
      "nextPage",
    ]);
  });

  it("reads how many columns a section runs its text in, and answers one for silence", () => {
    const sections = sectionsOf(`${ending(`<w:cols w:num="3" w:space="720"/>`)}${LETTER_SECTION}`);
    expect(sections.map((each) => each.columns.count)).toStrictEqual([3, 1]);
    expect(sections[0]?.columns.spaceTwips).toBe(720);
    expect(sections[0]?.columns.widthsTwips).toStrictEqual([]);
  });

  it("reads the width and the gap of each column where the section states them", () => {
    const stated =
      `<w:cols w:num="2" w:space="720" w:equalWidth="0">` +
      `<w:col w:w="4441" w:space="1416"/><w:col w:w="6343"/></w:cols>`;
    const sections = sectionsOf(`${ending(stated)}${LETTER_SECTION}`);
    expect(sections[0]?.columns.widthsTwips).toStrictEqual([4441, 6343]);
    expect(sections[0]?.columns.gapsTwips).toStrictEqual([1416, 0]);
  });

  // A section asking for equal widths is divided rather than read, whatever widths
  // Word left beside the request.
  it("divides a section asking for equal widths rather than reading the widths it wrote", () => {
    const stated = `<w:cols w:num="2" w:space="720"><w:col w:w="4441"/><w:col w:w="6343"/></w:cols>`;
    const sections = sectionsOf(`${ending(stated)}${LETTER_SECTION}`);
    expect(sections[0]?.columns.widthsTwips).toStrictEqual([]);
  });

  it("gives the last section's page to a reader asking for the document's", () => {
    const body = `${ending(`<w:pgSz w:w="15840" w:h="12240"/>`)}${LETTER_SECTION}`;
    expect(
      readSectionGeometry(openDocx(buildDocx({ "word/document.xml": wordDocument(body) }))),
    ).toStrictEqual(sectionsOf(body).at(-1)?.geometry);
  });
});

describe("endsASection", () => {
  const firstParagraphOf = (xml: string): XmlElement => {
    const root = partXml(
      openDocx(buildDocx({ "word/document.xml": wordDocument(xml + LETTER_SECTION) })),
      "word/document.xml",
    );
    const body = firstNamed(root, W_NS, "body");
    const paragraph = descendantsNamed(body ?? root, W_NS, "p")[0];
    if (paragraph === undefined) throw new Error("the body holds no paragraph");
    return paragraph;
  };

  it("says a paragraph carrying section properties closes a section", () => {
    expect(endsASection(firstParagraphOf(`<w:p><w:pPr><w:sectPr/></w:pPr></w:p>`))).toBe(true);
  });

  it("says an ordinary paragraph closes none", () => {
    expect(endsASection(firstParagraphOf(`<w:p><w:r><w:t>a</w:t></w:r></w:p>`))).toBe(false);
  });
});
