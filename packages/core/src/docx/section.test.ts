import { describe, expect, it } from "vitest";

import { isOnePagerError, OnePagerError } from "../errors.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { openDocx } from "./package.js";
import { readSectionGeometry } from "./section.js";

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

    expect(isOnePagerError(thrown)).toBe(true);
    if (!(thrown instanceof OnePagerError)) throw new Error("expected a OnePagerError");
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

    if (!(thrown instanceof OnePagerError)) throw new Error("expected a OnePagerError");
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

    if (!(thrown instanceof OnePagerError)) throw new Error("expected a OnePagerError");
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

    if (!(thrown instanceof OnePagerError)) throw new Error("expected a OnePagerError");
    expect(thrown.code).toBe("docx-unreadable");
  });
});
