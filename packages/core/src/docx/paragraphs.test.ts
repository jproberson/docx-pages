import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { openDocx } from "./package.js";
import { readParagraphs } from "./blocks.js";
import { paragraphText } from "./paragraphs.js";

const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const paragraphsOf = (body: string) =>
  readParagraphs(openDocx(buildDocx({ "word/document.xml": wordDocument(body) })));

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;

const firstTextOf = (body: string): string => {
  const [paragraph] = paragraphsOf(body);
  if (paragraph === undefined) throw new Error("expected at least one paragraph");
  return paragraphText(paragraph);
};

describe("readParagraphs", () => {
  it("returns body paragraphs in document order", () => {
    const found = paragraphsOf(`<w:p>${run("one")}</w:p><w:p>${run("two")}</w:p>`);
    expect(found.map(paragraphText)).toStrictEqual(["one", "two"]);
  });

  it("ignores paragraphs nested inside a floating text box", () => {
    const body = `<w:p>${run("outer")}
      <w:r><w:drawing><wp:anchor xmlns:wp="x"><w:txbxContent>
        <w:p>${run("inside the box")}</w:p>
      </w:txbxContent></wp:anchor></w:drawing></w:r></w:p>`;
    expect(paragraphsOf(body)).toHaveLength(1);
    expect(firstTextOf(body)).toBe("outer");
  });

  it("ignores the mc:Fallback duplicate of an alternate-content shape", () => {
    const body = `<w:p>${run("kept")}</w:p>
      <mc:AlternateContent xmlns:mc="${MC_NS}">
        <mc:Choice Requires="wps"><w:p>${run("choice")}</w:p></mc:Choice>
        <mc:Fallback><w:p>${run("fallback")}</w:p></mc:Fallback>
      </mc:AlternateContent>`;
    expect(paragraphsOf(body).map(paragraphText)).toStrictEqual(["kept", "choice"]);
  });

  it("reports an empty spacer paragraph as empty", () => {
    expect(
      paragraphsOf(`<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr></w:p>`),
    ).toHaveLength(1);
    expect(firstTextOf(`<w:p/>`)).toBe("");
  });

  it("concatenates the runs of one paragraph without separators", () => {
    expect(firstTextOf(`<w:p>${run("Wind")}${run("break")}</w:p>`)).toBe("Reference");
  });
});
