import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { openDocx } from "./package.js";
import { defaultHeaderPart, readRelationships, relationshipsPartFor } from "./relationships.js";

const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const HEADER_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";

const rels = (inner: string) =>
  `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}">${inner}</Relationships>`;

const sectionWithHeaders = (references: string) =>
  wordDocument(`<w:p/><w:sectPr>${references}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`);

const packageWith = (documentXml: string, relsXml?: string) =>
  openDocx(
    buildDocx({
      "word/document.xml": documentXml,
      "word/header1.xml": `<?xml version="1.0"?><w:hdr
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>`,
      ...(relsXml === undefined ? {} : { "word/_rels/document.xml.rels": relsXml }),
    }),
  );

describe("relationshipsPartFor", () => {
  it("puts the rels file beside the part it describes", () => {
    expect(relationshipsPartFor("word/document.xml")).toBe("word/_rels/document.xml.rels");
    expect(relationshipsPartFor("word/header1.xml")).toBe("word/_rels/header1.xml.rels");
  });
});

describe("readRelationships", () => {
  it("resolves targets against the directory of the owning part", () => {
    const pkg = packageWith(
      sectionWithHeaders(""),
      rels(`<Relationship Id="rId1" Type="${HEADER_TYPE}" Target="header1.xml"/>`),
    );
    expect(readRelationships(pkg, "word/document.xml").get("rId1")?.part).toBe("word/header1.xml");
  });

  it("normalises a target that walks up a directory", () => {
    const pkg = packageWith(
      sectionWithHeaders(""),
      rels(`<Relationship Id="rId1" Type="${HEADER_TYPE}" Target="../customXml/item1.xml"/>`),
    );
    expect(readRelationships(pkg, "word/document.xml").get("rId1")?.part).toBe(
      "customXml/item1.xml",
    );
  });

  it("is empty when the part has no rels file", () => {
    expect(readRelationships(packageWith(sectionWithHeaders("")), "word/document.xml").size).toBe(
      0,
    );
  });
});

describe("defaultHeaderPart", () => {
  it("follows the default header reference to its part", () => {
    const pkg = packageWith(
      sectionWithHeaders(`<w:headerReference w:type="default" r:id="rId7"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`),
      rels(`<Relationship Id="rId7" Type="${HEADER_TYPE}" Target="header1.xml"/>`),
    );
    expect(defaultHeaderPart(pkg)).toBe("word/header1.xml");
  });

  it("ignores first and even header references, which do not apply to page one by default", () => {
    const pkg = packageWith(
      sectionWithHeaders(`<w:headerReference w:type="even" r:id="rId7"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`),
      rels(`<Relationship Id="rId7" Type="${HEADER_TYPE}" Target="header1.xml"/>`),
    );
    expect(defaultHeaderPart(pkg)).toBeNull();
  });

  it("is null when the section declares no header", () => {
    expect(defaultHeaderPart(packageWith(sectionWithHeaders("")))).toBeNull();
  });

  it("is null when the reference points at a part the package does not contain", () => {
    const pkg = packageWith(
      sectionWithHeaders(`<w:headerReference w:type="default" r:id="rId7"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`),
      rels(`<Relationship Id="rId7" Type="${HEADER_TYPE}" Target="header9.xml"/>`),
    );
    expect(defaultHeaderPart(pkg)).toBeNull();
  });
});
