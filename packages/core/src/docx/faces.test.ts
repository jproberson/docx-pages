import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { openDocx } from "./package.js";
import { facesUsed } from "./faces.js";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

const NORMAL = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr></w:style></w:styles>`;

const facesOf = (parts: Record<string, string>) =>
  facesUsed(openDocx(buildDocx({ "word/styles.xml": NORMAL, ...parts })));

const body = (inner: string) => ({ "word/document.xml": wordDocument(inner) });

const named = (faces: ReturnType<typeof facesUsed>) => faces.map((face) => face.name);

describe("facesUsed", () => {
  it("reports the face a paragraph inherits from the default style", () => {
    expect(facesOf(body(`<w:p><w:r><w:t>a</w:t></w:r></w:p>`))).toStrictEqual([
      { name: "Arial", bold: false, italic: false, sizesPt: [12] },
    ]);
  });

  it("keeps a run's own face apart from the one around it", () => {
    const faces = facesOf(
      body(`<w:p><w:r><w:t>a</w:t></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Georgia"/></w:rPr><w:t>b</w:t></w:r></w:p>`),
    );

    expect(named(faces)).toStrictEqual(["Arial", "Georgia"]);
  });

  it("counts bold as a face of its own, since it is a separate file", () => {
    const faces = facesOf(
      body(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r>
        <w:r><w:t>b</w:t></w:r></w:p>`),
    );

    expect(named(faces)).toStrictEqual(["Arial", "Arial"]);
    expect([...faces].map((face) => face.bold).sort()).toStrictEqual([false, true]);
  });

  it("gathers every size a face is asked for, in order", () => {
    const faces = facesOf(
      body(`<w:p><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>a</w:t></w:r>
        <w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t>b</w:t></w:r></w:p>`),
    );

    expect(faces[0]?.sizesPt).toStrictEqual([8, 12, 20]);
  });

  it("reaches inside a text box, whose paragraphs are laid out in their own faces", () => {
    const textBox = `<w:p><w:r><w:drawing><wp:anchor xmlns:wp="${WP_NS}">
      <wp:extent cx="100" cy="100"/><wp:docPr id="1" name="Box"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData>
        <wps:wsp xmlns:wps="${WPS_NS}"><wps:txbx><w:txbxContent>
          <w:p><w:r><w:rPr><w:rFonts w:ascii="Verdana"/></w:rPr><w:t>in</w:t></w:r></w:p>
        </w:txbxContent></wps:txbx></wps:wsp>
      </a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;

    expect(named(facesOf(body(textBox)))).toContain("Verdana");
  });

  it("reads the header as well as the body, since both are drawn", () => {
    const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const faces = facesOf({
      "word/document.xml": `<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="${R_NS}"><w:body>
          <w:p><w:r><w:t>a</w:t></w:r></w:p>
          <w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr>
        </w:body></w:document>`,
      "word/header1.xml": `<?xml version="1.0"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:p><w:r><w:rPr><w:rFonts w:ascii="Tahoma"/></w:rPr><w:t>h</w:t></w:r></w:p></w:hdr>`,
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Target="header1.xml"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"/>
        </Relationships>`,
    });

    expect(named(faces)).toContain("Tahoma");
  });

  it("says a face is unnamed rather than dropping a run the cascade left without one", () => {
    const bare = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;
    const faces = facesOf({
      ...body(`<w:p><w:r><w:t>a</w:t></w:r></w:p>`),
      "word/styles.xml": bare,
    });

    expect(named(faces)).toStrictEqual([null]);
  });
});
