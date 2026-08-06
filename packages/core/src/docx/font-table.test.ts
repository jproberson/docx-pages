import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { readFaceShapes } from "./font-table.js";
import { openDocx } from "./package.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const BODY = `<w:p><w:r><w:t>a</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;

const packageWith = (fonts: string) =>
  openDocx(
    buildDocx({
      "word/document.xml": wordDocument(BODY),
      "word/fontTable.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="${W_NS}">${fonts}</w:fonts>`,
    }),
  );

// The classifications Word itself writes, copied out of documents it saved:
// Calibri is a sans (PANOSE style 15, rounded), Times New Roman a serif, Courier
// New a monospace by proportion, and Wingdings a pictorial face whose style byte
// classifies nothing.
const CALIBRI = `<w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:family w:val="swiss"/></w:font>`;
const TIMES = `<w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:family w:val="roman"/></w:font>`;
const COURIER = `<w:font w:name="Courier New"><w:panose1 w:val="02070309020205020404"/><w:family w:val="modern"/></w:font>`;
const WINGDINGS = `<w:font w:name="Wingdings"><w:panose1 w:val="05000000000000000000"/><w:family w:val="auto"/></w:font>`;

describe("readFaceShapes", () => {
  it("classifies a face off its PANOSE bytes", () => {
    const shapes = readFaceShapes(packageWith(`${CALIBRI}${TIMES}${COURIER}`));
    expect(shapes.get("calibri")).toBe("sans-serif");
    expect(shapes.get("times new roman")).toBe("serif");
    expect(shapes.get("courier new")).toBe("monospace");
  });

  it("falls to the family where the PANOSE bytes say nothing", () => {
    const zeroed = `<w:font w:name="Housemade Sans"><w:panose1 w:val="00000000000000000000"/><w:family w:val="swiss"/></w:font>`;
    expect(readFaceShapes(packageWith(zeroed)).get("housemade sans")).toBe("sans-serif");
  });

  it("classifies nothing for a face that says nothing", () => {
    expect(readFaceShapes(packageWith(WINGDINGS)).get("wingdings")).toBeUndefined();
    const bare = `<w:font w:name="Housemade Display"/>`;
    expect(readFaceShapes(packageWith(bare)).get("housemade display")).toBeUndefined();
  });

  it("comes back empty for a document without the part", () => {
    const pkg = openDocx(buildDocx({ "word/document.xml": wordDocument(BODY) }));
    expect(readFaceShapes(pkg).size).toBe(0);
  });
});
