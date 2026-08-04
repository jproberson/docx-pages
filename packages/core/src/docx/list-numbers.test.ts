import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { readBlocks } from "./blocks.js";
import { numberParagraphs } from "./list-numbers.js";
import { openDocx } from "./package.js";
import { W_NS } from "./section.js";
import { readStyleTable } from "./styles.js";

const BULLET_CHARACTER = "\uF0A7";

const level = (ilvl: number, format: string, text: string, extra = "") =>
  `<w:lvl w:ilvl="${String(ilvl)}"><w:numFmt w:val="${format}"/>
     <w:lvlText w:val="${text}"/>${extra}</w:lvl>`;

const numbering = (inner: string) => `<?xml version="1.0"?>
<w:numbering xmlns:w="${W_NS}">${inner}</w:numbering>`;

const list = (id: string, levels: string) =>
  `<w:abstractNum w:abstractNumId="${id}">${levels}</w:abstractNum>
   <w:num w:numId="${id}"><w:abstractNumId w:val="${id}"/></w:num>`;

const LISTS = numbering(
  `${list("1", `${level(0, "decimal", "%1.")}${level(1, "lowerLetter", "%1.%2)")}`)}
   ${list("2", level(0, "bullet", "&#xF0A7;"))}
   ${list("3", `${level(0, "upperRoman", "%1.")}${level(1, "decimal", "%2.", `<w:lvlRestart w:val="0"/>`)}`)}
   ${list("4", level(0, "decimalZero", "%1", `<w:start w:val="9"/>`))}
   ${list("5", level(0, "none", "%1."))}
   ${list("6", level(0, "ideographDigital", "%1"))}`,
);

const item = (numId: string, ilvl = 0) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="${String(ilvl)}"/>
     <w:numId w:val="${numId}"/></w:numPr></w:pPr></w:p>`;

const numbersOf = (body: string) => {
  const pkg = openDocx(
    buildDocx({ "word/document.xml": wordDocument(body), "word/numbering.xml": LISTS }),
  );
  return numberParagraphs(readBlocks(pkg), readStyleTable(pkg));
};

const textsOf = (body: string): readonly string[] => {
  const result = numbersOf(body);
  if (result.kind !== "numbered") throw new Error(result.kind);
  return [...result.numbers.values()].map((number) => number.text);
};

describe("numberParagraphs", () => {
  it("counts a list from one down the paragraphs that share its numId", () => {
    expect(textsOf(item("1").repeat(3))).toStrictEqual(["1.", "2.", "3."]);
  });

  it("passes over the paragraphs in between without counting them", () => {
    const body = `${item("1")}<w:p><w:r><w:t>aside</w:t></w:r></w:p>${item("1")}`;
    expect(textsOf(body)).toStrictEqual(["1.", "2."]);
  });

  it("counts each list of its own, even where they interleave", () => {
    const body = `${item("1")}${item("4")}${item("1")}`;
    expect(textsOf(body)).toStrictEqual(["1.", "09", "2."]);
  });

  it("repeats a bullet rather than counting it", () => {
    expect(textsOf(item("2").repeat(2))).toStrictEqual([BULLET_CHARACTER, BULLET_CHARACTER]);
  });

  it("writes a deeper level beside the count of the level above it", () => {
    const body = `${item("1")}${item("1", 1)}${item("1", 1)}`;
    expect(textsOf(body)).toStrictEqual(["1.", "1.a)", "1.b)"]);
  });

  it("starts a deeper level over once the level above it moves on", () => {
    const body = `${item("1")}${item("1", 1)}${item("1")}${item("1", 1)}`;
    expect(textsOf(body)).toStrictEqual(["1.", "1.a)", "2.", "2.a)"]);
  });

  it("runs a level on through the level above it when it says nothing restarts it", () => {
    const body = `${item("3")}${item("3", 1)}${item("3")}${item("3", 1)}`;
    expect(textsOf(body)).toStrictEqual(["I.", "1.", "II.", "2."]);
  });

  it("starts counting where the level says to", () => {
    expect(textsOf(item("4").repeat(2))).toStrictEqual(["09", "10"]);
  });

  it("writes no number for a level that asks for none", () => {
    expect(textsOf(item("5"))).toStrictEqual([""]);
  });

  it("keys each number by the paragraph it belongs to", () => {
    const result = numbersOf(`<w:p/>${item("1")}`);
    if (result.kind !== "numbered") throw new Error(result.kind);
    expect([...result.numbers.keys()]).toStrictEqual([1]);
  });

  it("stops on a format it cannot count in rather than numbering it wrong", () => {
    expect(numbersOf(`<w:p/>${item("6")}`)).toStrictEqual({
      kind: "unsupported",
      paragraphIndex: 1,
      numId: "6",
      ilvl: 0,
    });
  });
});
