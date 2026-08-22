import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { openDocx } from "./package.js";
import { numberingLevel, readNumberingTable } from "./numbering.js";
import { W_NS } from "./section.js";
import { firstNamed } from "./xml.js";

const numbering = (inner: string) => `<?xml version="1.0"?>
<w:numbering xmlns:w="${W_NS}">${inner}</w:numbering>`;

// Word writes a Wingdings bullet as a private use character, which has to survive
// both the XML parse and every lookup after it.
const BULLET_CHARACTER = "\uF0A7";

const BULLET = `<w:lvl w:ilvl="0">
  <w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A7;"/>
  <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr>
</w:lvl>`;

const DECIMAL = `<w:lvl w:ilvl="1">
  <w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/>
  <w:suff w:val="space"/><w:lvlRestart w:val="0"/>
  <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
</w:lvl>`;

const tableOf = (inner: string) => {
  const parts: Record<string, string> = { "word/document.xml": wordDocument("<w:p/>") };
  if (inner !== "") parts["word/numbering.xml"] = numbering(inner);
  return readNumberingTable(openDocx(buildDocx(parts)));
};

const levelOf = (inner: string, numId: string, ilvl: number) =>
  numberingLevel(tableOf(inner), numId, ilvl);

const ONE_LIST = `<w:abstractNum w:abstractNumId="7">${BULLET}${DECIMAL}</w:abstractNum>
  <w:num w:numId="4"><w:abstractNumId w:val="7"/></w:num>`;

describe("readNumberingTable", () => {
  it("resolves a numId through its abstract definition to the level at that ilvl", () => {
    expect(levelOf(ONE_LIST, "4", 0)).toMatchObject({
      ilvl: 0,
      format: "bullet",
      text: BULLET_CHARACTER,
      start: 1,
      suffix: "tab",
      restart: { kind: "any-higher" },
    });
  });

  it("reads the start, suffix and restart rule a level spells out", () => {
    expect(levelOf(ONE_LIST, "4", 1)).toMatchObject({
      format: "decimal",
      text: "%2.",
      start: 3,
      suffix: "space",
      restart: { kind: "never" },
    });
  });

  // **A start written out empty is a stated nought, and so is one left out.** Asked
  // of Word on 2026-08-22, three paragraphs a level: `w:start w:val=""` was marked
  // 0. 1. 2., exactly as `w:val="0"` and as a level stating no start at all were,
  // where `w:val="1"` was marked 1. 2. 3.
  it("numbers from nought where a level writes its start out empty", () => {
    const empty = `<w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0">
      <w:start w:val=""/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
      <w:num w:numId="4"><w:abstractNumId w:val="7"/></w:num>`;
    expect(levelOf(empty, "4", 0)?.start).toBe(0);
  });

  it("numbers from nought where a level states no start at all", () => {
    const silent = `<w:abstractNum w:abstractNumId="8"><w:lvl w:ilvl="0">
      <w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
      <w:num w:numId="5"><w:abstractNumId w:val="8"/></w:num>`;
    expect(levelOf(silent, "5", 0)?.start).toBe(0);
  });

  it("keeps the level's own properties for the style cascade to read", () => {
    const level = levelOf(ONE_LIST, "4", 0);
    if (level === null) throw new Error("expected a level");
    expect(firstNamed(level.properties, W_NS, "pPr")).not.toBeNull();
    expect(firstNamed(level.properties, W_NS, "rPr")).not.toBeNull();
  });

  it("has no level for a numId the numbering part never defines", () => {
    expect(levelOf(ONE_LIST, "9", 0)).toBeNull();
  });

  it("has no level past the deepest one its abstract definition gives", () => {
    expect(levelOf(ONE_LIST, "4", 2)).toBeNull();
  });

  it("reports a format it cannot number rather than guessing at one", () => {
    const list = `<w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="ideographDigital"/><w:lvlText w:val="%1"/></w:lvl>
      </w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;
    expect(levelOf(list, "1", 0)?.format).toBe("unsupported");
  });

  it("lets a level override replace the level the num points at", () => {
    const list = `<w:abstractNum w:abstractNumId="7">${BULLET}</w:abstractNum>
      <w:num w:numId="4"><w:abstractNumId w:val="7"/>
        <w:lvlOverride w:ilvl="0">
          <w:lvl w:ilvl="0"><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1)"/></w:lvl>
        </w:lvlOverride></w:num>`;
    expect(levelOf(list, "4", 0)).toMatchObject({ format: "upperRoman", text: "%1)" });
  });

  it("takes a start override without disturbing the rest of the level", () => {
    const list = `<w:abstractNum w:abstractNumId="7">${BULLET}</w:abstractNum>
      <w:num w:numId="4"><w:abstractNumId w:val="7"/>
        <w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>`;
    expect(levelOf(list, "4", 0)).toMatchObject({ format: "bullet", start: 5 });
  });

  it("leaves the numbering of one num alone when another overrides the same abstract", () => {
    const list = `<w:abstractNum w:abstractNumId="7">${BULLET}</w:abstractNum>
      <w:num w:numId="4"><w:abstractNumId w:val="7"/></w:num>
      <w:num w:numId="5"><w:abstractNumId w:val="7"/>
        <w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>`;
    expect(levelOf(list, "4", 0)?.start).toBe(1);
    expect(levelOf(list, "5", 0)?.start).toBe(5);
  });

  it("has no levels at all when the package carries no numbering part", () => {
    expect(tableOf("").levels.size).toBe(0);
  });
});
