import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { readBlocks, type Block } from "./blocks.js";
import { openDocx } from "./package.js";
import { paragraphText } from "./paragraphs.js";

const blocksOf = (body: string) =>
  readBlocks(openDocx(buildDocx({ "word/document.xml": wordDocument(body) })));

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const para = (text: string) => `<w:p>${run(text)}</w:p>`;
const cell = (inner: string, properties = "") =>
  `<w:tc><w:tcPr>${properties}</w:tcPr>${inner}</w:tc>`;
const row = (...cells: readonly string[]) => `<w:tr>${cells.join("")}</w:tr>`;
const table = (...rows: readonly string[]) => `<w:tbl>${rows.join("")}</w:tbl>`;

const label = (block: Block): string =>
  block.kind === "paragraph"
    ? paragraphText(block.paragraph)
    : `table(${String(block.rows.length)})`;

describe("readBlocks", () => {
  it("keeps paragraphs and tables apart in document order", () => {
    const found = blocksOf(`${para("before")}${table(row(cell(para("in"))))}${para("after")}`);
    expect(found.map(label)).toStrictEqual(["before", "table(1)", "after"]);
  });

  it("groups a row's cells side by side rather than in sequence", () => {
    const [block] = blocksOf(table(row(cell(para("left")), cell(para("right")))));
    if (block?.kind !== "table") throw new Error("expected a table");
    expect(block.rows[0]?.cells.map((each) => each.blocks.map(label))).toStrictEqual([
      ["left"],
      ["right"],
    ]);
  });

  it("numbers paragraphs across cells in document order", () => {
    const found = blocksOf(
      `${table(row(cell(para("a")), cell(`${para("b")}${para("c")}`)))}${para("d")}`,
    );
    const indices = new Map<string, number>();
    const walk = (blocks: readonly Block[]): void => {
      for (const block of blocks) {
        if (block.kind === "paragraph") indices.set(label(block), block.paragraph.index);
        else for (const each of block.rows) for (const one of each.cells) walk(one.blocks);
      }
    };
    walk(found);
    expect([...indices]).toStrictEqual([
      ["a", 0],
      ["b", 1],
      ["c", 2],
      ["d", 3],
    ]);
  });

  it("reads a cell's vertical alignment, defaulting to top", () => {
    const [block] = blocksOf(
      table(row(cell(para("a")), cell(para("b"), `<w:vAlign w:val="center"/>`))),
    );
    if (block?.kind !== "table") throw new Error("expected a table");
    expect(block.rows[0]?.cells.map((each) => each.verticalAlign)).toStrictEqual(["top", "center"]);
  });

  // What the table itself says and nothing more: a margin it leaves out is left
  // for the style chain to answer and for Word's own to stand behind that.
  it("states no margin at all for a table that asks for none", () => {
    const [block] = blocksOf(table(row(cell(para("in")))));
    if (block?.kind !== "table") throw new Error("expected a table");
    expect(block.statedInsets).toStrictEqual({
      indentTwips: 0,
      leftTwips: null,
      rightTwips: null,
      topTwips: null,
      bottomTwips: null,
    });
  });

  it("reads the margins and the indent a table asks for", () => {
    const properties =
      `<w:tblPr><w:tblInd w:w="-5" w:type="dxa"/>` +
      `<w:tblCellMar><w:left w:w="72" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>` +
      `</w:tblPr>`;
    const [block] = blocksOf(`<w:tbl>${properties}${row(cell(para("in")))}</w:tbl>`);
    if (block?.kind !== "table") throw new Error("expected a table");
    expect(block.statedInsets).toStrictEqual({
      indentTwips: -5,
      leftTwips: 72,
      rightTwips: 0,
      topTwips: null,
      bottomTwips: null,
    });
  });

  it("reads the margins a cell asks for itself, and nothing for the sides it leaves out", () => {
    const own = `<w:tcMar><w:top w:w="288" w:type="dxa"/><w:left w:w="0" w:type="dxa"/></w:tcMar>`;
    const [block] = blocksOf(table(row(cell(para("in"), own))));
    if (block?.kind !== "table") throw new Error("expected a table");
    expect(block.rows[0]?.cells[0]?.margins).toStrictEqual({
      topTwips: 288,
      leftTwips: 0,
      bottomTwips: null,
      rightTwips: null,
    });
  });

  it("reads the height a row asks for, and whether it is a floor or the whole of it", () => {
    const asked = (properties: string) =>
      blocksOf(table(`<w:tr><w:trPr>${properties}</w:trPr>${cell(para("in"))}</w:tr>`))[0];

    const floor = asked(`<w:trHeight w:val="1440"/>`);
    const exact = asked(`<w:trHeight w:val="1440" w:hRule="exact"/>`);
    if (floor?.kind !== "table" || exact?.kind !== "table") throw new Error("expected tables");

    expect(floor.rows[0]?.height).toStrictEqual({ twips: 1440, exact: false });
    expect(exact.rows[0]?.height).toStrictEqual({ twips: 1440, exact: true });
  });

  it("leaves a row that asks for no height of its own without one", () => {
    const [block] = blocksOf(table(row(cell(para("in")))));
    if (block?.kind !== "table") throw new Error("expected a table");
    expect(block.rows[0]?.height).toBeNull();
  });

  // A stated nought is an answer, and a side left out is not one: telling the two
  // apart is the whole reason this stops short of Word's own margin.
  it("tells a margin stated as nought from a side the table leaves out", () => {
    const properties = `<w:tblPr><w:tblCellMar><w:left w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>`;
    const [block] = blocksOf(`<w:tbl>${properties}${row(cell(para("in")))}</w:tbl>`);
    if (block?.kind !== "table") throw new Error("expected a table");
    expect(block.statedInsets.leftTwips).toBe(0);
    expect(block.statedInsets.rightTwips).toBeNull();
  });

  it("reads a table nested inside a cell as its own block", () => {
    const [block] = blocksOf(table(row(cell(table(row(cell(para("deep"))))))));
    if (block?.kind !== "table") throw new Error("expected a table");
    expect(block.rows[0]?.cells[0]?.blocks.map(label)).toStrictEqual(["table(1)"]);
  });

  it("leaves out the paragraphs of a floating text box", () => {
    const body = `<w:p>${run("outer")}
      <w:r><w:drawing><wp:anchor xmlns:wp="x"><w:txbxContent>
        ${para("inside the box")}
      </w:txbxContent></wp:anchor></w:drawing></w:r></w:p>`;
    expect(blocksOf(body).map(label)).toStrictEqual(["outer"]);
  });
});
