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
