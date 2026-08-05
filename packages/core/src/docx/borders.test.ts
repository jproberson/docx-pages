import { describe, expect, it } from "vitest";

import {
  borderExtentPt,
  readBorder,
  readBorders,
  readShading,
  readTableBorders,
  resolveCellBorders,
  NO_TABLE_BORDERS,
  type Border,
  type StatedBorders,
  type TableBorders,
} from "./borders.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { openDocx } from "./package.js";
import { readBlocks } from "./blocks.js";
import { W_NS } from "./section.js";
import { firstNamed, type XmlElement } from "./xml.js";
import { partXml } from "./package.js";

const xmlOf = (body: string): XmlElement =>
  partXml(openDocx(buildDocx({ "word/document.xml": wordDocument(body) })), "word/document.xml");

// The properties of the one paragraph in a document written round them.
function propertiesOf(inner: string): XmlElement {
  const body = firstNamed(xmlOf(`<w:p><w:pPr>${inner}</w:pPr></w:p>`), W_NS, "body");
  const paragraph = body === null ? null : firstNamed(body, W_NS, "p");
  const pPr = paragraph === null ? null : firstNamed(paragraph, W_NS, "pPr");
  if (pPr === null) throw new Error("expected paragraph properties");
  return pPr;
}

const oneBorder = (attributes: string): Border | null => {
  const properties = propertiesOf(`<w:pBdr><w:top ${attributes}/></w:pBdr>`);
  const container = firstNamed(properties, W_NS, "pBdr");
  return readBorder(container === null ? null : firstNamed(container, W_NS, "top"));
};

describe("readBorder", () => {
  it("reads a width in eighths of a point", () => {
    expect(oneBorder(`w:val="single" w:sz="12" w:color="FF0000"`)).toStrictEqual({
      style: "single",
      widthPt: 1.5,
      color: "#FF0000",
      spacePt: 0,
    });
  });

  it("draws nothing where the file asks for none, whatever width it states", () => {
    expect(oneBorder(`w:val="none" w:sz="8"`)).toBeNull();
    expect(oneBorder(`w:val="nil" w:sz="8"`)).toBeNull();
  });

  it("draws nothing for a pattern with no width behind it", () => {
    expect(oneBorder(`w:val="single" w:sz="0"`)).toBeNull();
    expect(oneBorder(`w:val="single"`)).toBeNull();
  });

  it("leaves an automatic colour to whatever draws the line", () => {
    expect(oneBorder(`w:val="single" w:sz="8" w:color="auto"`)?.color).toBeNull();
  });

  it("reads the room the line asks to stand off by, in points", () => {
    expect(oneBorder(`w:val="single" w:sz="8" w:space="12"`)?.spacePt).toBe(12);
  });

  it("reads a pattern it has no drawing of as the nearest one it has", () => {
    expect(oneBorder(`w:val="thick" w:sz="8"`)?.style).toBe("single");
    expect(oneBorder(`w:val="triple" w:sz="8"`)?.style).toBe("double");
    expect(oneBorder(`w:val="dotDash" w:sz="8"`)?.style).toBe("dashed");
  });

  it("reaches three times its own width where it is a double line", () => {
    expect(borderExtentPt(oneBorder(`w:val="double" w:sz="8"`))).toBe(3);
    expect(borderExtentPt(oneBorder(`w:val="single" w:sz="8"`))).toBe(1);
    expect(borderExtentPt(null)).toBe(0);
  });
});

describe("readShading", () => {
  it("reads the fill a paragraph asks for", () => {
    expect(readShading(propertiesOf(`<w:shd w:val="clear" w:fill="FFF2CC"/>`))).toBe("#FFF2CC");
  });

  it("leaves a paragraph that asks for none without one", () => {
    expect(readShading(propertiesOf(`<w:jc w:val="left"/>`))).toBeNull();
  });

  // Word reports a quarter of red over yellow as one colour, #FFBF00.
  it("mixes a pattern into the fill it stands on", () => {
    expect(
      readShading(propertiesOf(`<w:shd w:val="pct25" w:color="FF0000" w:fill="FFFF00"/>`)),
    ).toBe("#FFBF00");
  });

  it("leaves a pattern it cannot share out to the fill alone", () => {
    expect(
      readShading(propertiesOf(`<w:shd w:val="thinDiagStripe" w:color="FF0000" w:fill="FFFF00"/>`)),
    ).toBe("#FFFF00");
  });
});

const cellBorders = (inner: string): StatedBorders =>
  readBorders(propertiesOf(`<w:tcBorders>${inner}</w:tcBorders>`), "tcBorders");

const edge = (side: string, eighths: number, color: string): string =>
  `<w:${side} w:val="single" w:sz="${String(eighths)}" w:color="${color}"/>`;

const NOTHING = cellBorders("");

const tableBorders = (inner: string): TableBorders =>
  readTableBorders(propertiesOf(`<w:tblBorders>${inner}</w:tblBorders>`));

describe("resolveCellBorders", () => {
  it("gives a cell the table's outer line at the table's own edges", () => {
    const [row] = resolveCellBorders(
      [[NOTHING, NOTHING]],
      tableBorders(`${edge("top", 8, "FFC000")}${edge("left", 8, "FFC000")}`),
    );
    expect(row?.[0]?.left?.color).toBe("#FFC000");
    expect(row?.[1]?.left).toBeNull();
    expect(row?.[1]?.top?.color).toBe("#FFC000");
  });

  it("gives it the table's inside line between two of them", () => {
    const [row] = resolveCellBorders(
      [[NOTHING, NOTHING]],
      tableBorders(`<w:insideV w:val="single" w:sz="8" w:color="FFC000"/>`),
    );
    expect(row?.[0]?.right?.color).toBe("#FFC000");
    expect(row?.[1]?.left?.color).toBe("#FFC000");
  });

  it("lets the wider of two neighbours draw the line between them", () => {
    const [row] = resolveCellBorders(
      [[cellBorders(edge("right", 8, "FF0000")), cellBorders(edge("left", 24, "0070C0"))]],
      NO_TABLE_BORDERS,
    );
    expect(row?.[0]?.right?.color).toBe("#0070C0");
    expect(row?.[1]?.left?.color).toBe("#0070C0");
  });

  it("gives two of the same width to the cell on the left", () => {
    const [row] = resolveCellBorders(
      [[cellBorders(edge("right", 12, "FF0000")), cellBorders(edge("left", 12, "0070C0"))]],
      NO_TABLE_BORDERS,
    );
    expect(row?.[0]?.right?.color).toBe("#FF0000");
  });

  it("draws the line where one side asks for none and the other asks for one", () => {
    const [row] = resolveCellBorders(
      [[cellBorders(`<w:right w:val="nil"/>`), cellBorders(edge("left", 12, "0070C0"))]],
      NO_TABLE_BORDERS,
    );
    expect(row?.[0]?.right?.color).toBe("#0070C0");
  });

  it("lets a cell asking for none rub out the table's own edge", () => {
    const [row] = resolveCellBorders(
      [[cellBorders(`<w:left w:val="nil"/>`)]],
      tableBorders(edge("left", 8, "FFC000")),
    );
    expect(row?.[0]?.left).toBeNull();
  });

  it("shares the line between two rows with both of them", () => {
    const rows = resolveCellBorders(
      [[NOTHING], [cellBorders(edge("top", 24, "7030A0"))]],
      tableBorders(`<w:insideH w:val="single" w:sz="4" w:color="FFC000"/>`),
    );
    expect(rows[0]?.[0]?.bottom?.color).toBe("#7030A0");
    expect(rows[1]?.[0]?.top?.color).toBe("#7030A0");
  });
});

describe("readBlocks", () => {
  it("carries the lines and the fill a cell asks for", () => {
    const [block] = readBlocks(
      openDocx(
        buildDocx({
          "word/document.xml": wordDocument(
            `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblBorders>${edge("top", 8, "FFC000")}</w:tblBorders></w:tblPr>
              <w:tr><w:tc><w:tcPr><w:tcBorders>${edge("left", 4, "FF0000")}</w:tcBorders>
                <w:shd w:val="clear" w:fill="DEEBF7"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`,
          ),
        }),
      ),
    );
    if (block?.kind !== "table") throw new Error("expected a table");
    expect(block.styleId).toBe("Grid");
    expect(block.borders.top?.color).toBe("#FFC000");
    expect(block.rows[0]?.cells[0]?.borders.left?.widthPt).toBe(0.5);
    expect(block.rows[0]?.cells[0]?.fillColor).toBe("#DEEBF7");
  });
});
