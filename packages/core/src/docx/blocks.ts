import { MAIN_DOCUMENT_PART, partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import { attribute, firstNamed, type XmlElement } from "./xml.js";

export const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

export type Paragraph = {
  readonly index: number;
  readonly element: XmlElement;
};

export type CellVerticalAlign = "top" | "center" | "bottom";

export type TableCell = {
  readonly element: XmlElement;
  readonly blocks: readonly Block[];
  readonly verticalAlign: CellVerticalAlign;
};

export type TableRow = {
  readonly cells: readonly TableCell[];
};

// How far a table sits from the edge of the text it is laid out in, and how far
// its cells hold their own content off their edges.
export type TableInsets = {
  readonly indentTwips: number;
  readonly leftTwips: number;
  readonly rightTwips: number;
};

// What Word leaves at each side of a cell when the table asks for nothing: an
// eighth of an inch, which is the whole of the offset a bulleted list inside a
// table starts at.
export const DEFAULT_TABLE_INSETS: TableInsets = {
  indentTwips: 0,
  leftTwips: 108,
  rightTwips: 108,
};

export type Block =
  | { readonly kind: "paragraph"; readonly paragraph: Paragraph }
  | { readonly kind: "table"; readonly rows: readonly TableRow[]; readonly insets: TableInsets };

// Text box content is laid out inside its own frame, and mc:Fallback repeats the
// shape that mc:Choice already described.
export const isDetachedContent = (node: XmlElement): boolean =>
  (node.namespace === W_NS && node.name === "txbxContent") ||
  (node.namespace === MC_NS && node.name === "Fallback");

type NextIndex = () => number;

function collect(node: XmlElement, into: Block[], nextIndex: NextIndex): void {
  for (const child of node.children) {
    if (isDetachedContent(child)) continue;
    if (child.namespace === W_NS && child.name === "p") {
      into.push({ kind: "paragraph", paragraph: { index: nextIndex(), element: child } });
      continue;
    }
    if (child.namespace === W_NS && child.name === "tbl") {
      into.push(readTable(child, nextIndex));
      continue;
    }
    collect(child, into, nextIndex);
  }
}

function readTable(element: XmlElement, nextIndex: NextIndex): Block {
  const rows: TableRow[] = [];
  for (const tr of element.children) {
    if (tr.namespace !== W_NS || tr.name !== "tr") continue;
    const cells: TableCell[] = [];
    for (const tc of tr.children) {
      if (tc.namespace !== W_NS || tc.name !== "tc") continue;
      const blocks: Block[] = [];
      collect(tc, blocks, nextIndex);
      cells.push({ element: tc, blocks, verticalAlign: verticalAlignOf(tc) });
    }
    rows.push({ cells });
  }
  return { kind: "table", rows, insets: tableInsets(element) };
}

// A width in a table is written as a value and the units it is given in, and only
// twips are honoured here, which is what Word writes for both an indent and a
// cell margin.
function twipsIn(element: XmlElement | null): number | null {
  if (element === null) return null;
  const units = attribute(element, W_NS, "type");
  if (units !== undefined && units !== "dxa") return null;
  const value = Number(attribute(element, W_NS, "w") ?? Number.NaN);
  return Number.isFinite(value) ? value : null;
}

function tableInsets(element: XmlElement): TableInsets {
  const properties = firstNamed(element, W_NS, "tblPr");
  if (properties === null) return DEFAULT_TABLE_INSETS;

  const margins = firstNamed(properties, W_NS, "tblCellMar");
  const margin = (side: string): number | null =>
    margins === null ? null : twipsIn(firstNamed(margins, W_NS, side));

  return {
    indentTwips:
      twipsIn(firstNamed(properties, W_NS, "tblInd")) ?? DEFAULT_TABLE_INSETS.indentTwips,
    leftTwips: margin("left") ?? DEFAULT_TABLE_INSETS.leftTwips,
    rightTwips: margin("right") ?? DEFAULT_TABLE_INSETS.rightTwips,
  };
}

function verticalAlignOf(cell: XmlElement): CellVerticalAlign {
  const properties = firstNamed(cell, W_NS, "tcPr");
  const vAlign = properties === null ? null : firstNamed(properties, W_NS, "vAlign");
  const value = vAlign === null ? undefined : attribute(vAlign, W_NS, "val");
  if (value === "center") return "center";
  if (value === "bottom") return "bottom";
  return "top";
}

export function readBlocks(pkg: DocxPackage, part: string = MAIN_DOCUMENT_PART): readonly Block[] {
  return blocksIn(partXml(pkg, part));
}

// A text box's blocks are numbered from zero within it, since they stack in its
// own frame rather than in the part's.
export function blocksIn(element: XmlElement): readonly Block[] {
  const blocks: Block[] = [];
  let index = 0;
  collect(element, blocks, () => index++);
  return blocks;
}

export function blockParagraphs(blocks: readonly Block[]): readonly Paragraph[] {
  const found: Paragraph[] = [];
  const visit = (of: readonly Block[]): void => {
    for (const block of of) {
      if (block.kind === "paragraph") found.push(block.paragraph);
      else for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
    }
  };
  visit(blocks);
  return found;
}

export const readParagraphs = (
  pkg: DocxPackage,
  part: string = MAIN_DOCUMENT_PART,
): readonly Paragraph[] => blockParagraphs(readBlocks(pkg, part));
