import {
  readBorders,
  readShading,
  readTableBorders,
  type StatedBorders,
  type TableBorders,
} from "./borders.js";
import { MAIN_DOCUMENT_PART, partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import { attribute, firstNamed, toggledOn, type XmlElement } from "./xml.js";

export const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

export type Paragraph = {
  readonly index: number;
  readonly element: XmlElement;
};

export type CellVerticalAlign = "top" | "center" | "bottom";

// What a cell asks to hold its own content off its walls by, each side left out
// where it asks for nothing and the table's own margin stands.
export type CellMargins = {
  readonly leftTwips: number | null;
  readonly rightTwips: number | null;
  readonly topTwips: number | null;
  readonly bottomTwips: number | null;
};

// Where a cell stands in a run of rows merged down the page. A `restart` opens the
// run and holds everything drawn in it; a `continue` was swallowed by the one above
// and draws nothing at all.
export type CellMerge = "restart" | "continue";

export type TableCell = {
  readonly element: XmlElement;
  readonly blocks: readonly Block[];
  readonly verticalAlign: CellVerticalAlign;
  readonly margins: CellMargins;
  // How many of the table's grid columns the cell stands on, which is one unless it
  // says otherwise.
  readonly gridSpan: number;
  readonly merge: CellMerge | null;
  // What the cell asks for at each of its own edges, which the table's own lines
  // stand behind. Settling them takes the whole table: see `resolveCellBorders`.
  readonly borders: StatedBorders;
  readonly fillColor: string | null;
};

// How tall a row asks to be. A stated height is a floor under the row unless the
// row says it is exact, in which case it is the whole of the row whatever its
// cells hold.
export type RowHeight = {
  readonly twips: number;
  readonly exact: boolean;
};

export type TableRow = {
  readonly cells: readonly TableCell[];
  readonly height: RowHeight | null;
  readonly cantSplit: boolean;
};

// How far a table sits from the edge of the text it is laid out in, and how far
// its cells hold their own content off their edges.
export type TableInsets = {
  readonly indentTwips: number;
  readonly leftTwips: number;
  readonly rightTwips: number;
  readonly topTwips: number;
  readonly bottomTwips: number;
};

// The same as the file states it, each margin left out where the table's own
// `w:tblPr` names none. **A margin the table does not state is not the default
// yet**: its style may state one, and only where nothing in that cascade does
// either does Word's own stand.
export type StatedTableInsets = {
  readonly indentTwips: number;
  readonly leftTwips: number | null;
  readonly rightTwips: number | null;
  readonly topTwips: number | null;
  readonly bottomTwips: number | null;
};

// What Word leaves at each side of a cell when neither the table nor its style
// asks for anything: an eighth of an inch either side, which is the whole of the
// offset a bulleted list inside a table starts at, and nothing above or below.
export const DEFAULT_TABLE_INSETS: TableInsets = {
  indentTwips: 0,
  leftTwips: 108,
  rightTwips: 108,
  topTwips: 0,
  bottomTwips: 0,
};

// Which of a table style's conditional formats the table asks for, and how many
// rows and columns a band is. Word writes `w:tblLook` on every table it makes; a
// table that states none asks for none of them, which is what the attributes
// default to.
export type TableLook = {
  readonly firstRow: boolean;
  readonly lastRow: boolean;
  readonly firstColumn: boolean;
  readonly lastColumn: boolean;
  readonly horizontalBanding: boolean;
  readonly verticalBanding: boolean;
  // Undefined where the table states none of its own and the style's stands.
  readonly rowBandSize: number | undefined;
  readonly columnBandSize: number | undefined;
};

export const NO_TABLE_LOOK: TableLook = {
  firstRow: false,
  lastRow: false,
  firstColumn: false,
  lastColumn: false,
  horizontalBanding: false,
  verticalBanding: false,
  rowBandSize: undefined,
  columnBandSize: undefined,
};

// Where a table stands when `w:tblpPr` takes it out of the flow: what its own
// offsets are measured from, and how far off those origins it stands. `xSpec`
// names an edge instead of a distance, and stands in front of `xTwips` where a
// table states both.
export type TablePositioning = {
  readonly horizontalAnchor: "page" | "margin" | "column";
  readonly verticalAnchor: "page" | "margin" | "text";
  readonly xTwips: number;
  readonly yTwips: number;
  readonly xSpec: string | null;
  readonly ySpec: string | null;
  // How far the text has to stay off each side of it.
  readonly leftFromTextTwips: number;
  readonly rightFromTextTwips: number;
  readonly topFromTextTwips: number;
  readonly bottomFromTextTwips: number;
};

export type Block =
  | { readonly kind: "paragraph"; readonly paragraph: Paragraph }
  | {
      readonly kind: "table";
      readonly rows: readonly TableRow[];
      // Empty where the table declares no grid of its own.
      readonly gridTwips: readonly number[];
      readonly statedInsets: StatedTableInsets;
      readonly borders: TableBorders;
      // The table style the borders and the margins above stand in front of, which
      // the styles rather than the blocks can look up.
      readonly styleId: string | null;
      // Null in a table that flows with the text, which is all but ten of the 966.
      readonly positioning: TablePositioning | null;
      readonly look: TableLook;
    };

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
      const properties = firstNamed(tc, W_NS, "tcPr");
      cells.push({
        element: tc,
        blocks,
        verticalAlign: verticalAlignOf(tc),
        margins: cellMargins(tc),
        gridSpan: gridSpanOf(properties),
        merge: mergeOf(properties),
        borders: readBorders(properties, "tcBorders"),
        fillColor: readShading(properties),
      });
    }
    rows.push({ cells, height: rowHeight(tr), cantSplit: refusesToSplit(tr) });
  }
  const properties = firstNamed(element, W_NS, "tblPr");
  const style = properties === null ? null : firstNamed(properties, W_NS, "tblStyle");
  return {
    kind: "table",
    rows,
    gridTwips: gridColumns(element),
    statedInsets: statedTableInsets(properties),
    borders: readTableBorders(properties),
    styleId: style === null ? null : (attribute(style, W_NS, "val") ?? null),
    positioning: tablePositioning(properties),
    look: tableLook(properties),
  };
}

// **Word writes the switches twice**, once as attributes of their own and once as a
// hex mask in `w:val` that older files carry alone. The attributes win where they
// are there, and the mask answers where they are not.
function tableLook(properties: XmlElement | null): TableLook {
  const stated = properties === null ? null : firstNamed(properties, W_NS, "tblLook");
  if (stated === null) return NO_TABLE_LOOK;

  const mask = Number.parseInt(attribute(stated, W_NS, "val") ?? "", 16);
  const bit = (of: number): boolean => Number.isFinite(mask) && (mask & of) !== 0;
  const on = (name: string, at: number): boolean => {
    const value = attribute(stated, W_NS, name);
    return value === undefined ? bit(at) : value !== "0" && value !== "false" && value !== "off";
  };

  const bandSize = (name: string): number | undefined => {
    const held = properties === null ? null : firstNamed(properties, W_NS, name);
    const value = held === null ? Number.NaN : Number(attribute(held, W_NS, "val") ?? Number.NaN);
    return Number.isFinite(value) && value >= 1 ? Math.floor(value) : undefined;
  };

  return {
    firstRow: on("firstRow", 0x0020),
    lastRow: on("lastRow", 0x0040),
    firstColumn: on("firstColumn", 0x0080),
    lastColumn: on("lastColumn", 0x0100),
    // The mask says which banding is *off*, and the attributes say the same, so both
    // are read the same way round and both are turned over here.
    horizontalBanding: !on("noHBand", 0x0200),
    verticalBanding: !on("noVBand", 0x0400),
    rowBandSize: bandSize("tblStyleRowBandSize"),
    columnBandSize: bandSize("tblStyleColBandSize"),
  };
}

function tablePositioning(properties: XmlElement | null): TablePositioning | null {
  const stated = properties === null ? null : firstNamed(properties, W_NS, "tblpPr");
  if (stated === null) return null;

  const twips = (name: string): number => {
    const value = Number(attribute(stated, W_NS, name) ?? Number.NaN);
    return Number.isFinite(value) ? value : 0;
  };
  const horizontal = attribute(stated, W_NS, "horzAnchor");
  const vertical = attribute(stated, W_NS, "vertAnchor");

  return {
    horizontalAnchor: horizontal === "page" || horizontal === "margin" ? horizontal : "column",
    verticalAnchor: vertical === "page" || vertical === "margin" ? vertical : "text",
    xTwips: twips("tblpX"),
    yTwips: twips("tblpY"),
    xSpec: attribute(stated, W_NS, "tblpXSpec") ?? null,
    ySpec: attribute(stated, W_NS, "tblpYSpec") ?? null,
    leftFromTextTwips: twips("leftFromText"),
    rightFromTextTwips: twips("rightFromText"),
    topFromTextTwips: twips("topFromText"),
    bottomFromTextTwips: twips("bottomFromText"),
  };
}

function gridColumns(element: XmlElement): readonly number[] {
  const grid = firstNamed(element, W_NS, "tblGrid");
  if (grid === null) return [];
  return grid.children
    .filter((child) => child.namespace === W_NS && child.name === "gridCol")
    .map((column) => Number(attribute(column, W_NS, "w") ?? Number.NaN))
    .filter((twips) => Number.isFinite(twips));
}

// A span with no number on it is one column, and so is a number that is not one.
function gridSpanOf(properties: XmlElement | null): number {
  const span = properties === null ? null : firstNamed(properties, W_NS, "gridSpan");
  const value = span === null ? Number.NaN : Number(attribute(span, W_NS, "val") ?? Number.NaN);
  return Number.isFinite(value) && value >= 1 ? Math.floor(value) : 1;
}

// A `w:vMerge` saying nothing is a continuation, which is how Word writes every one
// of the 188 in the corpus document that holds most of this gap.
function mergeOf(properties: XmlElement | null): CellMerge | null {
  const merge = properties === null ? null : firstNamed(properties, W_NS, "vMerge");
  if (merge === null) return null;
  return attribute(merge, W_NS, "val") === "restart" ? "restart" : "continue";
}

// A height with no rule under it is a floor, which is what Word makes of one.
function rowHeight(row: XmlElement): RowHeight | null {
  const properties = firstNamed(row, W_NS, "trPr");
  const height = properties === null ? null : firstNamed(properties, W_NS, "trHeight");
  if (height === null) return null;

  const twips = Number(attribute(height, W_NS, "val") ?? Number.NaN);
  if (!Number.isFinite(twips)) return null;
  return { twips, exact: attribute(height, W_NS, "hRule") === "exact" };
}

function refusesToSplit(row: XmlElement): boolean {
  const properties = firstNamed(row, W_NS, "trPr");
  const stated = properties === null ? null : firstNamed(properties, W_NS, "cantSplit");
  return stated !== null && toggledOn(stated, W_NS);
}

function cellMargins(cell: XmlElement): CellMargins {
  const properties = firstNamed(cell, W_NS, "tcPr");
  const margins = properties === null ? null : firstNamed(properties, W_NS, "tcMar");
  const side = (name: string): number | null =>
    margins === null ? null : twipsIn(firstNamed(margins, W_NS, name));

  return {
    leftTwips: side("left"),
    rightTwips: side("right"),
    topTwips: side("top"),
    bottomTwips: side("bottom"),
  };
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

// What the table itself says, and nothing of what its style says. **A margin left
// out here is left out rather than defaulted**, since the style chain is read
// before Word's own eighth of an inch stands: `resolveTableInsets` finishes this.
export function statedTableInsets(properties: XmlElement | null): StatedTableInsets {
  const margins = properties === null ? null : firstNamed(properties, W_NS, "tblCellMar");
  const margin = (side: string): number | null =>
    margins === null ? null : twipsIn(firstNamed(margins, W_NS, side));

  return {
    indentTwips:
      (properties === null ? null : twipsIn(firstNamed(properties, W_NS, "tblInd"))) ??
      DEFAULT_TABLE_INSETS.indentTwips,
    leftTwips: margin("left"),
    rightTwips: margin("right"),
    topTwips: margin("top"),
    bottomTwips: margin("bottom"),
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
