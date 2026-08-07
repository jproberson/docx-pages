import { W_NS } from "./section.js";
import { attribute, firstNamed, type XmlElement } from "./xml.js";

// The lines a table, a cell or a paragraph asks to be drawn round itself, and the
// colour behind its text. Both are painted rather than laid out, though a border
// takes room: see `how-word-draws-a-border` for what was measured off Word's own
// pdf and what it was measured from.

// Word names dozens of border patterns. These are the four it draws differently
// enough to matter, and every other name is read as the nearest of them.
export type BorderStyle = "single" | "double" | "dashed" | "dotted";

export type Border = {
  readonly style: BorderStyle;
  // Eighths of a point as the file states them, in points.
  readonly widthPt: number;
  // Null where the file leaves the colour to whatever is drawing the line.
  readonly color: string | null;
  // How far the line stands off the thing it runs round.
  readonly spacePt: number;
};

export type BorderSide = "top" | "left" | "bottom" | "right";

export type Borders = Readonly<Record<BorderSide, Border | null>>;

// What one side of one thing asks for. A side that states nothing takes whatever
// the table says at that edge; one that states `nil` refuses the line, which is
// not the same answer.
export type StatedBorders = Readonly<Record<BorderSide, Border | null | undefined>>;

export const NO_BORDERS: Borders = { top: null, left: null, bottom: null, right: null };

export const NOTHING_STATED: StatedBorders = {
  top: undefined,
  left: undefined,
  bottom: undefined,
  right: undefined,
};

// A table states the lines round the whole of itself and the lines between its
// cells in the one place.
export type TableBorders = StatedBorders & {
  readonly insideHorizontal: Border | null | undefined;
  readonly insideVertical: Border | null | undefined;
};

export const NO_TABLE_BORDERS: TableBorders = {
  ...NOTHING_STATED,
  insideHorizontal: undefined,
  insideVertical: undefined,
};

export const SIDES: readonly BorderSide[] = ["top", "left", "bottom", "right"];

// The patterns that are not drawn as a plain line. `thick` is one, whatever its
// name says: it is the stated width and nothing more.
const STYLES: Readonly<Record<string, BorderStyle>> = {
  double: "double",
  triple: "double",
  doubleWave: "double",
  dashed: "dashed",
  dashSmallGap: "dashed",
  dashDotStroked: "dashed",
  dotDash: "dashed",
  dotDotDash: "dotted",
  dotted: "dotted",
};

// A pattern that draws nothing at all. Every other name is a line of some kind,
// and one this project has no drawing for is drawn as a plain line rather than
// left out.
const NOTHING = new Set(["none", "nil"]);

const EIGHTHS_PER_POINT = 8;

// Whether the pattern a border names is one of the four that are drawn. Every
// other name comes out as a plain line of the stated width, which is near enough
// to look right and is not what Word draws.
export const drawnAsStated = (value: string): boolean =>
  value === "single" || value === "thick" || NOTHING.has(value) || STYLES[value] !== undefined;

export function readBorder(element: XmlElement | null): Border | null {
  if (element === null) return null;

  const value = attribute(element, W_NS, "val") ?? "single";
  if (NOTHING.has(value)) return null;

  const eighths = Number(attribute(element, W_NS, "sz") ?? Number.NaN);
  const widthPt = Number.isFinite(eighths) ? eighths / EIGHTHS_PER_POINT : 0;
  // Word's own dialogue offers no width under an eighth of a point, and a stated
  // pattern with no width draws nothing.
  if (widthPt <= 0) return null;

  const space = Number(attribute(element, W_NS, "space") ?? Number.NaN);
  return {
    style: STYLES[value] ?? "single",
    widthPt,
    color: colorOf(attribute(element, W_NS, "color")),
    spacePt: Number.isFinite(space) ? space : 0,
  };
}

// "auto" leaves the colour to whatever the line is drawn on, which is not a
// colour this can name.
const colorOf = (value: string | undefined): string | null =>
  value === undefined || value === "auto" ? null : `#${value.replace("#", "")}`;

const stated = (container: XmlElement, ...names: readonly string[]): Border | null | undefined => {
  for (const name of names) {
    const element = firstNamed(container, W_NS, name);
    if (element !== null) return readBorder(element);
  }
  return undefined;
};

const sidesOf = (container: XmlElement | null): StatedBorders =>
  container === null
    ? NOTHING_STATED
    : {
        top: stated(container, "top"),
        left: stated(container, "left", "start"),
        bottom: stated(container, "bottom"),
        right: stated(container, "right", "end"),
      };

// The borders a `w:tcPr` or a `w:pPr` states, read from the element holding them.
export const readBorders = (
  properties: XmlElement | null,
  name: "tcBorders" | "pBdr",
): StatedBorders => sidesOf(properties === null ? null : firstNamed(properties, W_NS, name));

export function readTableBorders(properties: XmlElement | null): TableBorders {
  const container = properties === null ? null : firstNamed(properties, W_NS, "tblBorders");
  if (container === null) return NO_TABLE_BORDERS;
  return {
    ...sidesOf(container),
    insideHorizontal: stated(container, "insideH"),
    insideVertical: stated(container, "insideV"),
  };
}

// A cell's lines: the one drawn along each of its sides, and the one it and its
// neighbour both asked for.
//
// They are not the same line. **A side that refuses a line its neighbour asks for
// is still drawn one, and leaves no room for it**: measured on 2026-08-07 by the
// authored `lined-rows` document, where four rows each refusing a line at their top
// and asking for one at their foot stand exactly as far apart as four rows with no
// lines at all, and Word draws the line between every pair of them all the same.
//
// **What two sides asking for different widths agree to is a guess.** Every case
// measured has the two the same, since a table states one `w:insideH` and both its
// neighbours take it; the wider is taken here because that is the line drawn, and
// nothing has asked Word whether the room follows the drawn line or the narrower
// ask. A document where the two differ would settle it.
export type CellBorders = {
  readonly drawn: Borders;
  readonly agreed: Borders;
};

// The lines round every cell of a table, each settled twice over: first through
// the cascade, where a cell's own side stands instead of whatever the table asks
// for at that edge, and then between neighbours, since the line between two cells
// is one line and both of them have something to say about it.

export function resolveCellBorders(
  rows: readonly (readonly StatedBorders[])[],
  table: TableBorders,
): readonly (readonly CellBorders[])[] {
  const cascaded = rows.map((cells, row) =>
    cells.map((own, column) => ({
      top: instead(own.top, row === 0 ? table.top : table.insideHorizontal),
      bottom: instead(own.bottom, row === rows.length - 1 ? table.bottom : table.insideHorizontal),
      left: instead(own.left, column === 0 ? table.left : table.insideVertical),
      right: instead(own.right, column === cells.length - 1 ? table.right : table.insideVertical),
    })),
  );

  // The edge of the table has no neighbour to agree with, so what the cell asked
  // for there is the whole of the answer.
  const agreedWith = (neighbour: Border | null | undefined, own: Border | null): Border | null =>
    neighbour === undefined
      ? own
      : neighbour === null || own === null
        ? null
        : strongerBorder(neighbour, own);

  return cascaded.map((cells, row) =>
    cells.map((each, column) => ({
      drawn: {
        top: strongerBorder(cascaded[row - 1]?.[column]?.bottom ?? null, each.top),
        bottom: strongerBorder(each.bottom, cascaded[row + 1]?.[column]?.top ?? null),
        left: strongerBorder(cells[column - 1]?.right ?? null, each.left),
        right: strongerBorder(each.right, cells[column + 1]?.left ?? null),
      },
      agreed: {
        top: agreedWith(cascaded[row - 1]?.[column]?.bottom, each.top),
        bottom: agreedWith(cascaded[row + 1]?.[column]?.top, each.bottom),
        left: agreedWith(cells[column - 1]?.right, each.left),
        right: agreedWith(cells[column + 1]?.left, each.right),
      },
    })),
  );
}

// How far across a drawn line actually reaches. A double line is three bands of
// the stated width, line, gap and line, so it is three times as wide as the one
// number in the file says; everything else is the width itself.
export const borderExtentPt = (border: Border | null): number =>
  border === null ? 0 : border.style === "double" ? border.widthPt * 3 : border.widthPt;

// What a side asks for, or what stands behind it where it asks for nothing. A
// side that states `nil` has asked for no line at all, which is an answer: only a
// side that states nothing takes the one behind it.
export const instead = (
  own: Border | null | undefined,
  behind: Border | null | undefined,
): Border | null => (own === undefined ? behind : own) ?? null;

// Which of two neighbours draws the line between them. The wider wins whichever
// side asked for it, and a side asking for none never does: a cell that asks for
// no line beside one that asks for a line still gets the line. Two of the same
// width go to the first, which is the cell above or to the left.
export function strongerBorder(one: Border | null, other: Border | null): Border | null {
  if (one === null) return other;
  if (other === null) return one;
  return other.widthPt > one.widthPt ? other : one;
}

// A fill as a colour, with a pattern blended into it. Word reports `pct25` of red
// over yellow as one colour, #FFBF00, so a pattern is a share of the foreground
// mixed into the fill rather than anything drawn.
export function readShading(properties: XmlElement | null): string | null {
  const shd = properties === null ? null : firstNamed(properties, W_NS, "shd");
  if (shd === null) return null;

  const fill = colorOf(attribute(shd, W_NS, "fill"));
  const pattern = attribute(shd, W_NS, "val") ?? "clear";
  if (pattern === "clear" || pattern === "nil") return fill;

  const share = shareOf(pattern);
  const foreground = colorOf(attribute(shd, W_NS, "color"));
  if (share === null || foreground === null || fill === null) return fill;
  return blend(fill, foreground, share);
}

// `pctN` is that many hundredths of the foreground. The named patterns are a
// weave rather than a share and are left to the fill alone.
function shareOf(pattern: string): number | null {
  const percent = /^pct(\d+)$/.exec(pattern);
  if (percent === null) return null;
  const value = Number(percent[1]);
  return Number.isFinite(value) ? Math.min(1, value / 100) : null;
}

function blend(fill: string, foreground: string, share: number): string {
  const channels = (hex: string): readonly number[] =>
    [1, 3, 5].map((at) => Number.parseInt(hex.slice(at, at + 2), 16));
  const under = channels(fill);
  const over = channels(foreground);
  if (under.some(Number.isNaN) || over.some(Number.isNaN)) return fill;

  return `#${under
    .map((value, at) => Math.round(value * (1 - share) + (over[at] ?? 0) * share))
    .map((value) => value.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()}`;
}
