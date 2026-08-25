import {
  DEFAULT_TABLE_INSETS,
  statedTableInsets,
  type Paragraph,
  type StatedTableInsets,
  type TableInsets,
} from "./blocks.js";
import {
  readBorders,
  readShading,
  readTableBorders,
  NOTHING_STATED,
  NO_TABLE_BORDERS,
  type Borders,
  type StatedBorders,
  type TableBorders,
} from "./borders.js";
import {
  numberingLevel,
  readNumberingTable,
  type NumberingLevel,
  type NumberingTable,
} from "./numbering.js";
import { MATH_NS, readMathFont } from "./equations.js";
import { drawsInLine, paragraphRuns } from "./paragraphs.js";
import { TWIPS_PER_POINT } from "../layout/units.js";
import { partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import {
  attribute,
  childrenNamed,
  descendantsNamed,
  firstNamed,
  statedNumber,
  statesOn,
  toggledOn,
  type XmlElement,
} from "./xml.js";

export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

const STYLES_PART = "word/styles.xml";
const THEME_PART = "word/theme/theme1.xml";

export const WORD_DEFAULT_FONT_SIZE_PT = 10;

/**
 * The smallest a run may be, whatever size it states: half a point, which is the
 * smallest `w:sz` can spell, being written in half-points.
 *
 * **A stated size of nought is not nought, and it is not inherited either.** Measured
 * on 2026-08-13 against the room a footer takes out of the body, which is the one
 * place a height too small to see still decides something. Thirty-one documents whose
 * body is thirty-five lines of exactly 20pt in a body of 720, closed by one more line
 * told to be exactly so many points: that last line stays on page one until the room
 * runs out, so the largest that stays is 20 less what the footer took.
 *
 * | the footer's one paragraph | the last line that stayed | so it took |
 * | -------------------------- | ------------------------- | ---------- |
 * | none at all                | 20.00pt                   | nothing    |
 * | `w:sz w:val="0"`           | 19.30pt                   | 0.60-0.70  |
 * | `w:sz w:val="1"`, half a point | 19.30pt                | 0.60-0.70  |
 * | `w:sz w:val="2"`, one point   | 18.70pt                | 1.20-1.30  |
 *
 * **Nought and a half-point are the same answer, and a whole point is twice it**, so
 * the nought is held to the smallest size the attribute can state rather than to a
 * point or to the size it would otherwise inherit. A point was tried first and is
 * what a coarser reading of the same question gave: it is twice too much, and it cost
 * two corpus documents a picture apiece at the foot of a page.
 *
 * The corpus turns on it through a footer holding nothing but a sensitivity label in
 * an anchored box, whose one flow paragraph states a size of nought: measured as
 * nothing it holds the body off nothing, and the body then keeps a line Word sends to
 * the next page.
 */
export const SMALLEST_FONT_SIZE_PT = 0.5;

export type FontChoice =
  { readonly kind: "named"; readonly name: string } | { readonly kind: "unresolved" };

export type VerticalAlign = "baseline" | "superscript" | "subscript";

export type ParagraphMark = {
  readonly font: FontChoice;
  // What the run is actually set at, which a superscript or a subscript has
  // already shrunk; the size the file declared is not what Word draws.
  readonly fontSizePt: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  // How far the run sits off the line's baseline, upwards.
  readonly raisePt: number;
  // The line a run stands on is measured at the size it was declared at and moved
  // by the raise it asked for by name: **a script shrinks and lifts what is drawn
  // and leaves the line exactly as it was.** Measured on 2026-08-07 by the authored
  // `raised-text` document, where a 24pt superscript and a 24pt subscript beside
  // 12pt text made the same 29.28pt line, which is what 24pt text makes on its own
  // and is neither run's drawn size nor either raise.
  readonly lineSizePt: number;
  readonly lineRaisePt: number;
  // Null where the run leaves its colour to whatever it is drawn on.
  readonly color: string | null;
  // Extra width laid after every character of the run, the last one included:
  // `abcdef` at five points ran 33pt to 63pt, which is six characters' worth
  // rather than five gaps'. Negative tightens.
  readonly characterSpacingPt: number;
  // What every glyph's own advance is multiplied by, and what it is drawn stretched
  // to. One where the run states no scale of its own.
  readonly characterScale: number;
  // The size, in half-points, at and above which the run's pairs kern. Null where the
  // cascade states none at all, which is a run that does not kern.
  readonly kernFromHalfPoints: number | null;
  // What is painted behind the run, as a colour, or null for a run Word paints
  // nothing behind. A highlight is one of sixteen names rather than a colour of the
  // document's choosing, and `none` turns an inherited one off.
  readonly highlight: string | null;
  // Whether the run's letters are drawn as capitals, and at what size the ones that
  // were not capitals already are set. See `capitalised` in `runs.ts`.
  readonly capitals: Capitals;
};

// `w:caps` draws every letter as a capital at the run's own size. `w:smallCaps`
// draws a letter that was not a capital as one at four fifths of that size, and
// leaves a capital, a digit and a mark alone. Both stated is `w:caps`.
export type Capitals = "none" | "all" | "small";

type PartialMark = {
  readonly fontName: string | undefined;
  readonly fontSizeHalfPoints: number | undefined;
  readonly bold: boolean | undefined;
  readonly italic: boolean | undefined;
  readonly underline: boolean | undefined;
  readonly verticalAlign: VerticalAlign | undefined;
  readonly positionHalfPoints: number | undefined;
  readonly color: string | undefined;
  readonly characterSpacingTwentieths: number | undefined;
  readonly characterScalePercent: number | undefined;
  readonly kernFromHalfPoints: number | undefined;
  // Null rather than undefined for `w:highlight w:val="none"`, which is a run
  // turning an inherited highlight off rather than saying nothing.
  readonly highlight: string | null | undefined;
  readonly capitals: Capitals | undefined;
};

// What a stop does with the text that follows a tab reaching it: a left stop
// starts it there, the next three line it up on the stop, and a bar is not a place
// a tab lands at all but a line drawn down the page.
export type TabAlignment = "left" | "center" | "right" | "decimal" | "bar";

// A stop is either declared at a position or clears one the cascade already put
// there, so the two arrive in the same list and are resolved in order.
type TabStopEntry = {
  readonly positionTwips: number;
  readonly alignment: TabAlignment;
  readonly kind: "stop" | "clear";
};

export type TabStop = {
  readonly positionTwips: number;
  readonly alignment: TabAlignment;
};

type PartialFrame = {
  readonly alignment: ParagraphAlignment | undefined;
  readonly indentLeftTwips: number | undefined;
  readonly indentRightTwips: number | undefined;
  readonly indentFirstLineTwips: number | undefined;
  readonly spaceBeforeTwips: number | undefined;
  readonly automaticSpaceBefore: boolean | undefined;
  readonly spaceAfterTwips: number | undefined;
  readonly automaticSpaceAfter: boolean | undefined;
  readonly lineTwips: number | undefined;
  readonly lineRule: LineRule | undefined;
  readonly widowControl: boolean | undefined;
  readonly keepNext: boolean | undefined;
  readonly pageBreakBefore: boolean | undefined;
  readonly contextualSpacing: boolean | undefined;
  readonly tabStops: readonly TabStopEntry[] | undefined;
  readonly borders: StatedBorders;
  readonly fillColor: string | null | undefined;
};

// Either half can arrive on its own: a style names the list, a paragraph the
// level within it.
type PartialNumbering = {
  readonly numId: string | undefined;
  readonly ilvl: number | undefined;
};

// What a table style says about a cell standing in one of the places it names.
// Word writes each as a `w:tblStylePr` under the style, and the type is the place.
type ConditionalFormat = {
  readonly mark: PartialMark;
  readonly frame: PartialFrame;
};

type StyleDefinition = {
  readonly id: string;
  readonly basedOn: string | undefined;
  readonly mark: PartialMark;
  readonly frame: PartialFrame;
  readonly numbering: PartialNumbering;
  readonly tableBorders: TableBorders;
  // What a table style holds its cells' content off their walls by, each side left
  // out where the style states none.
  readonly tableInsets: StatedTableInsets;
  // Empty for everything that is not a table style, and for a table style that
  // formats every cell alike.
  readonly conditional: ReadonlyMap<string, ConditionalFormat>;
  // How many rows and columns a band is, which a table style states and a table
  // may state again over it.
  readonly rowBandSize: number | undefined;
  readonly columnBandSize: number | undefined;
};

export type StyleTable = {
  readonly byId: ReadonlyMap<string, StyleDefinition>;
  readonly defaultParagraphStyleId: string | undefined;
  readonly docDefaults: PartialMark;
  readonly docDefaultsFrame: PartialFrame;
  readonly themeFonts: ReadonlyMap<string, string>;
  readonly numbering: NumberingTable;
  // What a math run naming no face of its own is set in, which is the settings
  // part's own say and the only thing an equation needs that a style does not hold.
  readonly mathFont: string;
};

const EMPTY: PartialMark = {
  fontName: undefined,
  fontSizeHalfPoints: undefined,
  bold: undefined,
  italic: undefined,
  underline: undefined,
  verticalAlign: undefined,
  positionHalfPoints: undefined,
  color: undefined,
  characterSpacingTwentieths: undefined,
  characterScalePercent: undefined,
  kernFromHalfPoints: undefined,
  highlight: undefined,
  capitals: undefined,
};

const EMPTY_FRAME: PartialFrame = {
  alignment: undefined,
  indentLeftTwips: undefined,
  indentRightTwips: undefined,
  indentFirstLineTwips: undefined,
  spaceBeforeTwips: undefined,
  automaticSpaceBefore: undefined,
  spaceAfterTwips: undefined,
  automaticSpaceAfter: undefined,
  lineTwips: undefined,
  lineRule: undefined,
  widowControl: undefined,
  keepNext: undefined,
  pageBreakBefore: undefined,
  contextualSpacing: undefined,
  tabStops: undefined,
  borders: NOTHING_STATED,
  fillColor: undefined,
};

const NO_NUMBERING: PartialNumbering = { numId: undefined, ilvl: undefined };

const mergeNumbering = (base: PartialNumbering, over: PartialNumbering): PartialNumbering => ({
  numId: over.numId ?? base.numId,
  ilvl: over.ilvl ?? base.ilvl,
});

const mergeFrames = (base: PartialFrame, over: PartialFrame): PartialFrame => ({
  alignment: over.alignment ?? base.alignment,
  indentLeftTwips: over.indentLeftTwips ?? base.indentLeftTwips,
  indentRightTwips: over.indentRightTwips ?? base.indentRightTwips,
  indentFirstLineTwips: over.indentFirstLineTwips ?? base.indentFirstLineTwips,
  spaceBeforeTwips: over.spaceBeforeTwips ?? base.spaceBeforeTwips,
  automaticSpaceBefore: over.automaticSpaceBefore ?? base.automaticSpaceBefore,
  spaceAfterTwips: over.spaceAfterTwips ?? base.spaceAfterTwips,
  automaticSpaceAfter: over.automaticSpaceAfter ?? base.automaticSpaceAfter,
  lineTwips: over.lineTwips ?? base.lineTwips,
  lineRule: over.lineRule ?? base.lineRule,
  widowControl: over.widowControl ?? base.widowControl,
  keepNext: over.keepNext ?? base.keepNext,
  pageBreakBefore: over.pageBreakBefore ?? base.pageBreakBefore,
  contextualSpacing: over.contextualSpacing ?? base.contextualSpacing,
  // Tab stops add to the ones already inherited rather than replacing them, which
  // is why a clear has to travel with them.
  tabStops:
    over.tabStops === undefined ? base.tabStops : [...(base.tabStops ?? []), ...over.tabStops],
  // Each side of a border stands or falls on its own, so a paragraph that states
  // one of them keeps the other three from its style.
  borders: {
    top: over.borders.top === undefined ? base.borders.top : over.borders.top,
    left: over.borders.left === undefined ? base.borders.left : over.borders.left,
    bottom: over.borders.bottom === undefined ? base.borders.bottom : over.borders.bottom,
    right: over.borders.right === undefined ? base.borders.right : over.borders.right,
  },
  fillColor: over.fillColor ?? base.fillColor,
});

function toLineRule(value: string | undefined): LineRule | undefined {
  if (value === "exact") return "exact";
  if (value === "atLeast") return "atLeast";
  if (value === "auto") return "auto";
  return undefined;
}

function toAlignment(value: string | undefined): ParagraphAlignment | undefined {
  if (value === "right" || value === "end") return "right";
  if (value === "center") return "center";
  if (value === "both" || value === "distribute") return "justify";
  if (value === "left" || value === "start") return "left";
  return undefined;
}

function twipsAttribute(element: XmlElement | null, name: string): number | undefined {
  if (element === null) return undefined;
  const value = statedNumber(attribute(element, W_NS, name));
  return Number.isFinite(value) ? value : undefined;
}

// Reads the paragraph properties a style element and a paragraph both spell the
// same way, so one function serves the whole cascade.
function readFrame(container: XmlElement | null): PartialFrame {
  const pPr = container === null ? null : firstNamed(container, W_NS, "pPr");
  if (pPr === null) return EMPTY_FRAME;

  const jc = firstNamed(pPr, W_NS, "jc");
  const indent = firstNamed(pPr, W_NS, "ind");
  const spacing = firstNamed(pPr, W_NS, "spacing");
  const hanging = twipsAttribute(indent, "hanging");

  return {
    alignment: toAlignment(jc === null ? undefined : attribute(jc, W_NS, "val")),
    indentLeftTwips: twipsAttribute(indent, "left") ?? twipsAttribute(indent, "start"),
    indentRightTwips: twipsAttribute(indent, "right") ?? twipsAttribute(indent, "end"),
    // A hanging indent is a first line pulled back out of the left indent, so the
    // two spellings are one number.
    indentFirstLineTwips: hanging === undefined ? twipsAttribute(indent, "firstLine") : -hanging,
    spaceBeforeTwips: twipsAttribute(spacing, "before"),
    automaticSpaceBefore: onOffAttribute(spacing, "beforeAutospacing"),
    spaceAfterTwips: twipsAttribute(spacing, "after"),
    automaticSpaceAfter: onOffAttribute(spacing, "afterAutospacing"),
    lineTwips: twipsAttribute(spacing, "line"),
    lineRule: toLineRule(spacing === null ? undefined : attribute(spacing, W_NS, "lineRule")),
    widowControl: onOff(pPr, "widowControl"),
    keepNext: onOff(pPr, "keepNext"),
    pageBreakBefore: onOff(pPr, "pageBreakBefore"),
    contextualSpacing: onOff(pPr, "contextualSpacing"),
    tabStops: readTabStops(pPr),
    borders: readBorders(pPr, "pBdr"),
    fillColor: readShading(pPr),
  };
}

function readTabStops(pPr: XmlElement): readonly TabStopEntry[] | undefined {
  const tabs = firstNamed(pPr, W_NS, "tabs");
  if (tabs === null) return undefined;

  const entries: TabStopEntry[] = [];
  for (const tab of childrenNamed(tabs, W_NS, "tab")) {
    const positionTwips = twipsAttribute(tab, "pos");
    if (positionTwips === undefined) continue;
    const value = attribute(tab, W_NS, "val");
    entries.push({
      positionTwips,
      alignment: toTabAlignment(value),
      kind: value === "clear" ? "clear" : "stop",
    });
  }
  return entries;
}

function toTabAlignment(value: string | undefined): TabAlignment {
  if (value === "center" || value === "right" || value === "decimal" || value === "bar") {
    return value;
  }
  // `num` lines a number up the way `left` does, and `start` and `end` are the
  // names the same two stops go under in a document written left to right.
  return value === "end" ? "right" : "left";
}

function readNumbering(container: XmlElement | null): PartialNumbering {
  const pPr = container === null ? null : firstNamed(container, W_NS, "pPr");
  const numPr = pPr === null ? null : firstNamed(pPr, W_NS, "numPr");
  if (numPr === null) return NO_NUMBERING;

  const numId = firstNamed(numPr, W_NS, "numId");
  const ilvl = firstNamed(numPr, W_NS, "ilvl");
  const level = ilvl === null ? Number.NaN : Number(attribute(ilvl, W_NS, "val"));

  return {
    numId: numId === null ? undefined : attribute(numId, W_NS, "val"),
    ilvl: Number.isInteger(level) ? level : undefined,
  };
}

const merge = (base: PartialMark, over: PartialMark): PartialMark => ({
  fontName: over.fontName ?? base.fontName,
  fontSizeHalfPoints: over.fontSizeHalfPoints ?? base.fontSizeHalfPoints,
  bold: over.bold ?? base.bold,
  italic: over.italic ?? base.italic,
  underline: over.underline ?? base.underline,
  verticalAlign: over.verticalAlign ?? base.verticalAlign,
  positionHalfPoints: over.positionHalfPoints ?? base.positionHalfPoints,
  color: over.color ?? base.color,
  characterSpacingTwentieths: over.characterSpacingTwentieths ?? base.characterSpacingTwentieths,
  characterScalePercent: over.characterScalePercent ?? base.characterScalePercent,
  highlight: over.highlight === undefined ? base.highlight : over.highlight,
  capitals: over.capitals ?? base.capitals,
  kernFromHalfPoints: over.kernFromHalfPoints ?? base.kernFromHalfPoints,
});

// Whether a style is the one a paragraph naming none is written in. `w:default` is
// an on/off attribute like any other, and a producer that is not Word writes it
// "true" or "on": reading only "1" finds no default style at all, and every
// paragraph naming none then loses its font, its size, its spacing and its numbering.
export const statesDefaultStyle = (style: XmlElement): boolean => {
  const value = attribute(style, W_NS, "default");
  return value !== undefined && statesOn(value);
};

// The same three answers an on/off element gives, spelled as an attribute: an
// automatic space is asked for on `w:spacing` rather than under it.
function onOffAttribute(element: XmlElement | null, name: string): boolean | undefined {
  if (element === null) return undefined;
  const value = attribute(element, W_NS, name);
  return value === undefined ? undefined : statesOn(value);
}

// An on/off property is on when it is there without a value, so only an explicit
// off turns it back off further down the cascade.
function onOff(container: XmlElement, name: string): boolean | undefined {
  const element = firstNamed(container, W_NS, name);
  return element === null ? undefined : toggledOn(element, W_NS);
}

function readThemeFonts(pkg: DocxPackage): ReadonlyMap<string, string> {
  if (!pkg.parts.has(THEME_PART)) return new Map();
  const scheme = descend(partXml(pkg, THEME_PART), ["themeElements", "fontScheme"]);
  const fonts = new Map<string, string>();
  for (const [slot, element] of [
    ["major", scheme === null ? null : firstNamed(scheme, A_NS, "majorFont")],
    ["minor", scheme === null ? null : firstNamed(scheme, A_NS, "minorFont")],
  ] as const) {
    if (element === null) continue;
    const latin = firstNamed(element, A_NS, "latin");
    const typeface = latin === null ? undefined : attribute(latin, "", "typeface");
    if (typeface !== undefined && typeface !== "") fonts.set(slot, typeface);
  }
  return fonts;
}

function descend(root: XmlElement, path: readonly string[]): XmlElement | null {
  let node: XmlElement | null = root;
  for (const name of path) {
    if (node === null) return null;
    node = firstNamed(node, A_NS, name);
  }
  return node;
}

function readMark(
  container: XmlElement | null,
  themeFonts: ReadonlyMap<string, string>,
): PartialMark {
  if (container === null) return EMPTY;
  const rPr = firstNamed(container, W_NS, "rPr");
  if (rPr === null) return EMPTY;

  const fonts = firstNamed(rPr, W_NS, "rFonts");
  let fontName: string | undefined;
  if (fonts !== null) {
    const explicit = attribute(fonts, W_NS, "ascii") ?? attribute(fonts, W_NS, "hAnsi");
    const themed = attribute(fonts, W_NS, "asciiTheme") ?? attribute(fonts, W_NS, "hAnsiTheme");
    fontName = explicit ?? (themed === undefined ? undefined : themeFonts.get(themeSlot(themed)));
  }

  const size = firstNamed(rPr, W_NS, "sz");
  const halfPoints = statedNumber(size === null ? undefined : attribute(size, W_NS, "val"));

  return {
    fontName,
    fontSizeHalfPoints: Number.isFinite(halfPoints) ? halfPoints : undefined,
    bold: onOff(rPr, "b"),
    italic: onOff(rPr, "i"),
    underline: underlineOf(rPr),
    verticalAlign: verticalAlignOf(rPr),
    positionHalfPoints: positionOf(rPr),
    color: colorOf(rPr),
    characterSpacingTwentieths: characterSpacingOf(rPr),
    characterScalePercent: characterScaleOf(rPr),
    kernFromHalfPoints: twipsAttribute(firstNamed(rPr, W_NS, "kern"), "val"),
    highlight: highlightOf(rPr),
    capitals: capitalsOf(rPr),
  };
}

/**
 * Whether a run names a face in its own `w:rPr`, rather than being handed one by
 * the cascade above it.
 *
 * **An equation turns on the difference.** The document's math font sets a math
 * run that names no face of its own, and a face the cascade handed down is not a
 * face the run named. Measured on 2026-08-24 off Word's own pdf of the only two
 * corpus documents this project refused: their math runs resolve to Times New
 * Roman through `docDefaults`, they state Cambria Math in `m:mathPr`, and both
 * pdfs embed Cambria Math. Asking the resolved mark instead read the cascade's
 * answer as the run's own, so the math font was never applied and the documents
 * were refused for a face that sets no equations.
 */
export function statesItsOwnFace(
  run: XmlElement,
  themeFonts: ReadonlyMap<string, string>,
): boolean {
  return readMark(run, themeFonts).fontName !== undefined;
}

// **Both stated is `w:caps`.** Measured 2026-08-13: a run stating the two came out
// exactly as one stating only `w:caps`, every letter a capital at the run's own size
// and the same 153.09pt of advance.
function capitalsOf(rPr: XmlElement): Capitals | undefined {
  const all = onOff(rPr, "caps");
  const small = onOff(rPr, "smallCaps");
  if (all === true) return "all";
  if (small === true) return "small";
  if (all === false || small === false) return "none";
  return undefined;
}

// **The sixteen colours Word paints a highlight in, read off its own pdf.** Measured
// 2026-08-13, one word highlighted in each: they are the plain and the dark web
// colours, and nothing about them is the document's to choose. `none` is a run
// turning an inherited highlight off and paints nothing.
const HIGHLIGHTS: Readonly<Record<string, string>> = {
  yellow: "#ffff00",
  green: "#00ff00",
  cyan: "#00ffff",
  magenta: "#ff00ff",
  blue: "#0000ff",
  red: "#ff0000",
  darkBlue: "#00008b",
  darkCyan: "#008b8b",
  darkGreen: "#006400",
  darkMagenta: "#800080",
  darkRed: "#8b0000",
  darkYellow: "#808000",
  darkGray: "#a9a9a9",
  lightGray: "#d3d3d3",
  black: "#000000",
  white: "#ffffff",
};

function highlightOf(rPr: XmlElement): string | null | undefined {
  const stated = firstNamed(rPr, W_NS, "highlight");
  if (stated === null) return undefined;
  const name = attribute(stated, W_NS, "val");
  if (name === undefined || name === "none") return null;
  return HIGHLIGHTS[name] ?? null;
}

// **How wide every glyph of the run is drawn, as a percentage of its own width.**
// Measured against Word on 2026-08-14 over one line of the same word repeated: a run
// scaled to 103 came out 102.84% as wide, to 90 89.90%, to 150 149.90% and to 50
// 49.89%, each a tenth of a percent short of its own multiple where Word rounds a
// scaled advance.
//
// **It scales the glyph's advance and not the letter spacing beside it.** The same
// run scaled to 150 with a point of spacing came out 253.63pt against 225.66 scaled
// alone and 150.54 plain: 225.66 and then 28 for the spacing, where spacing first
// and scaling after would have given 267.8.
function characterScaleOf(rPr: XmlElement): number | undefined {
  const scale = firstNamed(rPr, W_NS, "w");
  if (scale === null) return undefined;
  const value = Number(attribute(scale, W_NS, "val"));
  return Number.isFinite(value) && value > 0 ? value : undefined;
}

// Inside `pPr` the same name is the room above and below a paragraph, which is
// read elsewhere and states no `w:val` at all.
function characterSpacingOf(rPr: XmlElement): number | undefined {
  const element = firstNamed(rPr, W_NS, "spacing");
  const value = element === null ? undefined : attribute(element, W_NS, "val");
  if (value === undefined) return undefined;
  const twentieths = Number(value);
  return Number.isFinite(twentieths) ? twentieths : undefined;
}

// An underline is not a toggle: it names the kind of line to draw, and the
// cascade turns one off with "none" rather than by leaving it out.
function underlineOf(rPr: XmlElement): boolean | undefined {
  const element = firstNamed(rPr, W_NS, "u");
  if (element === null) return undefined;
  return (attribute(element, W_NS, "val") ?? "single") !== "none";
}

// How far off its own baseline the run is drawn, in the half-points Word states
// it in and with Word's sign: positive lifts. It is a distance and not a share of
// the size, so a run raised twelve stands six points up whatever it is set in.
function positionOf(rPr: XmlElement): number | undefined {
  const element = firstNamed(rPr, W_NS, "position");
  const value = element === null ? undefined : attribute(element, W_NS, "val");
  if (value === undefined) return undefined;
  const halfPoints = Number(value);
  return Number.isFinite(halfPoints) ? halfPoints : undefined;
}

function verticalAlignOf(rPr: XmlElement): VerticalAlign | undefined {
  const element = firstNamed(rPr, W_NS, "vertAlign");
  const value = element === null ? undefined : attribute(element, W_NS, "val");
  if (value === "superscript" || value === "subscript") return value;
  return value === undefined ? undefined : "baseline";
}

// "auto" leaves the colour to the page, which is not a colour this can name.
function colorOf(rPr: XmlElement): string | undefined {
  const element = firstNamed(rPr, W_NS, "color");
  const value = element === null ? undefined : attribute(element, W_NS, "val");
  return value === undefined || value === "auto" ? undefined : `#${value.replace("#", "")}`;
}

const themeSlot = (reference: string): string =>
  reference.startsWith("major") ? "major" : "minor";

export function readStyleTable(pkg: DocxPackage): StyleTable {
  const themeFonts = readThemeFonts(pkg);
  const numbering = readNumberingTable(pkg);
  const mathFont = readMathFont(pkg);
  if (!pkg.parts.has(STYLES_PART)) {
    return {
      byId: new Map(),
      defaultParagraphStyleId: undefined,
      docDefaults: EMPTY,
      docDefaultsFrame: EMPTY_FRAME,
      themeFonts,
      numbering,
      mathFont,
    };
  }

  const root = partXml(pkg, STYLES_PART);
  const defaults = firstNamed(root, W_NS, "docDefaults");
  const runDefault = defaults === null ? null : firstNamed(defaults, W_NS, "rPrDefault");

  const byId = new Map<string, StyleDefinition>();
  let defaultParagraphStyleId: string | undefined;

  for (const style of root.children) {
    if (style.namespace !== W_NS || style.name !== "style") continue;
    const id = attribute(style, W_NS, "styleId");
    if (id === undefined) continue;
    const basedOnElement = firstNamed(style, W_NS, "basedOn");
    byId.set(id, {
      id,
      basedOn: basedOnElement === null ? undefined : attribute(basedOnElement, W_NS, "val"),
      mark: readMark(style, themeFonts),
      frame: readFrame(style),
      numbering: readNumbering(style),
      tableBorders: readTableBorders(firstNamed(style, W_NS, "tblPr")),
      tableInsets: statedTableInsets(firstNamed(style, W_NS, "tblPr")),
      conditional: readConditionalFormats(style, themeFonts),
      rowBandSize: bandSizeOf(style, "tblStyleRowBandSize"),
      columnBandSize: bandSizeOf(style, "tblStyleColBandSize"),
    });
    const isParagraph = (attribute(style, W_NS, "type") ?? "paragraph") === "paragraph";
    if (isParagraph && statesDefaultStyle(style)) defaultParagraphStyleId = id;
  }

  return {
    byId,
    defaultParagraphStyleId,
    docDefaults: readMark(runDefault, themeFonts),
    docDefaultsFrame: readFrame(
      defaults === null ? null : firstNamed(defaults, W_NS, "pPrDefault"),
    ),
    themeFonts,
    numbering,
    mathFont,
  };
}

// A paragraph inside a table reads the table's own style as well, between the
// document's defaults and its own style. Word's hierarchy puts a table style below
// a paragraph style and above the defaults, and the difference is not academic:
// where a document leaves `Normal` empty and states its spacing in `docDefaults`,
// which is what Word itself writes, the table style is the only thing standing
// between the two and it decides the height of every row.
//
// Measured against Word's own pdf of a real document: `TableGrid` asks for no space
// after a paragraph and single line spacing, `docDefaults` for 8pt and 1.08, and
// leaving the table style out made every row 23.0pt tall against the 13.9pt Word
// drew. The error accumulates down the page, so nothing below the first row lands.
// Where a cell stands in its table, which is what decides which of the table
// style's conditional formats reach the paragraphs inside it. A cell in no table
// stands nowhere.
export type CellPosition = {
  readonly firstRow: boolean;
  readonly lastRow: boolean;
  readonly firstColumn: boolean;
  readonly lastColumn: boolean;
  // Which band the cell's row and its column fall in, or null where the table asks
  // for no banding on that axis.
  readonly rowBand: 1 | 2 | null;
  readonly columnBand: 1 | 2 | null;
};

export type InTable = {
  readonly styleId: string | null;
  // Null for a paragraph that reads the table's style without standing in a cell of
  // it, which is what asking about the table itself does.
  readonly at: CellPosition | null;
};

// **The order the conditional formats are applied in, each standing in front of the
// one before it**: the bands, then the edges, then the corners.
//
// Measured on 2026-08-10 by the authored `conditional-table` document, where each of
// the thirteen places states an indent of its own five points from its neighbours',
// so the left Word drew a cell's line at names the format that won it outright. A
// table with every switch on drew its interior at the **vertical** band's indent
// wherever both bands reached it, its first and last columns at those columns'
// indents over any band, its first and last rows at those rows' over the columns',
// and each of its four corners at that corner's over all of them.
//
// **A vertical band beats a horizontal one**, which is the way round nothing would
// guess: the interior of the five by four table came out at 5pt and 10pt, the two
// vertical bands, and never at the 15 or 20 the horizontal ones asked for.
const CONDITIONAL_ORDER = [
  "wholeTable",
  "band1Horz",
  "band2Horz",
  "band1Vert",
  "band2Vert",
  "firstCol",
  "lastCol",
  "firstRow",
  "lastRow",
  "nwCell",
  "neCell",
  "swCell",
  "seCell",
] as const;

function reaches(type: (typeof CONDITIONAL_ORDER)[number], at: CellPosition): boolean {
  switch (type) {
    case "wholeTable":
      return true;
    case "band1Vert":
      return at.columnBand === 1;
    case "band2Vert":
      return at.columnBand === 2;
    case "band1Horz":
      return at.rowBand === 1;
    case "band2Horz":
      return at.rowBand === 2;
    case "firstCol":
      return at.firstColumn;
    case "lastCol":
      return at.lastColumn;
    case "firstRow":
      return at.firstRow;
    case "lastRow":
      return at.lastRow;
    case "nwCell":
      return at.firstRow && at.firstColumn;
    case "neCell":
      return at.firstRow && at.lastColumn;
    case "swCell":
      return at.lastRow && at.firstColumn;
    case "seCell":
      return at.lastRow && at.lastColumn;
  }
}

const conditionalFormats = (
  table: StyleTable,
  inTable: InTable | null,
): readonly ConditionalFormat[] => {
  const at = inTable?.at;
  if (at === null || at === undefined) return [];
  const found: ConditionalFormat[] = [];
  for (const style of styleChain(table, inTable?.styleId ?? undefined))
    for (const type of CONDITIONAL_ORDER) {
      const format = style.conditional.get(type);
      if (format !== undefined && reaches(type, at)) found.push(format);
    }
  return found;
};

const framesOver = (table: StyleTable, inTable: InTable | null): PartialFrame => {
  let resolved = table.docDefaultsFrame;
  for (const style of styleChain(table, inTable?.styleId ?? undefined)) {
    resolved = mergeFrames(resolved, style.frame);
  }
  for (const format of conditionalFormats(table, inTable)) {
    resolved = mergeFrames(resolved, format.frame);
  }
  return resolved;
};

const marksOver = (table: StyleTable, inTable: InTable | null): PartialMark => {
  let resolved = table.docDefaults;
  for (const style of styleChain(table, inTable?.styleId ?? undefined)) {
    resolved = merge(resolved, style.mark);
  }
  for (const format of conditionalFormats(table, inTable)) {
    resolved = merge(resolved, format.mark);
  }
  return resolved;
};

// **A band's depth is the table style's business and not the table's**, which is
// where Word writes it and where the authored document's two styles differ: reading
// it off the table alone put every column of the wide-banded case in the band next
// to the one Word drew it in.
function bandSizeOf(style: XmlElement, name: string): number | undefined {
  const properties = firstNamed(style, W_NS, "tblPr");
  const held = properties === null ? null : firstNamed(properties, W_NS, name);
  const value = held === null ? Number.NaN : Number(attribute(held, W_NS, "val") ?? Number.NaN);
  return Number.isFinite(value) && value >= 1 ? Math.floor(value) : undefined;
}

// The depths a table written in this style bands at, which the table itself may
// state again over them.
export function resolveBandSizes(
  table: StyleTable,
  styleId: string | null,
): { readonly rowBandSize: number | undefined; readonly columnBandSize: number | undefined } {
  let rowBandSize: number | undefined;
  let columnBandSize: number | undefined;
  for (const style of styleChain(table, styleId ?? undefined)) {
    rowBandSize = style.rowBandSize ?? rowBandSize;
    columnBandSize = style.columnBandSize ?? columnBandSize;
  }
  return { rowBandSize, columnBandSize };
}

// A `w:tblStylePr` holds the same `w:pPr` and `w:rPr` a style itself does, under a
// `w:type` naming where it applies.
function readConditionalFormats(
  style: XmlElement,
  themeFonts: ReadonlyMap<string, string>,
): ReadonlyMap<string, ConditionalFormat> {
  const formats = new Map<string, ConditionalFormat>();
  for (const child of style.children) {
    if (child.namespace !== W_NS || child.name !== "tblStylePr") continue;
    const type = attribute(child, W_NS, "type");
    if (type === undefined) continue;
    formats.set(type, { mark: readMark(child, themeFonts), frame: readFrame(child) });
  }
  return formats;
}

function styleChain(table: StyleTable, styleId: string | undefined): readonly StyleDefinition[] {
  const chain: StyleDefinition[] = [];
  const seen = new Set<string>();
  let current = styleId;
  while (current !== undefined && !seen.has(current)) {
    seen.add(current);
    const style = table.byId.get(current);
    if (style === undefined) break;
    chain.unshift(style);
    current = style.basedOn;
  }
  return chain;
}

// Word sets a script at about two thirds of the run's size, which is what both
// reference faces carry as their own superscript size, and moves it off the
// baseline by a share of the size it was declared at. **The two shares are not the
// same one**: measured on 2026-08-07 by the authored `raised-text` document at
// 12pt, 24pt and 36pt, a superscript went up 4.08, 7.92 and 12.00 and a subscript
// went down 0.96, 2.40 and 3.60. Word writes those on a grid of 0.24pt, so the
// superscript is a third of the size and the subscript a tenth, each to the one
// step the smallest of them is out by.
const SCRIPT_SIZE = 0.65;
const SCRIPT_RAISE = 1 / 3;
const SCRIPT_DROP = 1 / 10;

const scriptRaiseOf = (align: VerticalAlign | undefined, fontSizePt: number): number => {
  if (align === "superscript") return fontSizePt * SCRIPT_RAISE;
  if (align === "subscript") return -fontSizePt * SCRIPT_DROP;
  return 0;
};

// The two raises add. Measured by the same document: a 12pt superscript raised
// twelve half-points was drawn 10.08pt off the baseline, which is the 4.08 Word
// lifts a superscript by and the 6 the run asked for.
const raiseOf = (resolved: PartialMark, fontSizePt: number): number =>
  scriptRaiseOf(resolved.verticalAlign, fontSizePt) + positionRaiseOf(resolved);

const positionRaiseOf = (resolved: PartialMark): number => (resolved.positionHalfPoints ?? 0) / 2;

function markOf(resolved: PartialMark): ParagraphMark {
  const declaredPt =
    resolved.fontSizeHalfPoints === undefined
      ? WORD_DEFAULT_FONT_SIZE_PT
      : Math.max(SMALLEST_FONT_SIZE_PT, resolved.fontSizeHalfPoints / 2);
  const scripted =
    resolved.verticalAlign === "superscript" || resolved.verticalAlign === "subscript";

  return {
    font:
      resolved.fontName === undefined
        ? { kind: "unresolved" }
        : { kind: "named", name: resolved.fontName },
    fontSizePt: scripted ? declaredPt * SCRIPT_SIZE : declaredPt,
    bold: resolved.bold ?? false,
    italic: resolved.italic ?? false,
    underline: resolved.underline ?? false,
    raisePt: raiseOf(resolved, declaredPt),
    lineSizePt: declaredPt,
    lineRaisePt: positionRaiseOf(resolved),
    color: resolved.color ?? null,
    characterSpacingPt: (resolved.characterSpacingTwentieths ?? 0) / TWIPS_PER_POINT,
    characterScale: (resolved.characterScalePercent ?? 100) / 100,
    // **Kerning is opt-in**, so a run stating nothing carries nothing here rather
    // than a nought: the two are the same answer to Word and telling them apart
    // costs nothing.
    kernFromHalfPoints: resolved.kernFromHalfPoints ?? null,
    highlight: resolved.highlight ?? null,
    capitals: resolved.capitals ?? "none",
  };
}

function paragraphMarkOf(
  paragraph: Paragraph,
  table: StyleTable,
  inTable: InTable | null,
): PartialMark {
  let resolved = marksOver(table, inTable);
  for (const style of styleChain(table, styleIdOf(paragraph, table))) {
    resolved = merge(resolved, style.mark);
  }
  return merge(resolved, readMark(firstNamed(paragraph.element, W_NS, "pPr"), table.themeFonts));
}

export const resolveParagraphMark = (
  paragraph: Paragraph,
  table: StyleTable,
  inTable: InTable | null = null,
): ParagraphMark => markOf(paragraphMarkOf(paragraph, table, inTable));

// The number is drawn in the paragraph's own mark except where its level says
// otherwise, which is how a bullet ends up in a symbol face at the text's size.
export const resolveNumberMark = (
  paragraph: Paragraph,
  table: StyleTable,
  level: NumberingLevel,
): ParagraphMark =>
  markOf(
    merge(paragraphMarkOf(paragraph, table, null), readMark(level.properties, table.themeFonts)),
  );

function runStyleChain(table: StyleTable, run: XmlElement): readonly StyleDefinition[] {
  const rPr = firstNamed(run, W_NS, "rPr");
  const rStyle = rPr === null ? null : firstNamed(rPr, W_NS, "rStyle");
  const id = rStyle === null ? undefined : attribute(rStyle, W_NS, "val");
  return styleChain(table, id);
}

export type MarkedRun = {
  readonly run: XmlElement;
  readonly mark: ParagraphMark;
};

export const resolveRunMarks = (
  paragraph: Paragraph,
  table: StyleTable,
): readonly ParagraphMark[] => resolveRuns(paragraph, table).map((marked) => marked.mark);

export function resolveRuns(
  paragraph: Paragraph,
  table: StyleTable,
  inTable: InTable | null = null,
): readonly MarkedRun[] {
  let inherited = marksOver(table, inTable);
  for (const style of styleChain(table, styleIdOf(paragraph, table))) {
    inherited = merge(inherited, style.mark);
  }

  return paragraphRuns(paragraph).map((run) => ({
    run,
    mark: resolveRunMark(run, inherited, table),
  }));
}

function resolveRunMark(run: XmlElement, inherited: PartialMark, table: StyleTable): ParagraphMark {
  let resolved = inherited;
  for (const style of runStyleChain(table, run)) resolved = merge(resolved, style.mark);
  return markOf(merge(resolved, readMark(run, table.themeFonts)));
}

// What a paragraph pasted out of a web page gets where it asks for its space
// rather than stating it. Measured on 2026-08-13, twelve cases three times each:
// baselines 13.8 apart go to 28.1 with an automatic space above, and the same
// fourteen points sit on top of a 24pt line, so it neither follows the face nor
// scales with it. It wins over a `w:before` stated beside it.
const AUTOMATIC_SPACE_TWIPS = 280;

export type ParagraphAlignment = "left" | "right" | "center" | "justify";

// Word's three ways of spelling a line's height: a multiple of the natural one,
// a floor under it, or a fixed height that replaces it.
export type LineRule = "auto" | "exact" | "atLeast";

export type ParagraphFrame = {
  readonly alignment: ParagraphAlignment;
  readonly indentLeftTwips: number;
  readonly indentRightTwips: number;
  readonly indentFirstLineTwips: number;
  readonly spaceBeforeTwips: number;
  readonly spaceAfterTwips: number;
  // Whether that room is the fourteen points a paragraph asks for automatically
  // rather than a value it states, which is dropped in places a stated one is
  // kept: against the top of what holds the paragraph, and between two paragraphs
  // of one list.
  readonly automaticSpaceBefore: boolean;
  readonly automaticSpaceAfter: boolean;
  readonly lineTwips: number | null;
  readonly lineRule: LineRule;
  // Whether Word holds the paragraph's first line off the foot of a page and its
  // last line off the top of the next one. On unless the cascade says otherwise.
  readonly widowControl: boolean;
  // Whether Word moves the paragraph onto the page its next one begins, where the
  // two would otherwise be split. Off unless the cascade says otherwise.
  readonly keepNext: boolean;
  // Whether the paragraph starts a page of its own. A paragraph already standing
  // at the top of one makes no empty page to get there.
  readonly pageBreakBefore: boolean;
  // Whether the space the paragraph asks for above and below it is dropped where
  // the paragraph on that side is of the same style.
  readonly contextualSpacing: boolean;
  // In ascending order, measured from the left edge of the text area rather than
  // from the paragraph's own indent.
  readonly tabStops: readonly TabStop[];
  // The lines drawn round the paragraph and the colour drawn behind it, both of
  // which take room the paragraph has to leave for them.
  readonly borders: Borders;
  readonly fillColor: string | null;
};

// The lines a table style asks for round the table it is set on, which whatever
// the table states itself stands in front of. Word's conditional formats, which
// dress the first row or the banded ones differently, are not read.
export function resolveTableBorders(table: StyleTable, styleId: string | null): TableBorders {
  let resolved = NO_TABLE_BORDERS;
  for (const style of styleChain(table, styleId ?? undefined)) {
    resolved = mergeTableBorders(resolved, style.tableBorders);
  }
  return resolved;
}

/**
 * What a table holds its cells' content off their walls by: what the table itself
 * states, then what its style chain states, then Word's own eighth of an inch.
 *
 * **A style stating a margin of nought is stating one.** Read on 2026-08-18 off a
 * corpus document whose table names `Table Grid`, based on a `TableNormal` the
 * document defines itself with all four margins at 0: Word draws the text of every
 * cell at the cell's own left edge, and this drew it 5.4pt in, which is the built-in
 * 108 twips standing where the style had already answered. The arithmetic closes to
 * a tenth: the cell begins at 346.2, its paragraph states `w:ind w:left="407"`, and
 * Word draws the line at 366.43 against 346.2 + 20.35.
 *
 * **The built-in default is the last resort and not the default.** A document with
 * no table styles at all reaches it, which is what every authored document does and
 * why they state their margins by hand.
 *
 * **Only the left and the right are taken from a style, and that is measured rather
 * than left out.** Taking all four was tried first and priced over the 47 corpus
 * documents whose tables leave their margins to a style: seven improved by about
 * half, three more improved, and **two came apart**, `c81e5b6f3818` from 10 wrong
 * cells to 6379 and `be3c786f733a` from 456 to 10865. Read on the page, both are the
 * same thing: their styles state 100 twips on all four sides, every one of their
 * rows grew 5pt at the head and 5pt at the foot, and their pages then broke early.
 * `c81e5b6f3818` page 8 came out **31.7pt low with every left still agreeing to a
 * tenth** and two of Word's page 7 lines pushed onto it, which is a height and not a
 * width. Leaving the top and the bottom to the table's own gives ten better, none
 * worse, and 1601 fewer wrong cells over the 47.
 *
 * **What that does not settle is why.** Both documents are written by a producer
 * that spells its twips with a decimal point, `w:w="100.0"`; refusing a measurement
 * written that way was the first reading tried and the corpus refuted it outright,
 * 13 documents worse and 133620 wrong cells against 77768, so Word plainly reads
 * `100.0` as 100. Whether Word takes no top margin from a style at all, or takes one
 * and spends it somewhere this project does not, is unmeasured: it wants an authored
 * document of a styled table asked of Word directly.
 */
export function resolveTableInsets(
  table: StyleTable,
  styleId: string | null,
  stated: StatedTableInsets,
): TableInsets {
  // The chain arrives with the style a style is based on in front of it, so each
  // one stands over what it was based on, and the table stands over all of them.
  let inherited = NO_TABLE_INSETS_STATED;
  for (const style of styleChain(table, styleId ?? undefined)) {
    inherited = {
      indentTwips: 0,
      leftTwips: style.tableInsets.leftTwips ?? inherited.leftTwips,
      rightTwips: style.tableInsets.rightTwips ?? inherited.rightTwips,
      // **A style's top and bottom margins are not taken**, which is measured and
      // not a simplification: see the note over this function.
      topTwips: null,
      bottomTwips: null,
    };
  }

  const side = (name: "leftTwips" | "rightTwips" | "topTwips" | "bottomTwips"): number =>
    stated[name] ?? inherited[name] ?? DEFAULT_TABLE_INSETS[name];

  return {
    indentTwips: stated.indentTwips,
    leftTwips: side("leftTwips"),
    rightTwips: side("rightTwips"),
    topTwips: side("topTwips"),
    bottomTwips: side("bottomTwips"),
  };
}

const NO_TABLE_INSETS_STATED: StatedTableInsets = {
  indentTwips: 0,
  leftTwips: null,
  rightTwips: null,
  topTwips: null,
  bottomTwips: null,
};

// A side stated as `nil` is an answer like any other, so what a table states
// stands in front of its style even where it asks for no line at all.
export const mergeTableBorders = (base: TableBorders, over: TableBorders): TableBorders => ({
  top: over.top === undefined ? base.top : over.top,
  left: over.left === undefined ? base.left : over.left,
  bottom: over.bottom === undefined ? base.bottom : over.bottom,
  right: over.right === undefined ? base.right : over.right,
  insideHorizontal:
    over.insideHorizontal === undefined ? base.insideHorizontal : over.insideHorizontal,
  insideVertical: over.insideVertical === undefined ? base.insideVertical : over.insideVertical,
});

export type ParagraphNumbering = {
  readonly numId: string;
  readonly ilvl: number;
  readonly level: NumberingLevel;
};

// numId zero is how a paragraph says it wants none of the numbering its style
// would otherwise hand it.
const NO_LIST = "0";

export function resolveParagraphNumbering(
  paragraph: Paragraph,
  table: StyleTable,
): ParagraphNumbering | null {
  let resolved = NO_NUMBERING;
  for (const style of styleChain(table, styleIdOf(paragraph, table))) {
    resolved = mergeNumbering(resolved, style.numbering);
  }
  resolved = mergeNumbering(resolved, readNumbering(paragraph.element));

  const { numId } = resolved;
  if (numId === undefined || numId === NO_LIST) return null;

  const ilvl = resolved.ilvl ?? 0;
  const level = numberingLevel(table.numbering, numId, ilvl);
  return level === null ? null : { numId, ilvl, level };
}

// Which style the paragraph is set in, which is what "the same style" means to
// the properties that ask about their neighbours.
export function styleIdOf(paragraph: Paragraph, table: StyleTable): string | undefined {
  const pPr = firstNamed(paragraph.element, W_NS, "pPr");
  const pStyle = pPr === null ? null : firstNamed(pPr, W_NS, "pStyle");
  const named = pStyle === null ? undefined : attribute(pStyle, W_NS, "val");
  return named ?? table.defaultParagraphStyleId;
}

// A list level's own indents sit above the style's and below the paragraph's, so
// a bulleted paragraph is indented without losing an indent it sets itself.
// **An equation with a paragraph to itself is centred, and the paragraph's own w:jc
// does not reach it.** Measured 2026-08-13 over five cases: the same equation came
// out at 287.36 of a body running 36 to 576 whether the paragraph stated nothing,
// `left` or `right`, and only `m:oMathParaPr/m:jc` moved it, to 36.00. An equation
// sharing its line with text is not this and is drawn where the flow puts it.
//
// The settings part's own `m:mathPr/m:defJc` is not read: no document here states
// one, and what Word does with it was not measured.
function displayEquationAlignment(paragraph: Paragraph): ParagraphAlignment | null {
  // **A run that draws nothing does not count, and a single space does.** Measured
  // on 2026-08-13: the same fraction beside an empty run came out centred and full
  // size, and beside one space came out in the flow at the script size.
  const runs = paragraphRuns(paragraph).filter(drawsInLine);
  if (runs.length === 0 || runs.some((run) => run.namespace !== MATH_NS)) return null;
  const stated = descendantsNamed(paragraph.element, MATH_NS, "jc")[0];
  const alignment = stated === undefined ? undefined : attribute(stated, MATH_NS, "val");
  return toAlignment(alignment) ?? "center";
}

export function resolveParagraphFrame(
  paragraph: Paragraph,
  table: StyleTable,
  inTable: InTable | null = null,
): ParagraphFrame {
  let resolved = framesOver(table, inTable);
  for (const style of styleChain(table, styleIdOf(paragraph, table))) {
    resolved = mergeFrames(resolved, style.frame);
  }

  const numbering = resolveParagraphNumbering(paragraph, table);
  if (numbering !== null) {
    resolved = mergeFrames(resolved, readFrame(numbering.level.properties));
  }
  resolved = mergeFrames(resolved, readFrame(paragraph.element));

  return {
    alignment: displayEquationAlignment(paragraph) ?? resolved.alignment ?? "left",
    indentLeftTwips: resolved.indentLeftTwips ?? 0,
    indentRightTwips: resolved.indentRightTwips ?? 0,
    indentFirstLineTwips: resolved.indentFirstLineTwips ?? 0,
    spaceBeforeTwips:
      resolved.automaticSpaceBefore === true
        ? AUTOMATIC_SPACE_TWIPS
        : (resolved.spaceBeforeTwips ?? 0),
    spaceAfterTwips:
      resolved.automaticSpaceAfter === true
        ? AUTOMATIC_SPACE_TWIPS
        : (resolved.spaceAfterTwips ?? 0),
    automaticSpaceBefore: resolved.automaticSpaceBefore === true,
    automaticSpaceAfter: resolved.automaticSpaceAfter === true,
    lineTwips: resolved.lineTwips ?? null,
    lineRule: resolved.lineRule ?? "auto",
    widowControl: resolved.widowControl ?? true,
    keepNext: resolved.keepNext ?? false,
    pageBreakBefore: resolved.pageBreakBefore ?? false,
    contextualSpacing: resolved.contextualSpacing ?? false,
    tabStops: settledStops(resolved.tabStops),
    borders: {
      top: resolved.borders.top ?? null,
      left: resolved.borders.left ?? null,
      bottom: resolved.borders.bottom ?? null,
      right: resolved.borders.right ?? null,
    },
    fillColor: resolved.fillColor ?? null,
  };
}

function settledStops(entries: readonly TabStopEntry[] | undefined): readonly TabStop[] {
  const stops = new Map<number, TabStop>();
  for (const entry of entries ?? []) {
    if (entry.kind === "clear") stops.delete(entry.positionTwips);
    else
      stops.set(entry.positionTwips, {
        positionTwips: entry.positionTwips,
        alignment: entry.alignment,
      });
  }
  return [...stops.values()].sort((left, right) => left.positionTwips - right.positionTwips);
}
