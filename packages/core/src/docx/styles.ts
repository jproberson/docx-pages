import type { Paragraph } from "./blocks.js";
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
import { paragraphRuns } from "./paragraphs.js";
import { TWIPS_PER_POINT } from "../layout/units.js";
import { partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import { attribute, childrenNamed, firstNamed, type XmlElement } from "./xml.js";

export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

const STYLES_PART = "word/styles.xml";
const THEME_PART = "word/theme/theme1.xml";

export const WORD_DEFAULT_FONT_SIZE_PT = 10;

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
  // Null where the run leaves its colour to whatever it is drawn on.
  readonly color: string | null;
  // Extra width laid after every character of the run, the last one included:
  // `abcdef` at five points ran 33pt to 63pt, which is six characters' worth
  // rather than five gaps'. Negative tightens.
  readonly characterSpacingPt: number;
};

type PartialMark = {
  readonly fontName: string | undefined;
  readonly fontSizeHalfPoints: number | undefined;
  readonly bold: boolean | undefined;
  readonly italic: boolean | undefined;
  readonly underline: boolean | undefined;
  readonly verticalAlign: VerticalAlign | undefined;
  readonly color: string | undefined;
  readonly characterSpacingTwentieths: number | undefined;
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
  readonly spaceAfterTwips: number | undefined;
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

type StyleDefinition = {
  readonly id: string;
  readonly basedOn: string | undefined;
  readonly mark: PartialMark;
  readonly frame: PartialFrame;
  readonly numbering: PartialNumbering;
  // A table style says nothing about a paragraph's own frame; what it carries
  // here is the lines round the table it is set on.
  readonly tableBorders: TableBorders;
};

export type StyleTable = {
  readonly byId: ReadonlyMap<string, StyleDefinition>;
  readonly defaultParagraphStyleId: string | undefined;
  readonly docDefaults: PartialMark;
  readonly docDefaultsFrame: PartialFrame;
  readonly themeFonts: ReadonlyMap<string, string>;
  readonly numbering: NumberingTable;
};

const EMPTY: PartialMark = {
  fontName: undefined,
  fontSizeHalfPoints: undefined,
  bold: undefined,
  italic: undefined,
  underline: undefined,
  verticalAlign: undefined,
  color: undefined,
  characterSpacingTwentieths: undefined,
};

const EMPTY_FRAME: PartialFrame = {
  alignment: undefined,
  indentLeftTwips: undefined,
  indentRightTwips: undefined,
  indentFirstLineTwips: undefined,
  spaceBeforeTwips: undefined,
  spaceAfterTwips: undefined,
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
  spaceAfterTwips: over.spaceAfterTwips ?? base.spaceAfterTwips,
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
  const raw = attribute(element, W_NS, name);
  const value = raw === undefined ? Number.NaN : Number(raw);
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
    spaceAfterTwips: twipsAttribute(spacing, "after"),
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
  color: over.color ?? base.color,
  characterSpacingTwentieths: over.characterSpacingTwentieths ?? base.characterSpacingTwentieths,
});

// An on/off property is on when it is there without a value, so only an explicit
// off turns it back off further down the cascade.
function onOff(container: XmlElement, name: string): boolean | undefined {
  const element = firstNamed(container, W_NS, name);
  if (element === null) return undefined;
  const value = attribute(element, W_NS, "val");
  return value !== "0" && value !== "false" && value !== "off";
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
  const raw = size === null ? undefined : attribute(size, W_NS, "val");
  const halfPoints = raw === undefined ? undefined : Number(raw);

  return {
    fontName,
    fontSizeHalfPoints:
      halfPoints === undefined || !Number.isFinite(halfPoints) ? undefined : halfPoints,
    bold: onOff(rPr, "b"),
    italic: onOff(rPr, "i"),
    underline: underlineOf(rPr),
    verticalAlign: verticalAlignOf(rPr),
    color: colorOf(rPr),
    characterSpacingTwentieths: characterSpacingOf(rPr),
  };
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
  if (!pkg.parts.has(STYLES_PART)) {
    return {
      byId: new Map(),
      defaultParagraphStyleId: undefined,
      docDefaults: EMPTY,
      docDefaultsFrame: EMPTY_FRAME,
      themeFonts,
      numbering,
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
    });
    const isParagraph = (attribute(style, W_NS, "type") ?? "paragraph") === "paragraph";
    if (isParagraph && attribute(style, W_NS, "default") === "1") defaultParagraphStyleId = id;
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
  };
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

// Word sets a superscript or a subscript at about two thirds of the run's size,
// which is what both reference faces carry as their own superscript size, and
// moves it a third of that size off the baseline.
const SCRIPT_SIZE = 0.65;
const SCRIPT_RAISE = 1 / 3;

const raiseOf = (align: VerticalAlign | undefined, fontSizePt: number): number => {
  if (align === "superscript") return fontSizePt * SCRIPT_RAISE;
  if (align === "subscript") return -fontSizePt * SCRIPT_RAISE;
  return 0;
};

function markOf(resolved: PartialMark): ParagraphMark {
  const declaredPt =
    resolved.fontSizeHalfPoints === undefined
      ? WORD_DEFAULT_FONT_SIZE_PT
      : resolved.fontSizeHalfPoints / 2;
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
    raisePt: raiseOf(resolved.verticalAlign, declaredPt),
    color: resolved.color ?? null,
    characterSpacingPt: (resolved.characterSpacingTwentieths ?? 0) / TWIPS_PER_POINT,
  };
}

function paragraphMarkOf(paragraph: Paragraph, table: StyleTable): PartialMark {
  let resolved = table.docDefaults;
  for (const style of styleChain(table, styleIdOf(paragraph, table))) {
    resolved = merge(resolved, style.mark);
  }
  return merge(resolved, readMark(firstNamed(paragraph.element, W_NS, "pPr"), table.themeFonts));
}

export const resolveParagraphMark = (paragraph: Paragraph, table: StyleTable): ParagraphMark =>
  markOf(paragraphMarkOf(paragraph, table));

// The number is drawn in the paragraph's own mark except where its level says
// otherwise, which is how a bullet ends up in a symbol face at the text's size.
export const resolveNumberMark = (
  paragraph: Paragraph,
  table: StyleTable,
  level: NumberingLevel,
): ParagraphMark =>
  markOf(merge(paragraphMarkOf(paragraph, table), readMark(level.properties, table.themeFonts)));

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

export function resolveRuns(paragraph: Paragraph, table: StyleTable): readonly MarkedRun[] {
  let inherited = table.docDefaults;
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
export function resolveParagraphFrame(paragraph: Paragraph, table: StyleTable): ParagraphFrame {
  let resolved = table.docDefaultsFrame;
  for (const style of styleChain(table, styleIdOf(paragraph, table))) {
    resolved = mergeFrames(resolved, style.frame);
  }

  const numbering = resolveParagraphNumbering(paragraph, table);
  if (numbering !== null) {
    resolved = mergeFrames(resolved, readFrame(numbering.level.properties));
  }
  resolved = mergeFrames(resolved, readFrame(paragraph.element));

  return {
    alignment: resolved.alignment ?? "left",
    indentLeftTwips: resolved.indentLeftTwips ?? 0,
    indentRightTwips: resolved.indentRightTwips ?? 0,
    indentFirstLineTwips: resolved.indentFirstLineTwips ?? 0,
    spaceBeforeTwips: resolved.spaceBeforeTwips ?? 0,
    spaceAfterTwips: resolved.spaceAfterTwips ?? 0,
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
