import type { Paragraph } from "./paragraphs.js";
import { partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import { attribute, firstNamed, type XmlElement } from "./xml.js";

export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

const STYLES_PART = "word/styles.xml";
const THEME_PART = "word/theme/theme1.xml";

export const WORD_DEFAULT_FONT_SIZE_PT = 10;

export type FontChoice =
  { readonly kind: "named"; readonly name: string } | { readonly kind: "unresolved" };

export type ParagraphMark = {
  readonly font: FontChoice;
  readonly fontSizePt: number;
};

type PartialMark = {
  readonly fontName: string | undefined;
  readonly fontSizeHalfPoints: number | undefined;
};

type StyleDefinition = {
  readonly id: string;
  readonly basedOn: string | undefined;
  readonly mark: PartialMark;
};

export type StyleTable = {
  readonly byId: ReadonlyMap<string, StyleDefinition>;
  readonly defaultParagraphStyleId: string | undefined;
  readonly docDefaults: PartialMark;
  readonly themeFonts: ReadonlyMap<string, string>;
};

const EMPTY: PartialMark = { fontName: undefined, fontSizeHalfPoints: undefined };

const merge = (base: PartialMark, over: PartialMark): PartialMark => ({
  fontName: over.fontName ?? base.fontName,
  fontSizeHalfPoints: over.fontSizeHalfPoints ?? base.fontSizeHalfPoints,
});

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
  };
}

const themeSlot = (reference: string): string =>
  reference.startsWith("major") ? "major" : "minor";

export function readStyleTable(pkg: DocxPackage): StyleTable {
  const themeFonts = readThemeFonts(pkg);
  if (!pkg.parts.has(STYLES_PART)) {
    return { byId: new Map(), defaultParagraphStyleId: undefined, docDefaults: EMPTY, themeFonts };
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
    });
    const isParagraph = (attribute(style, W_NS, "type") ?? "paragraph") === "paragraph";
    if (isParagraph && attribute(style, W_NS, "default") === "1") defaultParagraphStyleId = id;
  }

  return {
    byId,
    defaultParagraphStyleId,
    docDefaults: readMark(runDefault, themeFonts),
    themeFonts,
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

export function resolveParagraphMark(paragraph: Paragraph, table: StyleTable): ParagraphMark {
  const pPr = firstNamed(paragraph.element, W_NS, "pPr");
  const pStyle = pPr === null ? null : firstNamed(pPr, W_NS, "pStyle");
  const named = pStyle === null ? undefined : attribute(pStyle, W_NS, "val");
  const styleId = named ?? table.defaultParagraphStyleId;

  let resolved = table.docDefaults;
  for (const style of styleChain(table, styleId)) resolved = merge(resolved, style.mark);
  resolved = merge(resolved, readMark(pPr, table.themeFonts));

  return {
    font:
      resolved.fontName === undefined
        ? { kind: "unresolved" }
        : { kind: "named", name: resolved.fontName },
    fontSizePt:
      resolved.fontSizeHalfPoints === undefined
        ? WORD_DEFAULT_FONT_SIZE_PT
        : resolved.fontSizeHalfPoints / 2,
  };
}
