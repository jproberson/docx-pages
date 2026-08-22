import { partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import { attribute, childrenNamed, firstNamed, statedNumber, type XmlElement } from "./xml.js";

export const NUMBERING_PART = "word/numbering.xml";

// The formats a level can be counted in. Anything else is named rather than
// guessed at, so a list this cannot number says so instead of numbering it wrong.
export type NumberFormat =
  | "decimal"
  | "decimalZero"
  | "lowerLetter"
  | "upperLetter"
  | "lowerRoman"
  | "upperRoman"
  | "bullet"
  | "none"
  | "unsupported";

// What Word puts between the number and the text: a tab to the next stop, a
// single space, or nothing at all.
export type NumberSuffix = "tab" | "space" | "nothing";

export type LevelRestart =
  | { readonly kind: "any-higher" }
  | { readonly kind: "never" }
  | { readonly kind: "after-level"; readonly ilvl: number };

export type NumberingLevel = {
  readonly ilvl: number;
  readonly format: NumberFormat;
  // The number as a pattern, with %1 through %9 standing for the count at each
  // level; a bullet's pattern is the character itself.
  readonly text: string;
  readonly start: number;
  readonly restart: LevelRestart;
  readonly suffix: NumberSuffix;
  // A level spells its indents and its font the way a style does, so the cascade
  // reads them straight off this element.
  readonly properties: XmlElement;
};

export type NumberingTable = {
  readonly levels: ReadonlyMap<string, NumberingLevel>;
};

const EMPTY: NumberingTable = { levels: new Map() };

const levelKey = (numId: string, ilvl: number): string => `${numId}:${String(ilvl)}`;

export const numberingLevel = (
  table: NumberingTable,
  numId: string,
  ilvl: number,
): NumberingLevel | null => table.levels.get(levelKey(numId, ilvl)) ?? null;

const FORMATS = new Map<string, NumberFormat>([
  ["decimal", "decimal"],
  ["decimalZero", "decimalZero"],
  ["lowerLetter", "lowerLetter"],
  ["upperLetter", "upperLetter"],
  ["lowerRoman", "lowerRoman"],
  ["upperRoman", "upperRoman"],
  ["bullet", "bullet"],
  ["none", "none"],
]);

const formatOf = (value: string | undefined): NumberFormat =>
  (value === undefined ? undefined : FORMATS.get(value)) ?? "unsupported";

const suffixOf = (value: string | undefined): NumberSuffix =>
  value === "space" ? "space" : value === "nothing" ? "nothing" : "tab";

function valueOf(level: XmlElement, name: string): string | undefined {
  const element = firstNamed(level, W_NS, name);
  return element === null ? undefined : attribute(element, W_NS, "val");
}

function integerOr(value: string | undefined, fallback: number): number {
  const parsed = statedNumber(value);
  return Number.isInteger(parsed) ? parsed : fallback;
}

// w:lvlRestart counts levels from one, and zero means the count runs on however
// deep the list goes.
function restartOf(level: XmlElement): LevelRestart {
  const value = valueOf(level, "lvlRestart");
  if (value === undefined) return { kind: "any-higher" };
  const parsed = integerOr(value, -1);
  if (parsed === 0) return { kind: "never" };
  if (parsed < 0) return { kind: "any-higher" };
  return { kind: "after-level", ilvl: parsed - 1 };
}

function readLevel(element: XmlElement): NumberingLevel | null {
  const ilvl = integerOr(attribute(element, W_NS, "ilvl"), -1);
  if (ilvl < 0) return null;

  return {
    ilvl,
    format: formatOf(valueOf(element, "numFmt")),
    text: valueOf(element, "lvlText") ?? "",
    // **A level stating no start at all begins at nought, not at one.** Asked of
    // Word on 2026-08-22: three paragraphs of a level writing no `w:start` were
    // marked 0. 1. 2., the same as a level stating `w:val="0"`, where a level
    // stating `w:val="1"` was marked 1. 2. 3. Word's own lists all write the
    // attribute out, which is why this went unseen.
    start: integerOr(valueOf(element, "start"), 0),
    restart: restartOf(element),
    suffix: suffixOf(valueOf(element, "suff")),
    properties: element,
  };
}

function readAbstract(element: XmlElement): ReadonlyMap<number, NumberingLevel> {
  const levels = new Map<number, NumberingLevel>();
  for (const child of childrenNamed(element, W_NS, "lvl")) {
    const level = readLevel(child);
    if (level !== null) levels.set(level.ilvl, level);
  }
  return levels;
}

// An override either replaces a level outright or only moves where it starts
// counting, which leaves everything else the abstract definition said.
function readOverrides(
  num: XmlElement,
  base: ReadonlyMap<number, NumberingLevel>,
): ReadonlyMap<number, NumberingLevel> {
  const overridden = new Map<number, NumberingLevel>();

  for (const override of childrenNamed(num, W_NS, "lvlOverride")) {
    const ilvl = integerOr(attribute(override, W_NS, "ilvl"), -1);
    if (ilvl < 0) continue;

    const replacement = firstNamed(override, W_NS, "lvl");
    if (replacement !== null) {
      const level = readLevel(replacement);
      if (level !== null) overridden.set(ilvl, { ...level, ilvl });
      continue;
    }

    const start = firstNamed(override, W_NS, "startOverride");
    const inherited = base.get(ilvl);
    if (start === null || inherited === undefined) continue;
    overridden.set(ilvl, {
      ...inherited,
      start: integerOr(attribute(start, W_NS, "val"), inherited.start),
    });
  }

  return overridden;
}

export function readNumberingTable(pkg: DocxPackage): NumberingTable {
  if (!pkg.parts.has(NUMBERING_PART)) return EMPTY;
  const root = partXml(pkg, NUMBERING_PART);

  const abstracts = new Map<string, ReadonlyMap<number, NumberingLevel>>();
  for (const abstract of childrenNamed(root, W_NS, "abstractNum")) {
    const id = attribute(abstract, W_NS, "abstractNumId");
    if (id !== undefined) abstracts.set(id, readAbstract(abstract));
  }

  const levels = new Map<string, NumberingLevel>();
  for (const num of childrenNamed(root, W_NS, "num")) {
    const numId = attribute(num, W_NS, "numId");
    const reference = firstNamed(num, W_NS, "abstractNumId");
    const abstractId = reference === null ? undefined : attribute(reference, W_NS, "val");
    if (numId === undefined || abstractId === undefined) continue;

    const base = abstracts.get(abstractId);
    if (base === undefined) continue;

    for (const [ilvl, level] of base) levels.set(levelKey(numId, ilvl), level);
    for (const [ilvl, level] of readOverrides(num, base)) {
      levels.set(levelKey(numId, ilvl), level);
    }
  }

  return { levels };
}
