import { blockParagraphs, type Block } from "./blocks.js";
import { numberingLevel, type NumberFormat, type NumberingLevel } from "./numbering.js";
import { resolveParagraphNumbering, type StyleTable } from "./styles.js";

export type ParagraphNumber = {
  readonly text: string;
  readonly level: NumberingLevel;
};

// A format this cannot count in stops the whole part rather than putting a wrong
// number in front of a paragraph.
export type ParagraphNumbers =
  | { readonly kind: "numbered"; readonly numbers: ReadonlyMap<number, ParagraphNumber> }
  | {
      readonly kind: "unsupported";
      readonly paragraphIndex: number;
      readonly numId: string;
      readonly ilvl: number;
    };

// A list has nine levels, and only nine, however deeply it is nested.
const DEEPEST_LEVEL = 8;

type Counters = Map<number, number>;

// Counting runs per list instance: two numIds over the same abstract definition
// are two lists, each with its own count.
export function numberParagraphs(blocks: readonly Block[], styles: StyleTable): ParagraphNumbers {
  const numbers = new Map<number, ParagraphNumber>();
  const counted = new Map<string, Counters>();

  for (const paragraph of blockParagraphs(blocks)) {
    const numbering = resolveParagraphNumbering(paragraph, styles);
    if (numbering === null) continue;

    const { numId, ilvl, level } = numbering;
    if (level.format === "unsupported") {
      return { kind: "unsupported", paragraphIndex: paragraph.index, numId, ilvl };
    }

    const counters = counted.get(numId) ?? new Map<number, number>();
    counted.set(numId, counters);
    counters.set(ilvl, (counters.get(ilvl) ?? level.start - 1) + 1);
    restartBelow(styles, numId, ilvl, counters);

    numbers.set(paragraph.index, { text: textOf(styles, numId, level, counters), level });
  }

  return { kind: "numbered", numbers };
}

// A deeper level counts again from its start once the level above it moves on,
// unless it says which level restarts it or that nothing does.
function restartBelow(styles: StyleTable, numId: string, ilvl: number, counters: Counters): void {
  for (let deeper = ilvl + 1; deeper <= DEEPEST_LEVEL; deeper += 1) {
    const level = numberingLevel(styles.numbering, numId, deeper);
    if (level === null || level.restart.kind === "never") continue;
    if (level.restart.kind === "after-level" && level.restart.ilvl !== ilvl) continue;
    counters.delete(deeper);
  }
}

// %1 through %9 stand for the count at each level, each written in the format the
// level it names is counted in.
const PLACEHOLDER = /%([1-9])/g;

function textOf(
  styles: StyleTable,
  numId: string,
  level: NumberingLevel,
  counters: Counters,
): string {
  if (level.format === "none") return "";
  if (level.format === "bullet") return level.text;

  return level.text.replace(PLACEHOLDER, (_placeholder, digit: string) => {
    const at = Number(digit) - 1;
    const named = numberingLevel(styles.numbering, numId, at);
    return written(counters.get(at) ?? named?.start ?? 1, named?.format ?? "decimal");
  });
}

function written(count: number, format: NumberFormat): string {
  switch (format) {
    case "decimalZero":
      return count < 10 && count >= 0 ? `0${String(count)}` : String(count);
    case "lowerLetter":
      return letters(count);
    case "upperLetter":
      return letters(count).toUpperCase();
    case "lowerRoman":
      return roman(count).toLowerCase();
    case "upperRoman":
      return roman(count);
    case "decimal":
    case "none":
    case "bullet":
    case "unsupported":
      return String(count);
  }
}

const ALPHABET = "abcdefghijklmnopqrstuvwxyz";

// Word runs out of letters at z and repeats the next one instead of carrying, so
// the twenty-seventh item is aa and the twenty-eighth bb.
function letters(count: number): string {
  if (count < 1) return String(count);
  const letter = ALPHABET[(count - 1) % ALPHABET.length] ?? "";
  return letter.repeat(Math.floor((count - 1) / ALPHABET.length) + 1);
}

const ROMAN: readonly (readonly [number, string])[] = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
];

function roman(count: number): string {
  if (count < 1 || count > 3999) return String(count);

  let rest = count;
  let written = "";
  for (const [amount, glyph] of ROMAN) {
    while (rest >= amount) {
      written += glyph;
      rest -= amount;
    }
  }
  return written;
}
