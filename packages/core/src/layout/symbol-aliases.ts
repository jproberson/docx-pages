// What a symbol face's characters mean, said in Unicode. A run written in
// Wingdings or Symbol does not hold text: it holds positions in that face's own
// page, stored either bare (0x6C) or on the private-use page Word writes them to
// (0xF06C), and only the face itself can draw them. On a machine without the
// face there is nothing to substitute: a text face asked for 0x6C draws the
// letter l where Word drew a bullet, which is worse than a missing-glyph box.
//
// Unicode names an equivalent for most of both pages, and translating first is
// how LibreOffice draws these documents too. The tables here are the curated
// middle of the published mappings: the bullets, checkboxes and arrows Word's
// own bullet library writes in Wingdings, and the Greek and operators
// equation-era documents write in Symbol. A character neither table carries
// falls through to the missing-glyph box as before. Webdings has not been met
// in a document and is left unmapped until it is.
//
// None of this claims Word's answer: what Word draws for a symbol face it does
// not have is a question about that machine (see `WORD_FALLBACK_FACES` on the
// unanswerable half of this), so the translation is only ever offered by the
// best-effort path, and the substitution it rides on is reported.

const WINGDINGS: ReadonlyMap<number, string> = new Map([
  [0x20, " "], // a space in a symbol run is still a space
  [0x6c, "●"],
  [0x6d, "❍"],
  [0x6e, "■"],
  [0x6f, "❑"],
  [0x70, "❒"],
  [0x75, "◆"],
  [0x76, "❖"],
  [0xa7, "▪"], // Word's stock square bullet
  [0xa8, "□"],
  [0xd8, "➢"],
  [0xdc, "➔"],
  [0xfc, "✓"],
  [0xfd, "☒"],
  [0xfe, "☑"],
]);

// Adobe's Symbol encoding, which the Symbol face keeps: Latin positions hold
// Greek, and the high half holds operators.
const SYMBOL: ReadonlyMap<number, string> = new Map([
  [0x20, " "],
  ...Array.from("ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ", (letter, at): [number, string] => [
    0x41 + at,
    letter,
  ]),
  ...Array.from("αβχδεφγηιϕκλμνοπθρστυϖωξψζ", (letter, at): [number, string] => [
    0x61 + at,
    letter,
  ]),
  [0xa3, "≤"],
  [0xa5, "∞"],
  [0xac, "←"],
  [0xad, "↑"],
  [0xae, "→"],
  [0xaf, "↓"],
  [0xb0, "°"],
  [0xb1, "±"],
  [0xb3, "≥"],
  [0xb4, "×"],
  [0xb6, "∂"],
  [0xb7, "•"],
  [0xb8, "÷"],
  [0xb9, "≠"],
  [0xbb, "≈"],
  [0xd5, "∏"],
  [0xd6, "√"],
  [0xd7, "⋅"],
  [0xe5, "∑"],
  [0xf2, "∫"],
]);

const FACES: ReadonlyMap<string, ReadonlyMap<number, string>> = new Map([
  ["wingdings", WINGDINGS],
  ["symbol", SYMBOL],
]);

const normalise = (name: string): string => name.trim().toLowerCase();

// Word stores a symbol character either bare or lifted onto the private-use
// page; both mean the same position in the face's own page.
const SYMBOL_PAGE = 0xf000;

const positionOf = (codePoint: number): number | null => {
  if (codePoint >= SYMBOL_PAGE && codePoint <= SYMBOL_PAGE + 0xff) return codePoint - SYMBOL_PAGE;
  return codePoint <= 0xff ? codePoint : null;
};

export const isAliasedSymbolFace = (faceName: string): boolean => FACES.has(normalise(faceName));

export function aliasedSymbolCharacter(faceName: string, codePoint: number): string | null {
  const page = FACES.get(normalise(faceName));
  if (page === undefined) return null;
  const position = positionOf(codePoint);
  return position === null ? null : (page.get(position) ?? null);
}

/**
 * The text a run written in a symbol face means, translated character by
 * character. A position the tables do not carry is lifted onto the private-use
 * page rather than left bare: bare, it would be painted as the substitute's own
 * letter, which is the one rendering worse than a box, and the box is what the
 * measurement gave it. Null for a face that is not a symbol face, whose text
 * already means itself.
 */
export function aliasedSymbolText(faceName: string, text: string): string | null {
  if (!isAliasedSymbolFace(faceName)) return null;
  return Array.from(text)
    .map((character) => {
      const codePoint = character.codePointAt(0) ?? 0;
      const alias = aliasedSymbolCharacter(faceName, codePoint);
      if (alias !== null) return alias;
      const position = positionOf(codePoint);
      return position === null ? character : String.fromCodePoint(SYMBOL_PAGE + position);
    })
    .join("");
}
