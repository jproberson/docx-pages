import type { MathStyle } from "./equations.js";

// **Word draws a math run in the Mathematical Alphanumeric block, not in a slanted
// face.** Measured 2026-08-13 by spelling one string six times over, once for each
// style, and reading back what the pdf says was drawn: `bandril` came out
// U+1D44F U+1D44E U+1D45B ... in Cambria Math regular, and our own advances for those
// code points sum to 36.94pt at 11pt against Word's 36.96.
//
// So the style is a mapping and not a weight. Nothing states bold or italic anywhere.

type Alphabet = {
  readonly upper: number;
  readonly lower: number;
  readonly upperGreek: number;
  readonly lowerGreek: number;
  // Null for a style Unicode holds no digits for, which the measurement bears out:
  // an italic run's digits were drawn as themselves, and a bold italic run's were
  // drawn as the bold ones.
  readonly digit: number | null;
};

const ALPHABETS: Readonly<Record<Exclude<MathStyle, "plain">, Alphabet>> = {
  bold: {
    upper: 0x1d400,
    lower: 0x1d41a,
    upperGreek: 0x1d6a8,
    lowerGreek: 0x1d6c2,
    digit: 0x1d7ce,
  },
  italic: { upper: 0x1d434, lower: 0x1d44e, upperGreek: 0x1d6e2, lowerGreek: 0x1d6fc, digit: null },
  "bold-italic": {
    upper: 0x1d468,
    lower: 0x1d482,
    upperGreek: 0x1d71c,
    lowerGreek: 0x1d736,
    digit: 0x1d7ce,
  },
};

// The one hole in the italic alphabet: U+1D455 is reserved, and Word draws the
// Planck constant in its place. Every other letter of every other style is where
// the offset says it is.
const ITALIC_H = 0x210e;

const within = (codePoint: number, first: number, last: number): boolean =>
  codePoint >= first && codePoint <= last;

function mappedTo(codePoint: number, alphabet: Alphabet, style: MathStyle): number {
  if (within(codePoint, 0x41, 0x5a)) return alphabet.upper + codePoint - 0x41;
  if (within(codePoint, 0x61, 0x7a)) {
    if (codePoint === 0x68 && style === "italic") return ITALIC_H;
    return alphabet.lower + codePoint - 0x61;
  }
  if (within(codePoint, 0x30, 0x39)) {
    return alphabet.digit === null ? codePoint : alphabet.digit + codePoint - 0x30;
  }
  // Greek runs Α to Ω and α to ω in one sequence in each math alphabet, so the
  // offset carries straight over, final sigma included.
  if (within(codePoint, 0x391, 0x3a9)) return alphabet.upperGreek + codePoint - 0x391;
  if (within(codePoint, 0x3b1, 0x3c9)) return alphabet.lowerGreek + codePoint - 0x3b1;
  return codePoint;
}

// What a math run spells once the style has had its say. `plain` and a run stating
// `m:nor` are drawn as they are written, which is the whole of what those two mean.
export function spelledAsMath(text: string, style: MathStyle | null): string {
  if (style === null || style === "plain") return text;
  const alphabet = ALPHABETS[style];
  let spelled = "";
  for (const character of text) {
    spelled += String.fromCodePoint(mappedTo(character.codePointAt(0) ?? 0, alphabet, style));
  }
  return spelled;
}
