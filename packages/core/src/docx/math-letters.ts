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

// **Word draws a hyphen in a maths run as a minus sign**, which is the same class of
// rule as the alphabet above: the character the file states is not the character Word
// draws. Measured 2026-08-14 by `equation-content-probe`, cases G and H, three repeats
// each: two letters with a hyphen between them and the same two with a minus came back
// out of Word's own pdf as the same string, `𝑎 − 𝑏`, drawn 25.454pt wide in both. The
// two characters are nothing alike in the face, 680 units against 1530, so a hyphen
// drawn as itself is 4.57pt short of what Word laid out at 11pt.
//
// A run stating `m:nor` is ordinary text and keeps its hyphen; the substitution is
// unmeasured for the other three styles, and stands for them because the style decides
// which alphabet the letters come out of and a hyphen is not a letter.
const DRAWN_AS: ReadonlyMap<number, number> = new Map([[0x2d, 0x2212]]);

// Whether Word draws this character out of one of the alphabets above, which is what
// the maths spacing asks to know either side of a gap: **the italic correction of a
// character stands unless the character after it comes out of the same alphabet.**
export const drawsFromAMathAlphabet = (codePoint: number): boolean =>
  codePoint === ITALIC_H || within(codePoint, 0x1d400, 0x1d7ff);

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

// What a math run spells once the style has had its say. `plain` keeps the letters it
// is written in, which is the whole of what that style means, and a run stating `m:nor`
// is not a maths run at all: it is drawn exactly as it is written.
export function spelledAsMath(text: string, style: MathStyle | null): string {
  if (style === null) return text;
  const alphabet = style === "plain" ? null : ALPHABETS[style];
  let spelled = "";
  for (const character of text) {
    const codePoint = character.codePointAt(0) ?? 0;
    const drawn = DRAWN_AS.get(codePoint);
    if (drawn !== undefined) spelled += String.fromCodePoint(drawn);
    else if (alphabet === null) spelled += character;
    else spelled += String.fromCodePoint(mappedTo(codePoint, alphabet, style));
  }
  return spelled;
}
