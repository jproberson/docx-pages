import { partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import { SETTINGS_PART } from "./settings.js";
import type { ParagraphMark } from "./styles.js";
import { attribute, descendantsNamed, firstNamed, toggledOn, type XmlElement } from "./xml.js";

// Where Word writes an equation, which is a language of its own rather than a run:
// its text is `m:t` inside `m:r`, so anything collecting `w:r` and reading `w:t`
// finds nothing at all in one and the paragraph holding it measures as empty.
export const MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

// What is read here is a run of text, a fraction and a delimiter, which is what the
// equations met in real documents are made of and nearly all of what they are made of.
// Anything else is refused whole and named, since drawing part of a structure is the
// plausible-looking page this project exists to avoid.
//
// **Reading a fraction is not setting one.** What this answers is the shape the file
// states; how tall a fraction stands against the line it is in, where its bar sits and
// how far a parenthesis stretches are questions only Word can answer, and until it has
// answered them an equation holding either still draws nothing and still says so.

// The face Word sets an equation in. A document states its own in `m:mathFont`, and
// Word writes that face onto every `m:r`'s `w:rPr` besides, so this answers for a run
// that names none of its own.
export const DEFAULT_MATH_FONT = "Cambria Math";

export function readMathFont(pkg: DocxPackage): string {
  if (!pkg.parts.has(SETTINGS_PART)) return DEFAULT_MATH_FONT;
  const properties = firstNamed(partXml(pkg, SETTINGS_PART), MATH_NS, "mathPr");
  const font = properties === null ? null : firstNamed(properties, MATH_NS, "mathFont");
  const named = font === null ? undefined : attribute(font, MATH_NS, "val");
  return named === undefined || named === "" ? DEFAULT_MATH_FONT : named;
}

// How Word sets the characters of a run inside an equation. **A run states this
// nowhere and is drawn slanted all the same**: no `m:r` in any equation met here
// carries a `w:i`, while the `m:ctrlPr` written beside it does, so the slant belongs
// to the equation rather than to the run's own properties and is the reader's to add.
export type MathStyle = "plain" | "bold" | "italic" | "bold-italic";

const STATED_STYLES: Readonly<Record<string, MathStyle>> = {
  p: "plain",
  b: "bold",
  i: "italic",
  bi: "bold-italic",
};

const WHERE_NOTHING_IS_STATED: MathStyle = "italic";

export type EquationRun = {
  readonly kind: "run";
  // The `m:r` itself, which holds its `w:rPr` exactly where a `w:r` holds one, so the
  // cascade that marks an ordinary run marks this one unchanged.
  readonly element: XmlElement;
  readonly text: string;
  // Null where the run states `m:nor`, Word's way of putting ordinary text inside an
  // equation: it keeps its own face and slant and is styled as an equation in no way.
  readonly style: MathStyle | null;
};

// A break inside an equation, which ends the line it stands on as a run's own break
// does. Word writes one as a `w:br` inside an `m:r`, and `element` is that run, since a
// break takes the height of the run holding it and is marked by the same cascade.
export type EquationBreak = {
  readonly kind: "break";
  readonly element: XmlElement;
};

// Word sets the two halves of a fraction one over the other with a bar between them,
// which is the whole of why an equation is not a line of runs.
export type EquationFraction = {
  readonly kind: "fraction";
  // The `m:ctrlPr` the file writes for the bar, which holds a `w:rPr` where a run holds
  // one and is marked by the same cascade. Null where the fraction states none.
  readonly control: XmlElement | null;
  readonly numerator: readonly EquationPiece[];
  readonly denominator: readonly EquationPiece[];
};

// What Word draws round a piece of an equation, and stretches to the height of it.
export type EquationDelimiter = {
  readonly kind: "delimiter";
  readonly control: XmlElement | null;
  // The characters drawn either side, and between one part and the next. Null where the
  // file states an empty one, which is how a delimiter open at one end is written.
  readonly opening: string | null;
  readonly closing: string | null;
  readonly separator: string | null;
  // **Null where the file states nothing**, which is every delimiter met in a real
  // document: whether Word stretches one that says nothing is unmeasured, so the reader
  // says what the file says and invents no answer of its own.
  readonly grows: boolean | null;
  readonly parts: readonly (readonly EquationPiece[])[];
};

export type EquationPiece = EquationRun | EquationBreak | EquationFraction | EquationDelimiter;

export type Equation =
  | { readonly kind: "read"; readonly content: readonly EquationPiece[] }
  // What was met and not read, named as the file names it and sorted, so a report can
  // say which equation is still missing and what is in the way of it.
  | { readonly kind: "refused"; readonly unreadable: readonly string[] };

export const equationsIn = (element: XmlElement): readonly XmlElement[] =>
  descendantsNamed(element, MATH_NS, "oMath");

export function readEquation(oMath: XmlElement): Equation {
  const unreadable = new Set<string>();
  const content = contentOf(oMath, unreadable);
  if (unreadable.size > 0) return { kind: "refused", unreadable: [...unreadable].sort() };
  return { kind: "read", content };
}

// The runs of an equation that is nothing but runs, and null for one holding anything
// that has to be set. **What is read is not yet what is drawn**: a fraction's own
// geometry is unmeasured, so an equation holding one is still drawn nowhere and the
// report still names it.
export function runsAlone(equation: Equation): readonly EquationRun[] | null {
  if (equation.kind === "refused") return null;
  const runs: EquationRun[] = [];
  for (const piece of equation.content) {
    if (piece.kind !== "run") return null;
    runs.push(piece);
  }
  return runs;
}

// How one run of an equation is set, asked of the run alone. `readEquation` answers
// the same for it, and this is for a caller holding the `m:r` and nothing else: the
// layout marks a run wherever it meets one rather than walking the equation again.
export const mathStyleOf = (run: XmlElement): MathStyle | null => {
  const properties = firstNamed(run, MATH_NS, "rPr");
  return properties === null ? WHERE_NOTHING_IS_STATED : styleStated(properties, new Set());
};

// What a mark of the paragraph's own cascade becomes once the equation has had its
// say. The face is the run's where the run names one, since Word writes the equation's
// face onto the run rather than leaving it to be inherited.
export function markedAsMath(
  mark: ParagraphMark,
  style: MathStyle | null,
  mathFont: string,
): ParagraphMark {
  if (style === null) return mark;
  return {
    ...mark,
    font: mark.font.kind === "named" ? mark.font : { kind: "named", name: mathFont },
    bold: style === "bold" || style === "bold-italic",
    italic: style === "italic" || style === "bold-italic",
  };
}

// What a paragraph writes round its text and draws nothing for. Word writes a
// bookmark or a proofing mark inside an equation as readily as outside one, and
// neither is a structure the reader has to understand.
const PLACES_NOTHING = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "proofErr",
  "commentRangeStart",
  "commentRangeEnd",
]);

const named = (element: XmlElement): string => {
  if (element.namespace === MATH_NS) return `m:${element.name}`;
  return element.namespace === W_NS ? `w:${element.name}` : element.name;
};

// Everything drawn inside one element of an equation, in the order the file writes it.
// An argument of a fraction or a delimiter holds the same content its equation does, so
// a fraction inside a delimiter inside a fraction is read by the one walk.
function contentOf(element: XmlElement, unreadable: Set<string>): readonly EquationPiece[] {
  const pieces: EquationPiece[] = [];
  for (const child of element.children) {
    if (child.namespace === MATH_NS) {
      if (child.name === "r") {
        readMathRun(child, pieces, unreadable);
        continue;
      }
      const piece = readPiece(child, unreadable);
      if (piece !== null) pieces.push(piece);
      continue;
    }
    if (child.namespace === W_NS && PLACES_NOTHING.has(child.name)) continue;
    unreadable.add(named(child));
  }
  return pieces;
}

function readPiece(element: XmlElement, unreadable: Set<string>): EquationPiece | null {
  if (element.name === "f") return readFraction(element, unreadable);
  if (element.name === "d") return readDelimiter(element, unreadable);
  unreadable.add(`m:${element.name}`);
  return null;
}

// Word sets a fraction stacked with a bar unless it says otherwise, and the three
// others it can say (skewed, linear and barless) are set differently enough that
// reading one as a stack would draw the wrong thing.
const STACKED_WITH_A_BAR = "bar";

function readFraction(element: XmlElement, unreadable: Set<string>): EquationFraction | null {
  let control: XmlElement | null = null;
  let numerator: readonly EquationPiece[] = [];
  let denominator: readonly EquationPiece[] = [];

  for (const child of element.children) {
    if (child.namespace !== MATH_NS) {
      if (!(child.namespace === W_NS && PLACES_NOTHING.has(child.name)))
        unreadable.add(named(child));
      continue;
    }
    switch (child.name) {
      case "fPr":
        control = fractionProperties(child, unreadable);
        break;
      case "num":
        numerator = contentOf(child, unreadable);
        break;
      case "den":
        denominator = contentOf(child, unreadable);
        break;
      default:
        unreadable.add(`m:${child.name}`);
    }
  }

  return { kind: "fraction", control, numerator, denominator };
}

function fractionProperties(properties: XmlElement, unreadable: Set<string>): XmlElement | null {
  let control: XmlElement | null = null;
  for (const child of properties.children) {
    if (child.namespace !== MATH_NS) {
      unreadable.add(named(child));
      continue;
    }
    switch (child.name) {
      case "ctrlPr":
        control = child;
        break;
      case "type":
        if (attribute(child, MATH_NS, "val") !== STACKED_WITH_A_BAR) unreadable.add("m:type");
        break;
      default:
        unreadable.add(`m:${child.name}`);
    }
  }
  return control;
}

// What Word draws either side of a delimiter, and between one part of it and the next,
// where the file states none of them itself.
const DEFAULT_OPENING = "(";
const DEFAULT_CLOSING = ")";
const DEFAULT_SEPARATOR = "|";

function readDelimiter(element: XmlElement, unreadable: Set<string>): EquationDelimiter | null {
  let opening: string | null = DEFAULT_OPENING;
  let closing: string | null = DEFAULT_CLOSING;
  let separator: string | null = DEFAULT_SEPARATOR;
  let grows: boolean | null = null;
  let control: XmlElement | null = null;
  const parts: (readonly EquationPiece[])[] = [];

  for (const child of element.children) {
    if (child.namespace !== MATH_NS) {
      if (!(child.namespace === W_NS && PLACES_NOTHING.has(child.name)))
        unreadable.add(named(child));
      continue;
    }
    if (child.name === "e") {
      parts.push(contentOf(child, unreadable));
      continue;
    }
    if (child.name !== "dPr") {
      unreadable.add(`m:${child.name}`);
      continue;
    }
    for (const stated of child.children) {
      if (stated.namespace !== MATH_NS) {
        unreadable.add(named(stated));
        continue;
      }
      switch (stated.name) {
        case "begChr":
          opening = characterStated(stated);
          break;
        case "endChr":
          closing = characterStated(stated);
          break;
        case "sepChr":
          separator = characterStated(stated);
          break;
        case "grow":
          grows = toggledOn(stated, MATH_NS);
          break;
        case "ctrlPr":
          control = stated;
          break;
        default:
          unreadable.add(`m:${stated.name}`);
      }
    }
  }

  return { kind: "delimiter", control, opening, closing, separator, grows, parts };
}

// A delimiter states the character it is drawn with in `m:val`, and an empty one to be
// drawn with nothing at all, which is how a bracket open at one end is written.
function characterStated(element: XmlElement): string | null {
  const value = attribute(element, MATH_NS, "val");
  return value === undefined || value === "" ? null : value;
}

// A run's text and its breaks, in the order they stand, so a break between two pieces
// of text ends the line between them rather than after both.
function readMathRun(run: XmlElement, pieces: EquationPiece[], unreadable: Set<string>): void {
  // What the run is set in is read before anything it holds, since a break in the
  // middle of one would otherwise leave the text in front of it styled by nothing.
  const properties = firstNamed(run, MATH_NS, "rPr");
  const style = properties === null ? WHERE_NOTHING_IS_STATED : styleStated(properties, unreadable);

  let text = "";
  const flush = (): void => {
    if (text !== "") pieces.push({ kind: "run", element: run, text, style });
    text = "";
  };

  for (const child of run.children) {
    if (child.namespace === MATH_NS && child.name === "t") {
      text += textOf(child);
      continue;
    }
    if (child.namespace === MATH_NS && child.name === "rPr") continue;
    // The run's own character properties, read by the cascade every run's are read by.
    if (child.namespace === W_NS && child.name === "rPr") continue;
    if (child.namespace === W_NS && PLACES_NOTHING.has(child.name)) continue;
    // A break ending the line is read; one starting a page or a column inside an
    // equation is a place nothing has measured, and is refused rather than guessed at.
    if (child.namespace === W_NS && child.name === "br") {
      if (attribute(child, W_NS, "type") === undefined) {
        flush();
        pieces.push({ kind: "break", element: run });
      } else unreadable.add("w:br");
      continue;
    }
    unreadable.add(named(child));
  }

  flush();
}

function styleStated(properties: XmlElement, unreadable: Set<string>): MathStyle | null {
  let style = WHERE_NOTHING_IS_STATED;
  let ordinaryText = false;
  for (const child of properties.children) {
    if (child.namespace !== MATH_NS) {
      unreadable.add(named(child));
      continue;
    }
    if (child.name === "nor") {
      ordinaryText = toggledOn(child, MATH_NS);
      continue;
    }
    if (child.name === "sty") {
      const stated = STATED_STYLES[attribute(child, MATH_NS, "val") ?? ""];
      if (stated === undefined) unreadable.add("m:sty");
      else style = stated;
      continue;
    }
    unreadable.add(`m:${child.name}`);
  }
  return ordinaryText ? null : style;
}

// The whitespace xml leaves insignificant, which is the only kind the edges of an
// `m:t` lose, by the rule a `w:t` is read by: Word writes `xml:space` on the `m:t`
// that ends an equation in a space, and that space is the equation's own.
const INSIGNIFICANT = /^[ \t\r\n]+|[ \t\r\n]+$/g;

const textOf = (element: XmlElement): string =>
  element.preservesSpace ? element.text : element.text.replace(INSIGNIFICANT, "");
