import { effectOf, NO_EFFECT, type WrapDistances } from "./anchors.js";
import { isDetachedContent, type Paragraph } from "./blocks.js";
import { readDrawingTurn } from "./drawing.js";
import {
  equationsIn,
  mathStyleOf,
  MATH_NS,
  needsSetting,
  readEquation,
  runElementsOf,
  runsIn,
  runsOf,
  type Equation,
  type EquationPiece,
} from "./equations.js";
import type { MarkedMath } from "../layout/math.js";
import { spelledAsMath } from "./math-letters.js";
import { WP_NS } from "./inlines.js";
import { W_NS } from "./section.js";
import {
  resolveRuns,
  statesItsOwnFace,
  type MarkedRun,
  type ParagraphMark,
  type StyleTable,
} from "./styles.js";
import { holdsALegacyPicture, inlinePictureOf } from "./vml.js";
import { attribute, firstNamed, statedNumber, type XmlElement } from "./xml.js";

export type RunPiece =
  | { readonly kind: "text"; readonly text: string }
  | { readonly kind: "tab" }
  // A break ends the line it stands on, and one of type "page" starts the line
  // under it on a page of its own. One of type "column" sends the rest of its
  // paragraph to the top of the next column, whatever room is left below it.
  | { readonly kind: "break"; readonly endsPage: boolean; readonly endsColumn: boolean }
  | {
      readonly kind: "drawing";
      readonly widthEmu: number;
      readonly heightEmu: number;
      // The extent is the drawing the right way up, so how far round it was turned
      // is part of how much of the line it takes.
      readonly turnDegrees: number;
      // How far past that extent the drawing's own effects reach, which is room the
      // line keeps on all four sides. See `placeInlines`.
      readonly effect: WrapDistances;
    }
  // An equation that has to be set rather than laid along the line: a fraction or a
  // delimiter, with the runs it holds gathered back into the shape the reader found.
  // The line measures it through `setMath` and a renderer draws what
  // `mathPrimitivesOf` hands out, so nothing between here and the page knows a
  // fraction from a delimiter.
  | { readonly kind: "equation"; readonly content: readonly MarkedMath[] };

export type TextRun = {
  readonly mark: ParagraphMark;
  readonly pieces: readonly RunPiece[];
};

// The whitespace xml leaves insignificant, which is the only kind the edges of a
// w:t lose. A no-break space is a character the text is made of rather than
// whitespace around it, so it stays whether or not the run asks for it.
const INSIGNIFICANT = /^[ \t\r\n]+|[ \t\r\n]+$/g;

// Whitespace at the edges of a w:t is insignificant unless something asks for it,
// which is why Word writes xml:space wherever a space has to survive. It need not be
// the w:t that asks: the attribute stands for everything under the element stating
// it, and a document in the corpus states it once on its own root.
const textOf = (element: XmlElement): string =>
  element.preservesSpace ? element.text : element.text.replace(INSIGNIFICANT, "");

function extentOf(inline: XmlElement): RunPiece {
  const extent = firstNamed(inline, WP_NS, "extent");
  const size = (name: string): number => {
    const raw = extent === null ? undefined : attribute(extent, "", name);
    const value = statedNumber(raw);
    return Number.isFinite(value) ? value : 0;
  };
  return {
    kind: "drawing",
    widthEmu: size("cx"),
    heightEmu: size("cy"),
    turnDegrees: readDrawingTurn(inline),
    effect: effectOf(firstNamed(inline, WP_NS, "effectExtent")),
  };
}

function collectPieces(node: XmlElement, into: RunPiece[]): void {
  for (const child of node.children) {
    if (isDetachedContent(child)) continue;

    const isText = (child.namespace === W_NS || child.namespace === MATH_NS) && child.name === "t";
    if (isText) {
      const text = textOf(child);
      if (text !== "") into.push({ kind: "text", text });
      continue;
    }
    if (child.namespace === W_NS && child.name === "tab") {
      into.push({ kind: "tab" });
      continue;
    }
    if (child.namespace === W_NS && child.name === "br") {
      into.push({
        kind: "break",
        endsPage: attribute(child, W_NS, "type") === "page",
        endsColumn: attribute(child, W_NS, "type") === "column",
      });
      continue;
    }
    // A carriage return is the line ending a `w:br` of no type is, spelled the way
    // the older producers spell it. It carries no type, so it ends neither a page
    // nor a column.
    if (child.namespace === W_NS && child.name === "cr") {
      into.push({ kind: "break", endsPage: false, endsColumn: false });
      continue;
    }
    if (child.namespace === WP_NS && child.name === "inline") {
      into.push(extentOf(child));
      continue;
    }
    // A floating anchor is out of flow, so nothing under it belongs to this line.
    if (child.namespace === WP_NS && child.name === "anchor") continue;
    if (holdsALegacyPicture(child.namespace, child.name)) {
      const picture = inlinePictureOf(child);
      if (picture !== null) {
        into.push({
          kind: "drawing",
          widthEmu: picture.widthEmu,
          heightEmu: picture.heightEmu,
          turnDegrees: 0,
          // The old form states no effect extent at all, so its drawings ask for
          // nothing beyond the size their style gives them.
          effect: NO_EFFECT,
        });
      }
      continue;
    }

    collectPieces(child, into);
  }
}

// What an equation is set in: the document's own math font, and the theme the
// question "does this run name a face itself" has to be asked through.
type MathFaces = {
  readonly font: string;
  readonly themeFonts: ReadonlyMap<string, string>;
};

// A math run names its own face in its w:rPr wherever Word wrote it, and the
// document's math font is what one that names none is set in. **The run is asked,
// not the mark**: the mark carries whatever the cascade handed down, and a face
// out of `docDefaults` is not a face the run named. See `statesItsOwnFace`.
const inTheMathFont = (mark: ParagraphMark, run: XmlElement, faces: MathFaces): ParagraphMark =>
  statesItsOwnFace(run, faces.themeFonts)
    ? mark
    : { ...mark, font: { kind: "named", name: faces.font } };

/**
 * **Four fifths, rounded to the nearest half point**, which is what a small capital
 * is set at. Measured 2026-08-13 against Word's own pdf, by where the text after a
 * word of small capitals starts rather than by a size the pdf reports, which is
 * quantised on the way out:
 *
 * | the run states | four fifths | Word set them at |
 * | -------------- | ----------- | ---------------- |
 * | 11pt           | 8.8         | **9.0**          |
 * | 12pt           | 9.6         | **9.5**          |
 * | 13pt           | 10.4        | **10.5**         |
 * | 14pt           | 11.2        | **11.0**         |
 * | 20pt           | 16.0        | **16.0**         |
 * | 21pt           | 16.8        | **17.0**         |
 *
 * The 11pt and 13pt rows are what say it is the nearest half point rather than the
 * one below, and a half point is what `w:sz` itself is written in.
 *
 * The line does not shrink with them: the 20pt run's line was the 20pt line, so the
 * mark keeps the size it was declared at for measuring and carries the smaller one
 * only for drawing.
 */
const SMALL_CAPITAL = 0.8;

const smallCapitalSize = (sizePt: number): number => Math.round(sizePt * SMALL_CAPITAL * 2) / 2;

// A letter Word sets small: one with a capital of its own. `ß` is one of them, and
// it is drawn as itself, since its capital is two letters rather than one.
const isSmall = (character: string): boolean =>
  character !== character.toUpperCase() && character === character.toLowerCase();

// **Every space in a run of small capitals is set small**, wherever it stands.
// Measured 2026-08-13 by where the text after each space begins: all four in the
// case came out 2.15pt against the 2.71 of a 12pt space, the two standing after a
// digit and after a bracket included. A `(`, a `-`, a `)` and a digit are all drawn
// at the run's own size, so the space is the only character with no capital of its
// own that Word sets small.
const SPACE = " ";

// **A capital that is two letters is not made**: `ß` came back out of Word's pdf as
// `ß` under both `w:caps` and `w:smallCaps`, at the small size under the second.
const asCapital = (character: string): string => {
  const upper = character.toUpperCase();
  return String.fromCodePoint(upper.codePointAt(0) ?? 0) === upper ? upper : character;
};

const capitalsOfText = (text: string): string => {
  let drawn = "";
  for (const character of text) drawn += asCapital(character);
  return drawn;
};

/**
 * A run drawn as capitals, which is one run to Word and several to a pdf: it draws a
 * small capital at its own size, so the run comes apart wherever the size changes.
 * `w:caps` never comes apart, since every letter keeps the run's own size.
 */
function capitalised(mark: ParagraphMark, pieces: readonly RunPiece[]): readonly TextRun[] {
  if (mark.capitals === "none") return [{ mark, pieces }];
  if (mark.capitals === "all") {
    return [
      {
        mark,
        pieces: pieces.map((piece) =>
          piece.kind === "text" ? { kind: "text", text: capitalsOfText(piece.text) } : piece,
        ),
      },
    ];
  }

  const small: ParagraphMark = { ...mark, fontSizePt: smallCapitalSize(mark.fontSizePt) };
  const runs: TextRun[] = [];
  let heldSmall = false;
  const add = (piece: RunPiece, isSmallPiece: boolean): void => {
    const wanted = isSmallPiece ? small : mark;
    const last = runs[runs.length - 1];
    if (last !== undefined && last.mark === wanted) {
      runs[runs.length - 1] = { mark: wanted, pieces: [...last.pieces, piece] };
      return;
    }
    runs.push({ mark: wanted, pieces: [piece] });
  };

  for (const piece of pieces) {
    if (piece.kind !== "text") {
      add(piece, false);
      continue;
    }
    let held = "";
    for (const character of piece.text) {
      const smallHere = character === SPACE || isSmall(character);
      if (held !== "" && smallHere !== heldSmall) {
        add({ kind: "text", text: held }, heldSmall);
        held = "";
      }
      heldSmall = smallHere;
      held += asCapital(character);
    }
    if (held !== "") add({ kind: "text", text: held }, heldSmall);
  }
  return runs;
}

// **An equation that has to be set is gathered back into one piece here**, which is
// what makes it safe for `paragraphRuns` to hand out the runs standing inside a
// fraction: they are marked one at a time by the cascade, as every run is, and then
// put back into the shape the reader found before anything reaches a line. A line
// never sees a half.
//
// An equation of runs alone is left flat, which is what it was before there was any
// geometry: its runs are drawn where a paragraph's own runs are.
function equationsToSet(
  paragraph: Paragraph,
): ReadonlyMap<XmlElement, { readonly at: XmlElement; readonly equation: Equation }> {
  const held = new Map<XmlElement, { at: XmlElement; equation: Equation }>();
  for (const oMath of equationsIn(paragraph.element)) {
    const equation = readEquation(oMath);
    if (!needsSetting(equation)) continue;
    // **Every element the equation holds is mapped, the runs carrying a break among
    // them**, so that none is handed out a second time as a run of the paragraph's own.
    // What answers for the whole of it is the first run with text in it, whose mark is
    // the face and the size the geometry is set at; a run holding only a break states
    // neither.
    const first = runsOf(equation)[0];
    if (first === undefined) continue;
    for (const element of runElementsOf(equation))
      held.set(element, { at: first.element, equation });
  }
  return held;
}

// The equation's own tree, with the mark each of its runs resolved to hung on it.
function markedMathOf(
  pieces: readonly EquationPiece[],
  marks: ReadonlyMap<XmlElement, ParagraphMark>,
  faces: MathFaces,
): readonly MarkedMath[] {
  const marked: MarkedMath[] = [];
  for (const piece of pieces) {
    if (piece.kind === "break") continue;
    if (piece.kind === "run") {
      const mark = marks.get(piece.element);
      if (mark === undefined) continue;
      marked.push({
        kind: "run",
        text: spelledAsMath(piece.text, piece.style),
        mark: inTheMathFont(mark, piece.element, faces),
      });
      continue;
    }

    // A structure takes the mark of the `m:ctrlPr` the file writes for it, which is
    // where a bar's colour and size come from, and the mark of the first run it holds
    // where the file wrote none. **The run may be any depth down**: a delimiter round
    // a fraction holds no run of its own, and taking only the runs standing directly
    // inside left it with no mark and dropped the whole delimiter.
    const first = runsIn([piece])[0];
    const mark = first === undefined ? undefined : marks.get(first.element);
    if (first === undefined || mark === undefined) continue;
    const own = inTheMathFont(mark, first.element, faces);

    if (piece.kind === "fraction") {
      marked.push({
        kind: "fraction",
        mark: own,
        numerator: markedMathOf(piece.numerator, marks, faces),
        denominator: markedMathOf(piece.denominator, marks, faces),
      });
      continue;
    }
    marked.push({
      kind: "delimiter",
      mark: own,
      opening: piece.opening === null ? null : (piece.opening.codePointAt(0) ?? null),
      closing: piece.closing === null ? null : (piece.closing.codePointAt(0) ?? null),
      // **A file stating nothing means it grows**: a parenthesis round a fraction came
      // back on the fourth rung of Cambria Math's ladder where the same one with
      // `m:grow` turned off came back on the first.
      grows: piece.grows ?? true,
      content: piece.parts.flatMap((part) => markedMathOf(part, marks, faces)),
    });
  }
  return marked;
}

export function readRuns(paragraph: Paragraph, styles: StyleTable): readonly TextRun[] {
  const marked = resolveRuns(paragraph, styles);
  const toSet = equationsToSet(paragraph);
  if (toSet.size === 0) return flatRuns(marked, styles);

  const marks = new Map(marked.map((each) => [each.run, each.mark] as const));
  const runs: TextRun[] = [];
  for (const each of marked) {
    const equation = toSet.get(each.run);
    if (equation === undefined) {
      runs.push(...flatRuns([each], styles));
      continue;
    }
    // Every run of one equation answers with the same piece, emitted where the first
    // of them stands and passed over at the rest.
    if (each.run !== equation.at) continue;
    const faces = { font: styles.mathFont, themeFonts: styles.themeFonts };
    runs.push(
      ...equationRuns(
        equation.equation.kind === "read" ? equation.equation.content : [],
        // The equation's own mark, which is the face its constants and its sizes are
        // taken from and not only the face its characters are drawn in, so it goes
        // through the math font exactly as the runs inside it do.
        inTheMathFont(each.mark, each.run, faces),
        marks,
        faces,
      ),
    );
  }
  return runs;
}

/**
 * **A break inside an equation ends the line it stands on, and what follows it is an
 * equation of its own.** So the pieces either side of one are handed out as separate
 * equations with an ordinary break between them, and nothing past here has to know the
 * three came out of one `m:oMath`.
 *
 * Measured on 2026-08-14 by the authored `equation-break-probe` document, nine cases
 * three times each against Word's own pdf. A display fraction alone is 27.36; the same
 * with a break after it is 42.00, with a break before it 40.08, with a run after the
 * break 40.08, and with a second fraction after the break 54.48, which is what the same
 * two equations with an ordinary `w:r` break between them come to as well. So the break
 * is ordinary on the content side: it ends the line, what follows stands on the next,
 * and it costs the same wherever in the equation it stands.
 *
 * **Which run holds the break is what its own line is measured from**, so it is handed
 * out under the mark of the `m:r` carrying it rather than under the equation's. The
 * probe separates that from the mark's line: see `heldOpenPt` in `layout/lines.ts`.
 *
 * A break standing inside a fraction's half or a delimiter is not this and is not
 * honoured: `markedMathOf` passes one over and `fidelity.ts` names the document.
 */
function equationRuns(
  content: readonly EquationPiece[],
  mark: ParagraphMark,
  marks: ReadonlyMap<XmlElement, ParagraphMark>,
  faces: MathFaces,
): readonly TextRun[] {
  const runs: TextRun[] = [];
  let held: EquationPiece[] = [];

  const flush = (): void => {
    const marked = markedMathOf(held, marks, faces);
    if (marked.length > 0) runs.push({ mark, pieces: [{ kind: "equation", content: marked }] });
    held = [];
  };

  for (const piece of content) {
    if (piece.kind !== "break") {
      held.push(piece);
      continue;
    }
    flush();
    runs.push({
      mark: marks.get(piece.element) ?? mark,
      // A break inside an equation that starts a page or a column is refused by the
      // reader rather than guessed at, so the one that reaches here ends a line only.
      pieces: [{ kind: "break", endsPage: false, endsColumn: false }],
    });
  }
  flush();
  return runs;
}

function flatRuns(marked: readonly MarkedRun[], styles: StyleTable): readonly TextRun[] {
  return marked.flatMap((each) => {
    const pieces: RunPiece[] = [];
    collectPieces(each.run, pieces);
    if (each.run.namespace !== MATH_NS) return capitalised(each.mark, pieces);

    const style = mathStyleOf(each.run);
    return capitalised(
      inTheMathFont(each.mark, each.run, {
        font: styles.mathFont,
        themeFonts: styles.themeFonts,
      }),
      pieces.map((piece) =>
        piece.kind === "text" ? { kind: "text", text: spelledAsMath(piece.text, style) } : piece,
      ),
    );
  });
}
