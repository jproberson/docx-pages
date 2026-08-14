import { isDetachedContent, type Paragraph } from "./blocks.js";
import { readDrawingTurn } from "./drawing.js";
import { mathStyleOf, MATH_NS } from "./equations.js";
import { spelledAsMath } from "./math-letters.js";
import { WP_NS } from "./inlines.js";
import { W_NS } from "./section.js";
import { resolveRuns, type ParagraphMark, type StyleTable } from "./styles.js";
import { holdsALegacyPicture, inlinePictureOf } from "./vml.js";
import { attribute, firstNamed, type XmlElement } from "./xml.js";

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
    };

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
    const value = raw === undefined ? Number.NaN : Number(raw);
    return Number.isFinite(value) ? value : 0;
  };
  return {
    kind: "drawing",
    widthEmu: size("cx"),
    heightEmu: size("cy"),
    turnDegrees: readDrawingTurn(inline),
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
        });
      }
      continue;
    }

    collectPieces(child, into);
  }
}

// A math run names its own face in its w:rPr wherever Word wrote it, and the
// document's math font is what one that names none is set in.
const inTheMathFont = (mark: ParagraphMark, mathFont: string): ParagraphMark =>
  mark.font.kind === "named" ? mark : { ...mark, font: { kind: "named", name: mathFont } };

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

export function readRuns(paragraph: Paragraph, styles: StyleTable): readonly TextRun[] {
  return resolveRuns(paragraph, styles).flatMap((marked) => {
    const pieces: RunPiece[] = [];
    collectPieces(marked.run, pieces);
    if (marked.run.namespace !== MATH_NS) return capitalised(marked.mark, pieces);

    const style = mathStyleOf(marked.run);
    return capitalised(
      inTheMathFont(marked.mark, styles.mathFont),
      pieces.map((piece) =>
        piece.kind === "text" ? { kind: "text", text: spelledAsMath(piece.text, style) } : piece,
      ),
    );
  });
}
