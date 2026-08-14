import type {
  AdvanceTable,
  FontMetrics,
  GlyphInk,
  InkBox,
  InkTable,
  MathConstants,
  MathTable,
  MathVariant,
} from "./font-metrics.js";
import type { ParagraphMark } from "../docx/styles.js";

// Where a fraction and a delimiter are set, measured against Word on 2026-08-13 over
// eleven authored pages and read off Word's own pdf.
//
// **Every number here is a length in points, measured from the box's own baseline with
// up positive.** That is the way the MATH table states its own constants, and the only
// way this arithmetic stays readable; a renderer walking down the page turns it round
// once, at the edge.
//
// **Word's pdf reports on a grid of 0.24pt** (1/300 inch), heights and thicknesses
// included, so a measurement quoted below is worth one step either way and no more. The
// constants the tests hold were derived by inverting these rules over the 11pt case,
// which is the one measured from every side; they are what Cambria Math's MATH table
// has to hold for Word's numbers to come out, and `font-file.ts` reading the real table
// is what will say whether it does.

// A face an equation can be set in: one carrying the MATH table it states its own
// constants in, the outlines a fraction's height is measured off, and the advances
// everything is laid along. It is the four tables `readFontFile` answers with, narrowed
// once, so that every function below is total and pure.
export type MathFace = {
  readonly unitsPerEm: number;
  readonly constants: MathConstants;
  readonly advanceOf: (codePoint: number) => number | null;
  readonly inkOf: GlyphInk;
  readonly inkOfGlyph: (glyph: number) => InkBox | null;
  // In the order the face keeps them, smallest first, the character's own plain glyph
  // usually the first of them. Empty for a character the face does not grow.
  readonly tallerVariantsOf: (codePoint: number) => readonly MathVariant[];
};

// **A face missing any of the three cannot set an equation at all**, and this answers
// null for one rather than inventing what the missing table would have said: all but a
// handful of faces state no MATH table, and a fraction set off another face's constants
// is a plausible-looking page rather than Word's.
export function mathFace(read: {
  readonly metrics: FontMetrics;
  readonly advances: AdvanceTable;
  readonly ink: InkTable;
  readonly math: MathTable;
}): MathFace | null {
  const { advances, ink, math } = read;
  if (math.kind !== "math" || ink.kind !== "ink" || advances.kind !== "advances") return null;

  return {
    unitsPerEm: read.metrics.unitsPerEm,
    constants: math.constants,
    advanceOf: advances.advanceFor,
    inkOf: ink.inkOf,
    inkOfGlyph: ink.inkOfGlyph,
    tallerVariantsOf: math.tallerVariantsOf,
  };
}

// Which of the two sets of constants a structure is set with. **Word sets a fraction
// standing alone in its paragraph in the display one and every other fraction in the
// text one, one standing inside another fraction included**: measured over a nested
// pair, where the inner fraction kept gaps of 0.72pt against the outer one's 1.20, and
// 0.72 is the rule thickness while 1.20 is a constant of its own.
export type MathSetting = "display" | "text";

// What a piece of an equation covers. The width is what it advances the line by; the
// ascent and the descent are the ink, since **a fraction's height is measured off the
// ink of its halves and not off the face's ascent**: the same fraction with an `l` in
// its numerator and without came out 2.64pt apart in height.
export type MathBox = {
  readonly widthPt: number;
  readonly ascentPt: number;
  readonly descentPt: number;
};

export type PlacedMathBox = MathBox & {
  readonly leftPt: number;
  // How far this box's own baseline stands above the baseline of the box holding it.
  readonly baselinePt: number;
};

export type FractionBar = {
  readonly leftPt: number;
  readonly widthPt: number;
  // The top of the bar above the fraction's baseline. It is filled downwards from
  // there: Word fills the bar rather than stroking it, which is how `pdf/fills.ts`
  // reads it back as a rectangle.
  readonly topPt: number;
  readonly thicknessPt: number;
};

export type FractionBox = MathBox & {
  readonly numerator: PlacedMathBox;
  readonly denominator: PlacedMathBox;
  readonly bar: FractionBar;
};

export type FractionRequest = {
  readonly numerator: MathBox;
  readonly denominator: MathBox;
  // **The fraction's own size, which is not its halves'.** Word shrinks the halves of a
  // fraction that shares its line with ordinary text and leaves the bar and the axis at
  // the size the run states: a fraction stated at 11pt beside text drew its halves at
  // 7.9 and kept the full-size 0.72pt bar.
  readonly sizePt: number;
  readonly setting: MathSetting;
  readonly face: MathFace;
};

// **Word's own arithmetic runs on a grid of 0.24pt**, which is a three-hundredth of
// an inch, and the bar is the one place measured where the grid shows in a length
// rather than in a position. Measured over the same fraction at seven sizes against
// Cambria's rule thickness of 133/2048: 9pt asks 0.584 and Word drew 0.480, 10pt asks
// 0.649 and drew 0.720, 13 asks 0.844 and drew 0.960, 15 asks 0.974 and drew 0.960,
// 17 asks 1.104 and drew 1.200, 20 asks 1.299 and drew 1.200. **The nearest multiple
// every time**, rounding up in four and down in three, so it is neither a floor nor a
// ceiling.
const DEVICE_GRID_PT = 0.24;

const onTheGrid = (lengthPt: number): number =>
  Math.round(lengthPt / DEVICE_GRID_PT) * DEVICE_GRID_PT;

// How far a fraction's box stands out past its own bar on each side. Measured over
// two fractions of very different widths, both at 11pt: the bar of the fraction above
// each stood 2.16pt wider and 1.08 further left than the bar below it, so what the
// upper bar spans is this box and not the text.
//
// **It keeps to the em rather than to a length.** The same nested fraction at 20pt put
// the inner bar 56.40 wide at 277.92 and the outer 60.24 at 275.76, which is 1.92 out
// on each side: a share of the em asks 1.964 and the 0.24 grid rounds that to exactly
// the 1.92 drawn, where a fixed length would have asked 1.08 again.
const BESIDE_THE_BAR_EM = 1.08 / 11;

export function fractionBox(request: FractionRequest): FractionBox {
  const { numerator, denominator, sizePt, setting, face } = request;
  const constants = face.constants;
  const scale = sizePt / face.unitsPerEm;
  const display = setting === "display";

  const thicknessPt = onTheGrid(constants.fractionRuleThickness * scale);
  const axisPt = constants.axisHeight * scale;
  const barTopPt = axisPt + thicknessPt / 2;
  const barBottomPt = axisPt - thicknessPt / 2;

  const gapAbovePt =
    (display ? constants.fractionNumDisplayStyleGapMin : constants.fractionNumeratorGapMin) * scale;
  const gapBelowPt =
    (display ? constants.fractionDenomDisplayStyleGapMin : constants.fractionDenominatorGapMin) *
    scale;

  // The face's own shift, unless the half is deep enough that holding to it would close
  // the gap the face asks to be left round the bar.
  const shiftUpPt = Math.max(
    (display
      ? constants.fractionNumeratorDisplayStyleShiftUp
      : constants.fractionNumeratorShiftUp) * scale,
    barTopPt + gapAbovePt + numerator.descentPt,
  );
  const shiftDownPt = Math.max(
    (display
      ? constants.fractionDenominatorDisplayStyleShiftDown
      : constants.fractionDenominatorShiftDown) * scale,
    denominator.ascentPt - barBottomPt + gapBelowPt,
  );

  // The bar is drawn as wide as the wider half and the halves are centred on one
  // another, measured over halves of two characters against ten: their centres came
  // back 0.03pt apart. **The box stands out past the bar**, which is what the bar of a
  // fraction above it spans.
  const barWidthPt = Math.max(numerator.widthPt, denominator.widthPt);
  const besidePt = BESIDE_THE_BAR_EM * sizePt;
  const widthPt = barWidthPt + 2 * besidePt;
  const centred = (box: MathBox): number => (widthPt - box.widthPt) / 2;

  return {
    widthPt,
    ascentPt: shiftUpPt + numerator.ascentPt,
    descentPt: shiftDownPt + denominator.descentPt,
    numerator: { ...numerator, leftPt: centred(numerator), baselinePt: shiftUpPt },
    denominator: { ...denominator, leftPt: centred(denominator), baselinePt: -shiftDownPt },
    bar: { leftPt: besidePt, widthPt: barWidthPt, topPt: barTopPt, thicknessPt },
  };
}

// What the halves of a fraction are set at where it shares its line with ordinary text.
// **Only that shrinks them**: a fraction alone in its paragraph kept its halves at the
// stated 11pt, one beside text drew them at 7.9, and one standing in another fraction's
// numerator kept the full 11 and only tightened its gaps. Where the boundary between
// those lies is unmeasured, and is the caller's to decide.
export const scriptSizePt = (sizePt: number, face: MathFace): number =>
  (sizePt * face.constants.scriptPercentScaleDown) / 100;

// What a line holding a fraction takes beyond the ink of the fraction itself. Measured
// over the three fractions of the probe, whose paragraphs stood 1.59, 1.42 and 1.55pt
// above the ink of their boxes: the face's own `mathLeading` is 1.62 at the size Word
// drew them, which is inside the grid the pdf reports on for all three. **How a line is
// measured is the line's own business**; this is here because the answer is a constant
// of the MATH table and nothing else reads one.
export const mathLeadingPt = (sizePt: number, face: MathFace): number =>
  (face.constants.mathLeading * sizePt) / face.unitsPerEm;

// One of the two characters a delimiter is drawn with, and which of the face's shapes
// for it was drawn. `variant` is the rung of the face's own ladder, and **a rung is a
// glyph with no character of its own**: a renderer that can only draw text cannot draw
// one. Null where the face keeps no ladder for the character at all, and the
// character's own glyph is what is drawn.
export type PlacedDelimiter = MathBox & {
  readonly codePoint: number;
  readonly variant: MathVariant | null;
  // **Whether what is drawn is a shape of the face's own rather than the character.**
  // The first rung of a ladder is the character's own glyph, so a delimiter that grew
  // no further than that is drawn as text like any other character, which is what Word
  // drew for one round a run and for one told not to grow. Anything above the first
  // rung has no character to name it at all.
  readonly grown: boolean;
  readonly leftPt: number;
  readonly baselinePt: number;
};

export type DelimiterBox = MathBox & {
  readonly opening: PlacedDelimiter | null;
  readonly closing: PlacedDelimiter | null;
  readonly content: PlacedMathBox;
  // **Whether what stands between the two is tall enough to be set as one.** Below the
  // face's own `delimitedSubFormulaMinHeight` Word draws the characters as ordinary
  // text on the line's own baseline and grows neither: a parenthesis round a run came
  // back in one string with the run, at the run's baseline, while the same parenthesis
  // round a fraction was hung on the axis 0.48pt off it.
  readonly setAsASubFormula: boolean;
  // Whether the content stands taller than the tallest rung the face keeps, which Word
  // fills by assembling the pieces it states beside the ladder and nothing here draws.
  // The tallest rung stands in, so the delimiter is drawn short.
  readonly grownShort: boolean;
};

export type DelimiterRequest = {
  // Null where the file states an empty character, which is a bracket open at that end
  // and draws nothing at all.
  readonly opening: number | null;
  readonly closing: number | null;
  readonly content: MathBox;
  readonly sizePt: number;
  // Whether the delimiter grows to its content. **A file stating nothing means it
  // does**: a parenthesis round a fraction came back on the fourth rung of Cambria
  // Math's ladder where the same one with `m:grow` turned off came back on the first.
  readonly grows: boolean;
  readonly face: MathFace;
};

export function delimiterBox(request: DelimiterRequest): DelimiterBox {
  const { opening, closing, content, sizePt, grows, face } = request;
  const scale = sizePt / face.unitsPerEm;
  const axisPt = face.constants.axisHeight * scale;

  const setAsASubFormula =
    content.ascentPt + content.descentPt >= face.constants.delimitedSubFormulaMinHeight * scale;

  // **A delimiter is measured about the axis rather than about the baseline**, since
  // that is where it is hung: every rung of Cambria Math's ladder above the first has
  // its own ink centred on the axis to the unit, and the first is 88 units off it,
  // which is the 0.48pt Word hung it by.
  const reachPt = 2 * Math.max(content.ascentPt - axisPt, content.descentPt + axisPt);

  let leftPt = 0;
  const place = (codePoint: number | null): PlacedDelimiter | null => {
    if (codePoint === null) return null;
    const drawn = !setAsASubFormula
      ? asOrdinaryText(codePoint, sizePt, face)
      : hungOnTheAxis(codePoint, grows ? reachPt : 0, sizePt, face);
    const placed = { ...drawn, codePoint, leftPt };
    leftPt += drawn.widthPt;
    return placed;
  };

  const drawnOpening = place(opening);
  const contentLeftPt = leftPt;
  leftPt += content.widthPt;
  const drawnClosing = place(closing);
  const placedContent = { ...content, leftPt: contentLeftPt, baselinePt: 0 };
  const grownShort =
    setAsASubFormula &&
    grows &&
    [opening, closing].some(
      (codePoint) => codePoint !== null && standsShortOf(codePoint, reachPt, sizePt, face),
    );

  const standing = [placedContent, drawnOpening, drawnClosing].filter(
    (each): each is PlacedMathBox | PlacedDelimiter => each !== null,
  );

  return {
    widthPt: leftPt,
    ascentPt: Math.max(...standing.map((each) => each.baselinePt + each.ascentPt)),
    descentPt: Math.max(...standing.map((each) => each.descentPt - each.baselinePt)),
    opening: drawnOpening,
    closing: drawnClosing,
    content: placedContent,
    setAsASubFormula,
    grownShort,
  };
}

// How far a box's ink reaches below its baseline, which is the other way up from the
// way a face states it. A box that stops on the baseline reaches nothing, rather than
// the negative nothing turning the sign over leaves behind.
const belowPt = (bottomUnits: number, scale: number): number =>
  bottomUnits === 0 ? 0 : -bottomUnits * scale;

type DrawnDelimiter = MathBox & {
  readonly variant: MathVariant | null;
  readonly grown: boolean;
  readonly baselinePt: number;
};

// The character drawn as itself, on the line's own baseline, which is what Word draws
// round anything too short to be set as a sub-formula.
function asOrdinaryText(codePoint: number, sizePt: number, face: MathFace): DrawnDelimiter {
  const scale = sizePt / face.unitsPerEm;
  const ink = face.inkOf(codePoint);
  return {
    variant: null,
    grown: false,
    widthPt: (face.advanceOf(codePoint) ?? 0) * scale,
    ascentPt: ink === null ? 0 : ink.top * scale,
    descentPt: ink === null ? 0 : belowPt(ink.bottom, scale),
    baselinePt: 0,
  };
}

// **The tallest rung that does not overhang what it stands round**, hung so that its own
// ink is centred on the axis. A `reachPt` of nothing asks for the first rung, which is
// the character's own shape and is what a delimiter told not to grow is drawn as.
//
// **The question of which rule this is, is closed, and here is why so that nobody
// re-derives it.** Three rules fitted the first measurement: this one, the content's
// height less a fixed drop of five points, and the height less a fixed factor of
// itself. Four delimiters have since been measured, round a run, round a shallow
// fraction, round an ordinary one and round a fraction of a fraction, reaching 1.087,
// 1.157, 1.181 and 1.0025 of the rung Word picked.
//
// **The factor is dead.** Round the nested fraction the content reaches 48.00 and Word
// drew the eighth rung at 47.88; a factor of 0.84 asks for 40.32 and would have picked
// the seventh at 41.04.
//
// **The drop cannot be told from this one by any document.** They differ only where the
// content reaches more than five points past the rung below it, which at that rung is a
// ratio above 1.229, and the deepest content a document can hold reaches 1.181: a
// fraction's halves are bounded by the ink of the letters, and the tallest and deepest
// Cambria Math carries in its italic alphabet are 1436 and -447 units. Sizes do not
// help, since the ladder is stated in the em and scales with everything else.
//
// So the two agree on every page there is, and this is the one of them that states no
// constant of its own.
function hungOnTheAxis(
  codePoint: number,
  reachPt: number,
  sizePt: number,
  face: MathFace,
): DrawnDelimiter {
  const scale = sizePt / face.unitsPerEm;
  const variants = face.tallerVariantsOf(codePoint);
  const withinReach = variants.filter((each) => each.measurement * scale <= reachPt);
  const chosen = withinReach[withinReach.length - 1] ?? variants[0];
  if (chosen === undefined) return asOrdinaryText(codePoint, sizePt, face);
  // The first rung is the character's own glyph, so a delimiter that reached no
  // further than it has not grown at all.
  const grown = chosen !== variants[0];

  const axisPt = face.constants.axisHeight * scale;
  const ink = chosen.ink ?? face.inkOfGlyph(chosen.glyph);
  if (ink === null) {
    return {
      variant: chosen,
      grown,
      widthPt: chosen.advance * scale,
      ascentPt: 0,
      descentPt: 0,
      baselinePt: 0,
    };
  }

  return {
    variant: chosen,
    grown,
    widthPt: chosen.advance * scale,
    ascentPt: ink.top * scale,
    descentPt: belowPt(ink.bottom, scale),
    baselinePt: axisPt - ((ink.top + ink.bottom) / 2) * scale,
  };
}

// Whether the content stands taller than every shape the face keeps for the character.
// One the face grows through no ladder at all stands short of nothing: it is drawn as
// itself, which is all Word has for it either.
function standsShortOf(
  codePoint: number,
  reachPt: number,
  sizePt: number,
  face: MathFace,
): boolean {
  const variants = face.tallerVariantsOf(codePoint);
  const tallest = variants[variants.length - 1];
  return tallest !== undefined && (tallest.measurement * sizePt) / face.unitsPerEm < reachPt;
}

// The box a run of an equation covers, which is its advance and the ink of the
// characters in it. A caller measuring a run its own way, with the kerning and the
// substitutions the rest of the layout knows about, hands the answer straight to
// `fractionBox` instead; this is here because the ink is what a fraction's height is
// made of and nothing else in the tree gathers it.
export function textBox(text: string, sizePt: number, face: MathFace): MathBox {
  const scale = sizePt / face.unitsPerEm;
  let advance = 0;
  let top: number | null = null;
  let bottom: number | null = null;

  for (const character of text) {
    const codePoint = character.codePointAt(0);
    if (codePoint === undefined) continue;
    advance += face.advanceOf(codePoint) ?? 0;
    const ink = face.inkOf(codePoint);
    if (ink === null) continue;
    top = top === null ? ink.top : Math.max(top, ink.top);
    bottom = bottom === null ? ink.bottom : Math.min(bottom, ink.bottom);
  }

  return {
    widthPt: advance * scale,
    ascentPt: (top ?? 0) * scale,
    descentPt: belowPt(bottom ?? 0, scale),
  };
}

// **What a set equation comes to on the page, and the whole of what a renderer is
// told.** Everything above answers where a thing stands; this says what the things
// are, so that nothing downstream has to know a fraction from a delimiter. A renderer
// draws these three in the order they arrive and has each of them already: a run of
// text, a filled rectangle and a glyph named by number.
//
// `baselinePt` and `topPt` run **down** the page from the origin the caller gives,
// which is the line's own baseline, because that is the way a renderer walks a page.
// Everything else in this file measures up from a baseline, and this is the one place
// the sign turns over.
export type MathPrimitive =
  | {
      readonly kind: "text";
      readonly text: string;
      readonly mark: ParagraphMark;
      readonly leftPt: number;
      readonly baselinePt: number;
      readonly sizePt: number;
    }
  | {
      readonly kind: "fill";
      readonly mark: ParagraphMark;
      readonly leftPt: number;
      readonly topPt: number;
      readonly widthPt: number;
      readonly heightPt: number;
    }
  // A glyph the face keeps for a character that grows, which has no character of its
  // own: a renderer that can only draw text cannot draw one of these.
  | {
      readonly kind: "glyph";
      readonly glyph: number;
      readonly mark: ParagraphMark;
      readonly leftPt: number;
      readonly baselinePt: number;
      readonly sizePt: number;
    };

// A piece of an equation that has been measured and placed, which is the geometry
// above with what is drawn hung on it. A caller builds these from the leaves up,
// asking `textBox`, `fractionBox` and `delimiterBox` for the geometry as it goes.
export type SetMath =
  | {
      readonly kind: "run";
      readonly text: string;
      // What the cascade resolved for the `m:r`, once `markedAsMath` has had its say.
      readonly mark: ParagraphMark;
      readonly sizePt: number;
      readonly box: MathBox;
    }
  | {
      readonly kind: "fraction";
      // What the bar is drawn in, which is the `m:ctrlPr`'s own mark where the file
      // wrote one and the fraction's own otherwise: a bar takes its colour from there.
      readonly mark: ParagraphMark;
      readonly box: FractionBox;
      readonly numerator: readonly SetMath[];
      readonly denominator: readonly SetMath[];
    }
  | {
      readonly kind: "delimiter";
      readonly mark: ParagraphMark;
      readonly sizePt: number;
      readonly box: DelimiterBox;
      readonly content: readonly SetMath[];
    };

export type MathOrigin = {
  readonly leftPt: number;
  // The line's own baseline, which the equation's own baseline sits on.
  readonly baselinePt: number;
};

/**
 * Everything a set equation draws, in the order it is painted.
 *
 * Nothing overlaps anything else in an equation, so the order is free and this walks
 * it the way it reads: a fraction's numerator, then its bar, then its denominator; a
 * delimiter's opening, then what it holds, then its closing.
 */
export function mathPrimitivesOf(
  pieces: readonly SetMath[],
  at: MathOrigin,
): readonly MathPrimitive[] {
  const primitives: MathPrimitive[] = [];
  let leftPt = at.leftPt;

  for (const piece of pieces) {
    primitivesInto(piece, { leftPt, baselinePt: at.baselinePt }, primitives);
    leftPt += widthOfSet(piece);
  }
  return primitives;
}

const widthOfSet = (piece: SetMath): number => piece.box.widthPt;

function primitivesInto(piece: SetMath, at: MathOrigin, into: MathPrimitive[]): void {
  if (piece.kind === "run") {
    into.push({
      kind: "text",
      text: piece.text,
      mark: piece.mark,
      leftPt: at.leftPt,
      baselinePt: at.baselinePt,
      sizePt: piece.sizePt,
    });
    return;
  }

  if (piece.kind === "fraction") {
    const { box } = piece;
    // A half's own baseline stands above the fraction's, and the page runs the other
    // way, so the placement's rise becomes a fall here.
    const half = (side: readonly SetMath[], placed: PlacedMathBox): void => {
      let leftPt = at.leftPt + placed.leftPt;
      for (const each of side) {
        primitivesInto(each, { leftPt, baselinePt: at.baselinePt - placed.baselinePt }, into);
        leftPt += widthOfSet(each);
      }
    };

    half(piece.numerator, box.numerator);
    into.push({
      kind: "fill",
      mark: piece.mark,
      leftPt: at.leftPt + box.bar.leftPt,
      topPt: at.baselinePt - box.bar.topPt,
      widthPt: box.bar.widthPt,
      heightPt: box.bar.thicknessPt,
    });
    half(piece.denominator, box.denominator);
    return;
  }

  const { box } = piece;
  const side = (placed: PlacedDelimiter | null): void => {
    if (placed === null) return;
    const baselinePt = at.baselinePt - placed.baselinePt;
    // A delimiter that grew is a glyph of the face's own with no character to name
    // it; one that did not is the character itself and goes down as text.
    if (!placed.grown || placed.variant === null) {
      into.push({
        kind: "text",
        text: String.fromCodePoint(placed.codePoint),
        mark: piece.mark,
        leftPt: at.leftPt + placed.leftPt,
        baselinePt,
        sizePt: piece.sizePt,
      });
      return;
    }
    into.push({
      kind: "glyph",
      glyph: placed.variant.glyph,
      mark: piece.mark,
      leftPt: at.leftPt + placed.leftPt,
      baselinePt,
      sizePt: piece.sizePt,
    });
  };

  side(box.opening);
  let leftPt = at.leftPt + box.content.leftPt;
  for (const each of piece.content) {
    primitivesInto(each, { leftPt, baselinePt: at.baselinePt }, into);
    leftPt += widthOfSet(each);
  }
  side(box.closing);
}

// An equation read and marked, before anything has measured it. The reader in
// `docx/equations.ts` answers the shape and `docx/runs.ts` hangs a mark on each run;
// this is what those two come to, and `setMath` is what turns it into geometry.
export type MarkedMath =
  | { readonly kind: "run"; readonly text: string; readonly mark: ParagraphMark }
  | {
      readonly kind: "fraction";
      readonly mark: ParagraphMark;
      readonly numerator: readonly MarkedMath[];
      readonly denominator: readonly MarkedMath[];
    }
  | {
      readonly kind: "delimiter";
      readonly mark: ParagraphMark;
      readonly opening: number | null;
      readonly closing: number | null;
      readonly grows: boolean;
      readonly content: readonly MarkedMath[];
    };

// How wide a run of an equation is and how far its ink reaches. The caller measures
// rather than this file, because a run's width is the line's own question: the pairs
// it closes up by, the characters it borrows from another face and the face it is
// stood in by are all settled there and none of them is geometry.
export type MathMeasurer = (text: string, mark: ParagraphMark, sizePt: number) => MathBox | null;

// Several pieces standing side by side, as one box: a fraction's half and a
// delimiter's content are each a row of them.
function rowOf(pieces: readonly SetMath[]): MathBox {
  let widthPt = 0;
  let ascentPt = 0;
  let descentPt = 0;
  for (const piece of pieces) {
    widthPt += piece.box.widthPt;
    ascentPt = Math.max(ascentPt, piece.box.ascentPt);
    descentPt = Math.max(descentPt, piece.box.descentPt);
  }
  return { widthPt, ascentPt, descentPt };
}

export type SetMathRequest = {
  // The size the equation itself is set at, which is not always the size its halves
  // are drawn at: see `scriptSizePt`.
  readonly sizePt: number;
  readonly halfSizePt: number;
  readonly setting: MathSetting;
  readonly face: MathFace;
  readonly measure: MathMeasurer;
};

/**
 * An equation measured and placed, from the leaves up.
 *
 * Answers null where a run could not be measured at all, which is a face nothing on
 * the machine can answer for: the caller refuses the paragraph rather than laying it
 * out in nothing, as it does for any other run.
 */
export function setMath(
  pieces: readonly MarkedMath[],
  request: SetMathRequest,
): readonly SetMath[] | null {
  const set: SetMath[] = [];

  for (const piece of pieces) {
    if (piece.kind === "run") {
      const box = request.measure(piece.text, piece.mark, request.halfSizePt);
      if (box === null) return null;
      set.push({
        kind: "run",
        text: piece.text,
        mark: piece.mark,
        sizePt: request.halfSizePt,
        box,
      });
      continue;
    }

    if (piece.kind === "fraction") {
      // **A fraction standing inside another is set in the text constants**, whatever
      // the one round it was set in: measured over the nested pair, where the inner
      // fraction's two baselines stood 12.24pt apart at 11pt against the outer pair's
      // 15.84, which is the text shifts of 1200 and 1030 against the display ones of
      // 1550 and 1370. A delimiter passes the setting through untouched: the fraction
      // inside one, alone in its paragraph, was drawn in display like any other.
      const inside = { ...request, setting: "text" as const };
      const numerator = setMath(piece.numerator, inside);
      const denominator = setMath(piece.denominator, inside);
      if (numerator === null || denominator === null) return null;
      set.push({
        kind: "fraction",
        mark: piece.mark,
        box: fractionBox({
          numerator: rowOf(numerator),
          denominator: rowOf(denominator),
          sizePt: request.sizePt,
          setting: request.setting,
          face: request.face,
        }),
        numerator,
        denominator,
      });
      continue;
    }

    const content = setMath(piece.content, request);
    if (content === null) return null;
    set.push({
      kind: "delimiter",
      mark: piece.mark,
      sizePt: request.sizePt,
      box: delimiterBox({
        opening: piece.opening,
        closing: piece.closing,
        content: rowOf(content),
        sizePt: request.sizePt,
        grows: piece.grows,
        face: request.face,
      }),
      content,
    });
  }

  return set;
}

// The whole of what a set equation takes on the line it stands on.
export const mathRowOf = (pieces: readonly SetMath[]): MathBox => rowOf(pieces);
