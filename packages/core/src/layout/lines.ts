import type { RunPiece, TextRun } from "../docx/runs.js";
import type { ParagraphMark } from "../docx/styles.js";
import {
  advanceWidthPt,
  ascentPt,
  lineHeightPt,
  type AdvancesUnavailable,
  type BorrowedGlyph,
  type FaceElsewhere,
  type FaceRequest,
  type FontMetrics,
  type GlyphAdvances,
  type MetricsLookup,
  runKerns,
  type PairKerning,
  type RunKerning,
} from "./font-metrics.js";
import { NO_INK, NO_MATH } from "./font-metrics.js";
import {
  mathFace,
  mathLeadingPt,
  mathRowOf,
  scriptSizePt,
  setMath,
  textBox,
  type MathFace,
  type SetMath,
} from "./math.js";
import { nextTabStop, type TabStopPt } from "./tab-stops.js";
import { roomForTurn } from "./turns.js";
import { emuToPoints } from "./units.js";

export type MetricsResolver = ((request: FaceRequest) => MetricsLookup) & {
  // Whether a run whose style cascade names no face at all may be put to the
  // resolver anyway, as a request for the empty name. Only a resolver that has
  // an answer for every name says so; without it such a run refuses the
  // document, which is the exact behaviour the suites hold to.
  readonly answersForUnresolved?: true;
};

// Where the run sits along its line, measured from the line's own start. A tab
// opens a gap the runs after it never account for, so each one carries the place
// it reached rather than leaving it to be added up.
export type LineSegment =
  | {
      readonly kind: "text";
      readonly mark: ParagraphMark;
      readonly text: string;
      readonly widthPt: number;
      readonly offsetPt: number;
    }
  | {
      readonly kind: "drawing";
      readonly widthPt: number;
      readonly heightPt: number;
      readonly offsetPt: number;
    }
  // A set equation, already placed about its own baseline, which the line's baseline
  // is. A renderer asks `mathPrimitivesOf` what to draw and nothing here has to know
  // a fraction from a delimiter.
  | {
      readonly kind: "equation";
      readonly pieces: readonly SetMath[];
      readonly widthPt: number;
      readonly ascentPt: number;
      readonly descentPt: number;
      readonly offsetPt: number;
    };

export type TextLine = {
  readonly segments: readonly LineSegment[];
  readonly widthPt: number;
  readonly heightPt: number;
  // How far below the line's own top its baseline falls, which is the seat below
  // plus however far the tallest thing on the line reaches above that baseline.
  readonly ascentPt: number;
  // The room the floor below opened above everything on the line, which a drawing
  // shorter than the faces beside it stands under. Word answers for a paragraph
  // from here rather than from the line's own top.
  readonly seatPt: number;
  // The tallest line any face on this one would make on its own, which is not the
  // line's height once a drawing stands on it: a drawing reaches above the text
  // without being measured from a face. This is what a line multiple is taken of,
  // and what a line holding nothing but a drawing is never shorter than.
  readonly fontHeightPt: number;
  // The line the run holding the break at the end of this one would have made, for
  // a line with nothing measured on it to be held open by. **Word measures such a
  // line from that run rather than from the paragraph's mark**, which only the
  // paragraph's last line answers to: measured on 2026-08-11 by the authored
  // `empty-line-size` document. Null on a line no break ends, which is where the
  // mark stands.
  //
  // **The line a break ends and the line it opens are held open by different things,
  // and it is worth keeping them apart**, since a reader meeting the rule above will
  // take it for both. Measured on 2026-08-14 by the authored `equation-break-probe`
  // document, whose break stands in an 11pt maths run inside a paragraph whose own
  // mark is 12pt: a fraction with the break **before** it came out 40.08, which is the
  // fraction's 27.36 and the break's own run, and the same fraction with the break
  // **after** it came out 42.00, which is 27.36 and the paragraph's 12pt mark. Setting
  // that mark to 36pt moved the second to 71.28 and left the break's run where it was,
  // so the opened line is the mark's whatever the break was written in. That is the
  // rule above and `emptyLine(null)` already between them, and the probe is what says
  // a maths run is no exception.
  readonly heldOpenPt: number | null;
};

export type MeasureFailure =
  | { readonly kind: "unresolved-font" }
  | { readonly kind: "unknown-font-metrics"; readonly fontName: string }
  | {
      readonly kind: "unmeasurable-text";
      readonly fontName: string;
      readonly reason: AdvancesUnavailable;
    }
  | {
      readonly kind: "unmapped-character";
      readonly fontName: string;
      readonly codePoint: number;
    }
  | {
      // The face an equation asked to be set in carries no MATH table, or none of
      // the outlines its halves are measured off. This is not a face that cannot
      // be measured: ordinary text in it lays out perfectly well, and only an
      // equation wants the two tables.
      readonly kind: "unsettable-equation";
      readonly fontName: string;
    };

export type LineBreaking =
  | { readonly kind: "lines"; readonly lines: readonly TextLine[] }
  | { readonly kind: "unmeasurable"; readonly failure: MeasureFailure };

// A paragraph part way through being broken into lines. The width each line has is
// asked for as that line is taken, since an object beside it can leave it less room
// than the line above had, and asking takes nothing away: the same flow answers
// again at another width, which is what a line that has to be broken a second time
// needs.
export type LineFlow = {
  readonly next: (roomPt: number) => FlowedLine | null;
  // The narrowest the next line can be made, which is what it has to be given
  // before it can be drawn at all: a run of space narrower than this holds no
  // line, however the rest of the paragraph is broken.
  readonly leastPt: number;
  // Whether a page break stands in front of whatever this flow has left, so the
  // line it hands back begins a page. Asked of the flow rather than of the line,
  // since a flow with nothing left still has to answer: a break the paragraph ends
  // on is what puts the paragraph after it on a page of its own.
  readonly startsPage: boolean;
};

export type FlowedLine = {
  readonly line: TextLine;
  readonly rest: LineFlow;
};

export type LineFlowStart =
  | { readonly kind: "flow"; readonly flow: LineFlow }
  | { readonly kind: "unmeasurable"; readonly failure: MeasureFailure };

// Tab stops are measured from the left edge of the text area, so a line has to say
// how far its own start sits from that edge for a tab to land on the right one. A
// hanging first line starts outside that edge, and a tab on it reaches the stop at
// the indent that the lines below it start from.
export type LineTabs = {
  readonly stopsPt: readonly TabStopPt[];
  readonly originPt: number;
  readonly firstLineOriginPt?: number;
  // How far apart the stops the document falls back on stand, which is where a tab
  // past the last one the paragraph declares lands.
  readonly defaultStopPt?: number;
};

export type FlowInput = {
  readonly runs: readonly TextRun[];
  readonly metricsFor: MetricsResolver;
  readonly tabs?: LineTabs;
  // Whether the paragraph fills its lines out to both edges, which is what lets a
  // line take a word it has not the room for.
  readonly justified?: boolean;
};

export type BreakLinesInput = FlowInput & {
  readonly widthPt: number;
  // What the first line alone has room for, which a hanging indent makes wider
  // than the lines under it and a first-line indent makes narrower.
  readonly firstLineWidthPt?: number;
};

const NO_TABS: LineTabs = { stopsPt: [], originPt: 0 };

// Widths are compared, not accumulated into a coordinate, so this only has to
// absorb the last bits of a sum of exact ratios.
const EPSILON = 1e-9;

export const faceRequestFor = (mark: ParagraphMark): FaceRequest => ({
  name: mark.font.kind === "named" ? mark.font.name : "",
  bold: mark.bold,
  italic: mark.italic,
});

// What a run of one mark states about kerning, which is what the two predicates in
// `font-metrics.ts` are asked. A mark carries the threshold and the size; the face
// carries the pairs.
export const kerningOf = (mark: ParagraphMark): RunKerning => ({
  ...faceRequestFor(mark),
  kernFromHalfPoints: mark.kernFromHalfPoints,
  sizePt: mark.fontSizePt,
});

type Face = {
  readonly metrics: FontMetrics;
  readonly advanceFor: GlyphAdvances;
  readonly elsewhere: FaceElsewhere | null;
  // How far a pair of the face's characters moves together, in font units, or
  // nothing where the face states no pairs or nobody asked its file for them.
  readonly kerningBetween: PairKerning | null;
};

type Fragment = {
  readonly mark: ParagraphMark;
  readonly text: string;
  readonly widthPt: number;
  // How far into the fragment its first decimal point stands, or null where it
  // holds none. A decimal stop lines its text up on that point, and the width is
  // free while the characters are being added up anyway.
  readonly beforePointPt: number | null;
  // How far the fragment reaches above and below the line's own baseline, which a
  // raised run reaches off: the raise moves both ends and neither passes the
  // baseline going the other way.
  readonly heightPt: number;
  readonly ascentPt: number;
  // The line the fragment's faces make with nothing raised, which is what a
  // multiple line rule is taken of and what floors a line holding a short drawing.
  readonly fontHeightPt: number;
  // How far the fragment reaches after each of its characters, which is what a word
  // too long for the line it stands alone on is cut by. The total says nothing about
  // where a character inside it ends, since the characters of a face are not all one
  // width, and the reaches are added up here anyway.
  readonly reachPt: readonly number[];
};

type Unit =
  | { readonly kind: "word" | "space"; readonly fragments: readonly Fragment[] }
  | { readonly kind: "tab" }
  | {
      readonly kind: "break";
      readonly endsPage: boolean;
      // The line the break's own run would have made had it held text, which is
      // what holds open a line the break leaves nothing measured on.
      readonly heightPt: number;
    }
  | {
      readonly kind: "equation";
      readonly widthPt: number;
      readonly ascentPt: number;
      readonly descentPt: number;
      readonly fontHeightPt: number;
      readonly pieces: readonly SetMath[];
    }
  | {
      readonly kind: "drawing";
      readonly widthPt: number;
      readonly heightPt: number;
      // The line the drawing's own run would have made had it held text, which is
      // all the face on that run has to say about the line it stands on.
      readonly fontHeightPt: number;
    };

type Measured<T> = { readonly kind: "ok"; readonly value: T } | { readonly kind: "failed" };

// Word advances one code point at a time, so splitting there is the right
// granularity for measuring even though it is not grapheme-aware.
const charactersOf = (text: string): readonly string[] => Array.from(text);

const widthOf = (fragments: readonly Fragment[]): number =>
  fragments.reduce((sum, fragment) => sum + fragment.widthPt, 0);

class Measurer {
  private readonly faces = new Map<ParagraphMark, Face>();
  failure: MeasureFailure | null = null;

  constructor(private readonly metricsFor: MetricsResolver) {}

  // The face an equation is set out of, which wants more of a file than a line of
  // text does: the outlines a fraction's height is measured off and the MATH table it
  // takes its every constant from. A face carrying neither cannot set one at all.
  //
  // **What could not be had is recorded here**, because nothing above records it.
  // A null used to leave the measurer with no failure at all, and `measureText`
  // and `breakParagraph` each fill that hole with `unresolved-font`, so the only
  // two documents the corpus refuses were both refused for a cause that had not
  // happened: their runs name a face, and the face simply does not set equations.
  mathFaceFor(mark: ParagraphMark): MathFace | null {
    const fontName = mark.font.kind === "named" ? mark.font.name : "";
    const lookup = this.metricsFor(faceRequestFor(mark));
    if (lookup.kind === "missing") {
      this.failure ??= { kind: "unknown-font-metrics", fontName };
      return null;
    }

    const face = mathFace({
      metrics: lookup.metrics,
      advances: lookup.advances,
      ink: lookup.ink ?? NO_INK,
      math: lookup.math ?? NO_MATH,
    });
    if (face === null) this.failure ??= { kind: "unsettable-equation", fontName };
    return face;
  }

  private faceFor(mark: ParagraphMark): Face | null {
    const cached = this.faces.get(mark);
    if (cached !== undefined) return cached;

    if (mark.font.kind === "unresolved" && this.metricsFor.answersForUnresolved !== true) {
      this.failure ??= { kind: "unresolved-font" };
      return null;
    }

    const fontName = mark.font.kind === "named" ? mark.font.name : "";
    const lookup = this.metricsFor(faceRequestFor(mark));
    if (lookup.kind === "missing") {
      this.failure ??= { kind: "unknown-font-metrics", fontName };
      return null;
    }
    if (lookup.advances.kind !== "advances") {
      this.failure ??= { kind: "unmeasurable-text", fontName, reason: lookup.advances.reason };
      return null;
    }

    const face = {
      metrics: lookup.metrics,
      advanceFor: lookup.advances.advanceFor,
      elsewhere: lookup.elsewhere ?? null,
      kerningBetween: lookup.kerning?.kind === "kerning" ? lookup.kerning.kerningBetween : null,
    };
    this.faces.set(mark, face);
    return face;
  }

  // The line a face makes, asked without asking for any of its glyphs: a drawing
  // needs its run's face for nothing else, and a face with metrics but no
  // advances still answers for one.
  lineHeight(mark: ParagraphMark): number | null {
    if (mark.font.kind === "unresolved" && this.metricsFor.answersForUnresolved !== true) {
      this.failure ??= { kind: "unresolved-font" };
      return null;
    }
    const lookup = this.metricsFor(faceRequestFor(mark));
    if (lookup.kind === "missing") {
      this.failure ??= {
        kind: "unknown-font-metrics",
        fontName: mark.font.kind === "named" ? mark.font.name : "",
      };
      return null;
    }
    return lineHeightPt(lookup.metrics, mark.lineSizePt);
  }

  // A character drawn out of another face raises the fragment as a run in that
  // face would have: it stands as tall as the tallest face drawn in it and seats
  // its baseline under the deepest ascent among them.
  fragment(mark: ParagraphMark, text: string): Fragment | null {
    const face = this.faceFor(mark);
    if (face === null) return null;

    let widthPt = 0;
    const reachPt: number[] = [];
    let beforePointPt: number | null = null;
    let abovePt = ascentPt(face.metrics, mark.lineSizePt);
    let belowPt = lineHeightPt(face.metrics, mark.lineSizePt) - abovePt;

    // **The pairs are summed in font units and scaled once**, which is what the same
    // line at 12pt and at 10.5pt measured: it closed up by 11.94pt and by 10.45, and
    // 11.94 scaled by 10.5/12 is 10.4475. So they are gathered here rather than
    // converted a pair at a time.
    const kerningBetween = runKerns(kerningOf(mark)) ? face.kerningBetween : null;
    let kerningUnits = 0;
    let lastCodePoint: number | null = null;

    for (const character of text) {
      if (beforePointPt === null && character === DECIMAL_POINT) beforePointPt = widthPt;
      const codePoint = character.codePointAt(0) ?? 0;
      const drawn = drawnBy(face, codePoint);
      if (drawn === null) {
        this.failure ??= {
          kind: "unmapped-character",
          fontName: mark.font.kind === "named" ? mark.font.name : "",
          codePoint,
        };
        return null;
      }

      if (kerningBetween !== null && lastCodePoint !== null) {
        kerningUnits += kerningBetween(lastCodePoint, codePoint);
      }
      lastCodePoint = codePoint;

      // A space takes the spacing as a letter does. A tab never reaches here:
      // it leaves no segment behind, and ends at its stop regardless.
      // **The scale multiplies the glyph's own advance and the spacing is added
      // after it**, which is what a run scaled to 150 with a point of spacing
      // measured: 225.66pt and then 28, rather than 267.8 the other way round.
      widthPt +=
        advanceWidthPt(drawn.advance, drawn.metrics, mark.fontSizePt) * mark.characterScale +
        mark.characterSpacingPt;
      reachPt.push(widthPt + advanceWidthPt(kerningUnits, face.metrics, mark.fontSizePt));

      // The fragment already stands on its own face, so only a borrowed character
      // can raise it.
      if (drawn.metrics !== face.metrics) {
        const above = ascentPt(drawn.metrics, mark.lineSizePt);
        abovePt = Math.max(abovePt, above);
        belowPt = Math.max(belowPt, lineHeightPt(drawn.metrics, mark.lineSizePt) - above);
      }
    }

    // A run that asked to be raised off its baseline carries its whole line with
    // it: it reaches that much further above and that much less below, and neither
    // end crosses the baseline. Measured on 2026-08-07 by the authored `raised-text`
    // document, which is where the clamp comes from: a 12pt run raised six points,
    // alone on its line, left the line 17.52pt tall rather than 14.64, so what the
    // run no longer reaches below counts for nothing rather than pulling the next
    // line up.
    const raisedAbovePt = Math.max(0, abovePt + mark.lineRaisePt);
    const raisedBelowPt = Math.max(0, belowPt - mark.lineRaisePt);

    return {
      mark,
      text,
      widthPt: widthPt + advanceWidthPt(kerningUnits, face.metrics, mark.fontSizePt),
      beforePointPt,
      heightPt: raisedAbovePt + raisedBelowPt,
      ascentPt: raisedAbovePt,
      fontHeightPt: abovePt + belowPt,
      reachPt,
    };
  }
}

function drawnBy(face: Face, codePoint: number): BorrowedGlyph | null {
  const own = face.advanceFor(codePoint);
  if (own !== null) return { metrics: face.metrics, advance: own };
  return face.elsewhere === null ? null : face.elsewhere(codePoint);
}

// The character a decimal stop lines its text up on. Word takes it from the
// system's own number format; these documents are all written in one where it is
// the full stop.
const DECIMAL_POINT = ".";

/**
 * A no-break space is what its name says: text runs on through it, so it belongs
 * to the word around it rather than opening a place the line can break. Every
 * other run of whitespace is a gap between words.
 *
 * **Which spaces those are was measured on 2026-08-13**, by an authored document
 * putting each of twenty of them between the second and third of three words in a
 * column holding only two: a break there leaves `x1111 x2222` above `x3333`, and
 * no break leaves `x1111` above `x2222 x3333`. Word breaks at U+0020, U+2002,
 * U+2003, U+2005, U+3000, U+200B and a tab, and does not break at U+00A0, U+2007,
 * U+202F or U+FEFF, which are the three Unicode calls glue and the one it calls a
 * word joiner.
 *
 * **The nine left out are left out on purpose.** Word also declined to break at
 * U+2000, U+2001, U+2004, U+2006, U+2008, U+2009, U+200A, U+1680 and U+205F, which
 * Unicode all class as break-after, and U+2000 is canonically the same character as
 * U+2002 which Word does break at. The pdf draws two items on the line in several of
 * those cases, which is a face standing in for one the document's own has no glyph
 * for, so the answer is about the substitution rather than about the character. Of
 * the nine, only U+2008 turns up in the corpus at all: 15 documents and 111
 * characters, against U+202F's 20 documents and 1169. **Ask again with a face that
 * carries all of them before believing any of the nine.**
 */
const NO_BREAK = "\\u00a0\\u2007\\u202f\\ufeff";
const GAP = new RegExp(`([^\\S${NO_BREAK}]+)`);
const IS_GAP = new RegExp(`^[^\\S${NO_BREAK}]+$`);

/**
 * Word lets a line end on a hyphen inside a word, so each one closes the word it
 * belongs to and the rest of it becomes a word of its own.
 *
 * **Which characters those are was measured on 2026-08-25** by
 * `probes/hyphen-break-probe.ts`, twelve of them put between the halves of one
 * compound in a column wide enough for the compound alone and not for the word before
 * it: a break leaves the first half at the end of the first line, and no break carries
 * the whole compound to the second. Word breaks after **U+002D, U+2010, U+2012,
 * U+2013 and U+2014**, the hyphen-minus, the hyphen and the figure, en and em dashes.
 *
 * **It does not break after U+2011, U+00AD, U+2043, U+002F or U+2044**: the
 * non-breaking hyphen, a soft hyphen written as a character rather than as
 * `w:softHyphen`, the hyphen bullet, the solidus or the fraction slash. U+2015, the
 * horizontal bar, is the one that answers neither way: the compound went whole to the
 * second line and then broke again, which is a break before the bar and not after it,
 * and nothing here reads that.
 *
 * `8d31983d2189` is the reading in the wild. It ends a line on U+2010 where this
 * carried the compound on, and the two pages below that line each came out one line
 * out of place.
 */
const HYPHENS = "\\u002d\\u2010\\u2012\\u2013\\u2014";
const AFTER_HYPHEN = new RegExp(`(?<=[${HYPHENS}])`, "u");
const ENDS_ON_HYPHEN = new RegExp(`[${HYPHENS}]$`, "u");

const endsOnHyphen = (unit: Unit): boolean =>
  unit.kind === "word" && ENDS_ON_HYPHEN.test(unit.fragments.at(-1)?.text ?? "");

// **What decides the size a fraction's halves are set at is drawn text on the line,
// not the shape of the paragraph.** Measured on 2026-08-13: a paragraph holding a
// fraction and an empty run drew the halves full size and centred, one holding the
// fraction and a single space drew them at the script size and in the flow, and a
// delimiter standing alone was full size like anything else. So a run holding nothing
// does not count and a space does.
const drawsBesideAnEquation = (runs: readonly TextRun[]): boolean =>
  runs.some((run) =>
    run.pieces.some(
      (piece) =>
        (piece.kind === "text" && piece.text !== "") ||
        piece.kind === "tab" ||
        piece.kind === "drawing",
    ),
  );

function tokenize(runs: readonly TextRun[], measurer: Measurer): Measured<readonly Unit[]> {
  const units: Unit[] = [];
  const setting = drawsBesideAnEquation(runs) ? "text" : "display";

  const append = (kind: "word" | "space", fragment: Fragment): void => {
    const last = units.at(-1);
    if (last !== undefined && last.kind === kind && !endsOnHyphen(last)) {
      units[units.length - 1] = { kind, fragments: [...last.fragments, fragment] };
      return;
    }
    units.push({ kind, fragments: [fragment] });
  };

  for (const run of runs) {
    for (const piece of run.pieces) {
      if (!addPiece(piece, run.mark, units, append, measurer, setting)) return { kind: "failed" };
    }
  }

  return { kind: "ok", value: units };
}

function addPiece(
  piece: RunPiece,
  mark: ParagraphMark,
  units: Unit[],
  append: (kind: "word" | "space", fragment: Fragment) => void,
  measurer: Measurer,
  setting: "display" | "text",
): boolean {
  if (piece.kind === "equation") {
    const face = measurer.mathFaceFor(mark);
    const fontHeightPt = measurer.lineHeight(mark);
    if (face === null || fontHeightPt === null) return false;

    // The halves shrink where the line holds anything else that draws; the bar and
    // the axis keep to the size the run states either way.
    const halfSizePt = setting === "text" ? scriptSizePt(mark.fontSizePt, face) : mark.fontSizePt;
    const pieces = setMath(piece.content, {
      sizePt: mark.fontSizePt,
      halfSizePt,
      setting,
      face,
      measure: (text, runMark, sizePt) => {
        const own = measurer.mathFaceFor(runMark);
        return own === null ? null : textBox(text, sizePt, own);
      },
    });
    if (pieces === null) return false;

    // **A line holding an equation is the ink of it and the face's own
    // `mathLeading`, and the leading is room above.** Measured on 2026-08-13 over
    // the three fractions of the authored probe, whose paragraphs stood 1.59, 1.42
    // and 1.55pt above the ink of their boxes against a `mathLeading` of 1.62 at the
    // size Word drew them. It is asked for here rather than in the builder because
    // that room is what `fontHeightPt` already means: a line is never shorter than
    // it, and what it opens over the line's own content is a seat.
    const box = mathRowOf(pieces);
    units.push({
      kind: "equation",
      widthPt: box.widthPt,
      ascentPt: box.ascentPt,
      descentPt: box.descentPt,
      fontHeightPt: Math.max(
        fontHeightPt,
        box.ascentPt + box.descentPt + mathLeadingPt(mark.fontSizePt, face),
      ),
      pieces,
    });
    return true;
  }

  if (piece.kind === "tab") {
    units.push({ kind: "tab" });
    return true;
  }
  if (piece.kind === "break") {
    // The break's run is asked for nothing but its line height, which is what a
    // line the break leaves nothing on is held open by.
    const heightPt = measurer.lineHeight(mark);
    if (heightPt === null) return false;
    units.push({ kind: "break", endsPage: piece.endsPage, heightPt });
    return true;
  }
  if (piece.kind === "drawing") {
    // The run's face is asked for nothing but its line height, which is what a
    // drawing's line is held open by and what a multiple over it is taken of.
    const fontHeightPt = measurer.lineHeight(mark);
    if (fontHeightPt === null) return false;
    const room = roomForTurn(
      { widthPt: emuToPoints(piece.widthEmu), heightPt: emuToPoints(piece.heightEmu) },
      piece.turnDegrees,
    );
    units.push({ kind: "drawing", widthPt: room.widthPt, heightPt: room.heightPt, fontHeightPt });
    return true;
  }

  for (const token of piece.text.split(GAP).filter((each) => each !== "")) {
    const space = IS_GAP.test(token);
    for (const part of space ? [token] : token.split(AFTER_HYPHEN)) {
      const fragment = measurer.fragment(mark, part);
      if (fragment === null) return false;
      append(space ? "space" : "word", fragment);
    }
  }
  return true;
}

// What a stop that lines its text up reaches over: everything from the tab to the
// next tab, the next break, or the end of the paragraph, and how far into that its
// first decimal point stands. A tab's width can only be settled once this is
// known, which is why the line has to look ahead of itself to open one.
type TabbedSpan = {
  readonly widthPt: number;
  readonly toPointPt: number | null;
};

function spanAfterTab(units: readonly Unit[], from: number): TabbedSpan {
  let widthPt = 0;
  // A space the span ends on hangs past the stop rather than being lined up with
  // it, the way a space a line ends on hangs past the margin.
  let trailingSpacePt = 0;
  let toPointPt: number | null = null;

  for (let at = from; at < units.length; at += 1) {
    const unit = units[at];
    if (unit === undefined || unit.kind === "tab" || unit.kind === "break") break;

    if (unit.kind === "drawing" || unit.kind === "equation") {
      widthPt += unit.widthPt;
      trailingSpacePt = 0;
      continue;
    }

    for (const fragment of unit.fragments) {
      if (toPointPt === null && fragment.beforePointPt !== null) {
        toPointPt = widthPt + fragment.beforePointPt;
      }
      widthPt += fragment.widthPt;
    }
    trailingSpacePt = unit.kind === "space" ? trailingSpacePt + widthOf(unit.fragments) : 0;
  }

  return { widthPt: widthPt - trailingSpacePt, toPointPt };
}

// How far in front of the stop the text it lines up begins. A decimal stop puts
// the first point of that text on itself, and text with no point in it ends there,
// as a right stop's does.
function leadOf(stop: TabStopPt, span: TabbedSpan): number {
  if (stop.alignment === "center") return span.widthPt / 2;
  if (stop.alignment === "right") return span.widthPt;
  if (stop.alignment === "decimal") return span.toPointPt ?? span.widthPt;
  return 0;
}

// What became of a unit offered to a line: either the line took it, or the line is
// full and whatever is left of the unit starts the next one.
type Taken =
  { readonly kind: "taken" } | { readonly kind: "full"; readonly rest: readonly Fragment[] | null };

const TAKEN: Taken = { kind: "taken" };

// Word breaks greedily: it fills a line until the next word will not fit, then
// starts a new one, and the spaces it broke at hang past the edge rather than
// opening the next line. One line at a time, since the room the next one has is
// only known once this one has been placed.
class LineBuilder {
  private segments: LineSegment[] = [];
  private committedPt = 0;
  private pending: LineSegment[] = [];
  private pendingPt = 0;
  // The line the faces of the pending spaces would have made, kept because a space
  // never raises a line it shares and does make one of its own where a drawing is
  // moved off it. See `drawing`.
  private pendingFontHeightPt = 0;
  private ascentPt = 0;
  private descentPt = 0;
  private fontHeightPt = 0;
  private tabbed = false;

  constructor(
    private readonly room: number,
    private readonly tabs: LineTabs,
    private readonly index: number,
    // The line above ended because it filled up, so a gap at the head of this one
    // hangs there rather than opening this one.
    private readonly wrapped: boolean,
    private readonly justified: boolean,
  ) {}

  // **A justified line takes a word it has not the room for so long as its own
  // spaces can give up the difference: every space gives up at most a quarter of
  // itself, and the line at most a third of the advance of the word it is being
  // asked to take, counting the space in front of it, plus 0.2307 of that advance
  // over the line's spaces less a half.**
  //
  // Measured on 2026-08-10 off Word's own pdf, over 18 sweeps a twip apart: four to
  // thirty two spaces, three sizes, two faces, and words advancing from 5.6pt to
  // 43.8pt. The quarter is the bound where the word is a wide one, and Word took
  // 10.8125pt of a line offering 10.8516 and refused 10.8625. Everywhere else the
  // word's own advance is the bound, and it is that advance and not the em: a line
  // in Times New Roman, whose space is a quarter of its em where Calibri's is
  // 0.2261, stopped at the same fraction of it. The line's width does not enter at
  // all, a line of `mm` seven times ending in a short word stopping where the same
  // short word did on a line half as wide. The 0.2307 is bracketed to 0.2299-0.2315
  // by the eight and sixteen space sweeps, and the third is what the sweeps out to
  // thirty two spaces settle on.
  //
  // The squeeze itself is on the spaces and nowhere else, which Word's own report of
  // every character of such a line says: the letters advance by what the face makes
  // them and the spaces come out at 78% of theirs.
  private squeezePt(fragments: readonly Fragment[]): number {
    if (!this.justified) return 0;
    const held = [...this.segments, ...this.pending];
    const gapsPt = held.reduce(
      (widthPt, segment) => widthPt + (spaceCountOf(segment) === 0 ? 0 : segment.widthPt),
      0,
    );
    const spaces = held.reduce((count, segment) => count + spaceCountOf(segment), 0);
    if (spaces === 0) return 0;

    const advancePt =
      widthOf(fragments) + this.pending.reduce((widthPt, segment) => widthPt + segment.widthPt, 0);
    return Math.min(gapsPt / 4, advancePt * (1 / 3 + 0.2307 / (spaces - 0.5)));
  }

  // A line reaches as far above the baseline as the highest thing on it and as far
  // below as the deepest, which are not always the same thing: a drawing reaches
  // above the baseline and never below it.
  private raise(ascentPt: number, descentPt: number, fontHeightPt: number): void {
    this.ascentPt = Math.max(this.ascentPt, ascentPt);
    this.descentPt = Math.max(this.descentPt, descentPt);
    this.fontHeightPt = Math.max(this.fontHeightPt, fontHeightPt);
  }

  private commit(segments: readonly LineSegment[], widthPt: number): void {
    let offsetPt = this.committedPt + this.pendingPt;
    this.segments.push(...this.pending);
    for (const segment of segments) {
      this.segments.push(startingAt(segment, offsetPt));
      offsetPt += segment.widthPt;
    }
    this.committedPt += this.pendingPt + widthPt;
    this.pending = [];
    this.pendingPt = 0;
    this.pendingFontHeightPt = 0;
    this.tabbed = false;
  }

  private get filled(): number {
    return this.committedPt + this.pendingPt;
  }

  private get empty(): boolean {
    return this.segments.length === 0;
  }

  // A trailing space hangs past the edge, but a trailing tab holds the line open
  // as far as the stop it reached: Word wraps that line around the width the tab
  // gave it, even when nothing is drawn there. A line with neither is no line.
  finish(): TextLine | null {
    if (this.empty && !this.tabbed) return null;
    // A line is never shorter than the line its own faces would have made, and
    // what that floor opens is room above everything on it: a drawing shorter than
    // the faces beside it stands at the foot of the line rather than part way up.
    const contentPt = this.ascentPt + this.descentPt;
    const heightPt = Math.max(contentPt, this.fontHeightPt);
    return {
      segments: this.segments,
      widthPt: this.committedPt + (this.tabbed ? this.pendingPt : 0),
      heightPt,
      ascentPt: this.ascentPt,
      seatPt: heightPt - contentPt,
      fontHeightPt: this.fontHeightPt,
      heldOpenPt: null,
    };
  }

  // **A space is measured across and never up.** Word takes its width from the
  // face it is written in and its height from nothing at all, wherever on the line
  // it falls: measured on 2026-08-07 by the authored `trailing-space` document,
  // where a 24pt space came out the height of the 12pt mark behind it in front of
  // the text, at the end of it and between two of its words alike.
  space(fragments: readonly Fragment[]): Taken {
    if (this.empty && this.wrapped) return TAKEN;
    for (const fragment of fragments) {
      this.pending.push(startingAt(segmentOf(fragment), this.committedPt + this.pendingPt));
      this.pendingPt += fragment.widthPt;
      this.pendingFontHeightPt = Math.max(this.pendingFontHeightPt, fragment.fontHeightPt);
    }
    return TAKEN;
  }

  tab(span: TabbedSpan): Taken {
    if (this.empty && this.wrapped) return TAKEN;
    const { stopsPt } = this.tabs;
    const originPt =
      this.index === 0 ? (this.tabs.firstLineOriginPt ?? this.tabs.originPt) : this.tabs.originPt;
    const stop = nextTabStop(originPt + this.filled, stopsPt, this.tabs.defaultStopPt);
    // A stop that lines its text up never pulls it back over what the line already
    // holds: where there is not room for the text in front of the stop, the tab
    // opens none at all and the text carries on from where it was.
    const startPt = Math.max(this.filled, stop.positionPt - originPt - leadOf(stop, span));
    this.pendingPt = startPt - this.committedPt;
    this.tabbed = true;
    return TAKEN;
  }

  // A drawing stands on the baseline and reaches nothing below it, however deep
  // the text beside it goes.
  // An equation straddles the baseline where a drawing sits on it, so it raises the
  // line by its own ascent and its own descent rather than by one height.
  equation(unit: Extract<Unit, { kind: "equation" }>): Taken {
    if (!this.empty && this.filled + unit.widthPt > this.room + EPSILON) {
      return { kind: "full", rest: null };
    }
    this.raise(unit.ascentPt, unit.descentPt, unit.fontHeightPt);
    this.commit(
      [
        {
          kind: "equation",
          pieces: unit.pieces,
          widthPt: unit.widthPt,
          ascentPt: unit.ascentPt,
          descentPt: unit.descentPt,
          offsetPt: 0,
        },
      ],
      unit.widthPt,
    );
    return TAKEN;
  }

  /**
   * A drawing on the line, which moves to the next one where it does not fit.
   *
   * **Spaces in front of it are enough to move it, and they stay behind on a line
   * of their own.** A space is held pending rather than committed, so a line
   * holding nothing else reads as empty and used to take the drawing however wide
   * it was. Measured on 2026-08-24 in a 540pt frame: a 560pt picture opening a line
   * after one space was drawn a whole line below the same picture opening a line
   * with nothing in front of it, and a 400pt one went the same way. Four corpus
   * documents of one template hold that shape, each drawing its picture a line too
   * high and the width of one space too far right.
   *
   * The spaces left behind draw nothing, and the line they hold open is the line
   * their own face would have made: they raise nothing, as a space never does, and
   * `finish` floors a line at its faces.
   */
  drawing(widthPt: number, heightPt: number, fontHeightPt: number): Taken {
    const overflows = this.filled + widthPt > this.room + EPSILON;
    if (overflows && !this.empty) {
      return { kind: "full", rest: null };
    }
    if (overflows && this.pending.length > 0) {
      this.raise(0, 0, this.pendingFontHeightPt);
      this.commit([], 0);
      return { kind: "full", rest: null };
    }
    this.raise(heightPt, 0, fontHeightPt);
    this.commit([{ kind: "drawing", widthPt, heightPt, offsetPt: 0 }], widthPt);
    return TAKEN;
  }

  word(fragments: readonly Fragment[], cutting: boolean): Taken {
    const widthPt = widthOf(fragments);
    if (this.filled + widthPt <= this.room + this.squeezePt(fragments) + EPSILON) {
      this.take(fragments);
      return TAKEN;
    }
    if (!this.empty) return { kind: "full", rest: fragments };
    // A word with no line of its own to fit on is cut where it overflows, unless
    // what is being asked is how narrow the line can be made without cutting.
    if (!cutting) {
      this.take(fragments);
      return TAKEN;
    }

    const [head, tail] = splitFragments(fragments, this.room - this.filled);
    this.take(head);
    return { kind: "full", rest: tail };
  }

  private take(fragments: readonly Fragment[]): void {
    for (const fragment of fragments) {
      this.raise(fragment.ascentPt, fragment.heightPt - fragment.ascentPt, fragment.fontHeightPt);
    }
    this.commit(fragments.map(segmentOf), widthOf(fragments));
  }
}

// How far a paragraph has been broken: which unit is next, what is left of a word
// the line above cut in two, whether that line ended by filling up, whether it
// ended at a page break, and whether a line break opened a line that nothing has
// been put on yet.
type Cursor = {
  readonly at: number;
  readonly rest: readonly Fragment[] | null;
  readonly wrapped: boolean;
  readonly index: number;
  readonly startsPage: boolean;
  readonly opened: boolean;
};

const START: Cursor = {
  at: 0,
  rest: null,
  wrapped: false,
  index: 0,
  startsPage: false,
  opened: false,
};

// The line a break leaves behind it where nothing stood on it. **Every break opens
// a line under it, and that line stands whether anything is written on it or not**,
// so this is handed back rather than stepped over. A line a break ends is held open
// by the run that break stands in; one the paragraph merely ran out on is held open
// by the paragraph's mark, which is what `heldOpenPt` of null asks for.
const emptyLine = (heldOpenPt: number | null): TextLine => ({
  segments: [],
  widthPt: 0,
  heightPt: 0,
  ascentPt: 0,
  seatPt: 0,
  fontHeightPt: 0,
  heldOpenPt,
});

// One line's worth of units, or nothing left to draw one from. A line held open by
// nothing but a line break carries no text and takes no room, so the flow steps
// over it and keeps looking rather than handing back a line that is not there.
function fillFrom(
  units: readonly Unit[],
  tabs: LineTabs,
  justified: boolean,
  from: Cursor,
  roomPt: number,
  cutting: boolean,
): Drawn | null {
  let cursor = from;

  for (;;) {
    const built = fillLine(units, tabs, justified, cursor, roomPt, cutting);
    if (built === null) return null;
    if (built.line !== null) return { line: built.line, cursor: built.cursor };
    cursor = built.cursor;
  }
}

// That line, and the flow for whatever the paragraph has left after it.
function buildLine(
  units: readonly Unit[],
  tabs: LineTabs,
  justified: boolean,
  from: Cursor,
  roomPt: number,
): FlowedLine | null {
  const built = fillFrom(units, tabs, justified, from, roomPt, true);
  if (built === null) return null;
  return {
    line: built.line,
    rest: flowFrom(units, tabs, justified, { ...built.cursor, index: from.index + 1 }),
  };
}

const flowFrom = (
  units: readonly Unit[],
  tabs: LineTabs,
  justified: boolean,
  cursor: Cursor,
): LineFlow => ({
  // Room below nothing is the same as none. A paragraph can ask for indents wider
  // than the frame it stands in, which a real document does inside a narrow cell,
  // and the width left over is then negative. Given that, the builder took nothing
  // and the cursor never moved, so `buildLine` asked again for ever: a document out
  // of the wild hung the layout outright. Nothing here decides what such a
  // paragraph should look like, only that it is measured at all.
  next: (roomPt) => buildLine(units, tabs, justified, cursor, Math.max(0, roomPt)),
  leastPt: leastRoomPt(units, tabs, justified, cursor),
  startsPage: cursor.startsPage,
});

// The line the paragraph makes when it is given no room at all and is not allowed
// to cut a word in two: it runs to the first place a break is legal, and how far
// that reaches is what any run of space has to hold. A word is unbreakable, a tab
// holds the line open to the stop it reached, and a gap hangs past the edge.
//
// **This asks for the line alone and never for the flow after it.** Building that
// flow asks its own least room, which built the flow after that, so making one flow
// walked every break the paragraph had left: a paragraph of a few thousand words ran
// the stack out, and every line before that one cost the whole remainder again.
function leastRoomPt(
  units: readonly Unit[],
  tabs: LineTabs,
  justified: boolean,
  cursor: Cursor,
): number {
  return fillFrom(units, tabs, justified, cursor, 0, false)?.line.widthPt ?? 0;
}

type Filled = { readonly line: TextLine | null; readonly cursor: Cursor };

type Drawn = { readonly line: TextLine; readonly cursor: Cursor };

function fillLine(
  units: readonly Unit[],
  tabs: LineTabs,
  justified: boolean,
  cursor: Cursor,
  roomPt: number,
  cutting: boolean,
): Filled | null {
  const builder = new LineBuilder(roomPt, tabs, cursor.index, cursor.wrapped, justified);
  let at = cursor.at;

  if (cursor.rest !== null) {
    const took = builder.word(cursor.rest, cutting);
    if (took.kind === "full") {
      return filledAt(builder, {
        ...cursor,
        rest: took.rest,
        wrapped: true,
        startsPage: false,
        opened: false,
      });
    }
  }

  while (at < units.length) {
    const unit = units[at];
    if (unit === undefined) break;
    at += 1;

    // A break ends the line it is on wherever it stands, and the line under it
    // starts with whatever gap follows rather than losing it.
    //
    // **The line it opens stands whether anything is written on it or not.**
    // Measured on 2026-08-08 by the authored `breaks-in-a-paragraph` document, over
    // eight cases written out three times: a paragraph came out one line taller for
    // every break in it, wherever the breaks stood, and one holding a break and
    // nothing else came out two lines tall.
    if (unit.kind === "break") {
      const line = builder.finish();
      return {
        line: line === null ? emptyLine(unit.heightPt) : { ...line, heldOpenPt: unit.heightPt },
        cursor: {
          at,
          rest: null,
          wrapped: false,
          index: cursor.index,
          startsPage: unit.endsPage,
          // A page break carries its ask to the page under it rather than to a line
          // of its own, and a paragraph ending on one draws nothing more.
          opened: !unit.endsPage,
        },
      };
    }

    const took = tookOf(builder, unit, cutting, () => spanAfterTab(units, at));
    if (took.kind === "full") {
      // A unit the line could not take at all waits where it is; one it cut in two
      // hands the rest of itself to the line below.
      const rest = took.rest;
      return filledAt(builder, {
        at: rest === null ? at - 1 : at,
        rest,
        wrapped: true,
        index: cursor.index,
        startsPage: false,
        opened: false,
      });
    }
  }

  const line = builder.finish();
  const spent: Cursor = {
    at,
    rest: null,
    wrapped: false,
    index: cursor.index,
    startsPage: false,
    opened: false,
  };
  // Nothing left to write, and a break that opened a line for it: the line is
  // handed back empty, which is what makes a paragraph ending in a break one line
  // taller than its text.
  if (line === null) return cursor.opened ? { line: emptyLine(null), cursor: spent } : null;
  return { line, cursor: spent };
}

const filledAt = (builder: LineBuilder, cursor: Cursor): Filled => ({
  line: builder.finish(),
  cursor,
});

function tookOf(
  builder: LineBuilder,
  unit: Unit,
  cutting: boolean,
  spanAfter: () => TabbedSpan,
): Taken {
  switch (unit.kind) {
    case "word":
      return builder.word(unit.fragments, cutting);
    case "space":
      return builder.space(unit.fragments);
    case "tab":
      return builder.tab(spanAfter());
    case "equation":
      return builder.equation(unit);
    case "drawing":
      return builder.drawing(unit.widthPt, unit.heightPt, unit.fontHeightPt);
    case "break":
      return TAKEN;
  }
}

const segmentOf = (fragment: Fragment): LineSegment => ({
  kind: "text",
  mark: fragment.mark,
  text: fragment.text,
  widthPt: fragment.widthPt,
  offsetPt: 0,
});

const startingAt = (segment: LineSegment, offsetPt: number): LineSegment =>
  segment.kind === "text" ? { ...segment, offsetPt } : { ...segment, offsetPt };

// A word with no line of its own to fit on is cut at the character that overflows,
// and never before the first one, so a narrow column still makes progress.
//
// **The cut falls at the character wherever the word came from**, and a word written
// in more than one run is one word. This used to send a whole fragment to the next
// line as soon as it did not fit, which put the cut at the boundary between two runs:
// measured on 2026-08-08 by the authored `insignificant-space` document, Word cut
// `aaaaaaaaaaaa` and `bbbbbbbbbbbb` written as two runs in exactly the place it cut
// the same twenty four characters written as one, and this project cut at the twelfth
// wherever the runs were divided.
function splitFragments(
  fragments: readonly Fragment[],
  availablePt: number,
): readonly [readonly Fragment[], readonly Fragment[]] {
  const head: Fragment[] = [];
  const tail: Fragment[] = [];
  let taken = 0;
  let filled = 0;

  for (const fragment of fragments) {
    if (tail.length > 0) {
      tail.push(fragment);
      continue;
    }
    if (filled + fragment.widthPt <= availablePt + EPSILON) {
      head.push(fragment);
      filled += fragment.widthPt;
      taken += charactersOf(fragment.text).length;
      continue;
    }

    const [left, right] = splitFragment(fragment, availablePt - filled, taken === 0);
    if (left !== null) {
      head.push(left);
      filled += left.widthPt;
      taken += charactersOf(left.text).length;
    }
    if (right !== null) tail.push(right);
  }

  return [head, tail];
}

// **The character the cut falls at is the one whose own advance overflows.** The
// fragment's width used to be shared out evenly over its characters, which put the
// cut wherever the average said and handed both halves a width no face makes: eight
// characters of `iiiiMMMM` average to nothing any of them measures. So the reaches
// the fragment was measured by are what it is cut by, and the head comes out the
// width the face would make of the head's own text.
//
// The tail keeps whatever the head left of the fragment's width, which is a kern for
// the pair the cut now stands between. Nothing says what Word does with that pair.
function splitFragment(
  fragment: Fragment,
  availablePt: number,
  atLeastOne: boolean,
): readonly [Fragment | null, Fragment | null] {
  const characters = charactersOf(fragment.text);
  const reachOf = (count: number): number => fragment.reachPt[count - 1] ?? 0;

  let count = 0;
  while (count < characters.length && reachOf(count + 1) <= availablePt + EPSILON) {
    count += 1;
  }
  if (count === 0 && atLeastOne) count = 1;
  const filled = reachOf(count);

  const head = characters.slice(0, count).join("");
  const tail = characters.slice(count).join("");
  return [
    head === ""
      ? null
      : { ...fragment, text: head, widthPt: filled, reachPt: reachesTo(fragment, count) },
    tail === ""
      ? null
      : {
          ...fragment,
          text: tail,
          widthPt: fragment.widthPt - filled,
          reachPt: reachesFrom(fragment, count, filled),
        },
  ];
}

const reachesTo = (fragment: Fragment, count: number): readonly number[] =>
  fragment.reachPt.slice(0, count);

const reachesFrom = (fragment: Fragment, count: number, filledPt: number): readonly number[] =>
  fragment.reachPt.slice(count).map((each) => each - filledPt);

// Word justifies a line by handing every space character on it an equal share of
// the room the line did not fill, whatever size that space is set in. A tab takes
// none, since it holds the stop it reached; a no-break space takes none either,
// being part of the word around it rather than a gap between words; and a space
// the line ended on has already hung past the edge and is not on the line at all.
//
// The share is negative where the line took a word it had not the room for, and the
// spaces give it back: measured over every character of such a line, they come out
// at 78% of the width the face makes them and the letters at the width they always
// were.
export function justifyLine(line: TextLine, roomPt: number): TextLine {
  const slackPt = roomPt - line.widthPt;
  const spaces = line.segments.reduce((count, segment) => count + spaceCountOf(segment), 0);
  if (Math.abs(slackPt) <= EPSILON || spaces === 0) return line;

  const sharePt = slackPt / spaces;
  let shiftPt = 0;

  const segments = line.segments.map((segment) => {
    const moved = { ...segment, offsetPt: segment.offsetPt + shiftPt };
    const grownPt = spaceCountOf(segment) * sharePt;
    if (grownPt === 0) return moved;

    shiftPt += grownPt;
    return { ...moved, widthPt: segment.widthPt + grownPt };
  });

  return { ...line, segments, widthPt: roomPt };
}

const spaceCountOf = (segment: LineSegment): number =>
  segment.kind === "text" && IS_GAP.test(segment.text) ? charactersOf(segment.text).length : 0;

export type TextMeasurement =
  | {
      readonly kind: "measured";
      readonly widthPt: number;
      readonly heightPt: number;
      readonly ascentPt: number;
    }
  | { readonly kind: "unmeasurable"; readonly failure: MeasureFailure };

// A run of text that never breaks, which is what a list number is.
export function measureText(
  text: string,
  mark: ParagraphMark,
  metricsFor: MetricsResolver,
): TextMeasurement {
  const measurer = new Measurer(metricsFor);
  const fragment = measurer.fragment(mark, text);
  if (fragment === null) {
    return { kind: "unmeasurable", failure: measurer.failure ?? { kind: "unresolved-font" } };
  }
  return {
    kind: "measured",
    widthPt: fragment.widthPt,
    heightPt: fragment.heightPt,
    ascentPt: fragment.ascentPt,
  };
}

// Everything a paragraph's text is broken from, measured once: what fails here
// fails whatever width the lines are then given.
export function beginLines(input: FlowInput): LineFlowStart {
  const measurer = new Measurer(input.metricsFor);
  const tokens = tokenize(input.runs, measurer);
  if (tokens.kind === "failed") {
    return { kind: "unmeasurable", failure: measurer.failure ?? { kind: "unresolved-font" } };
  }

  const tabs = input.tabs ?? NO_TABS;
  return { kind: "flow", flow: flowFrom(tokens.value, tabs, input.justified === true, START) };
}

// Every line of a paragraph at one width, which is what a stack with nothing
// standing beside it asks for.
export function breakLines(input: BreakLinesInput): LineBreaking {
  const started = beginLines(input);
  if (started.kind === "unmeasurable") return started;

  const lines: TextLine[] = [];
  let flow = started.flow;
  for (;;) {
    const taken = flow.next(
      lines.length === 0 ? (input.firstLineWidthPt ?? input.widthPt) : input.widthPt,
    );
    if (taken === null) return { kind: "lines", lines };
    lines.push(taken.line);
    flow = taken.rest;
  }
}
