import { describe, expect, it } from "vitest";

import { buildSfnt, type FontFixture } from "../testing/build-font.js";
import { readFontFile } from "./font-file.js";
import type { MathConstants } from "./font-metrics.js";
import {
  delimiterBox,
  mathPrimitivesOf,
  fractionBox,
  mathFace,
  mathLeadingPt,
  mathRowOf,
  scriptSizePt,
  setMath,
  textBox,
  type MarkedMath,
  type MathBox,
  type MathFace,
  type MathSetting,
  type SetMath,
} from "./math.js";
import type { ParagraphMark } from "../docx/styles.js";

// Word's own pdf reports on a grid of 0.24pt, so a number here agrees with the
// measurement when it stands within one step of it. Nothing below is asserted more
// tightly than the oracle can answer.
const GRID_PT = 0.24;

const expectAsWordDrewIt = (actual: number, measuredPt: number): void => {
  expect(Math.abs(actual - measuredPt)).toBeLessThanOrEqual(GRID_PT + 1e-9);
};

const UNITS_PER_EM = 2048;

// Cambria Math's own constants, read off
// `/Applications/Microsoft Word.app/Contents/Resources/DFonts/Cambria.ttc` on 2026-08-13
// by `font-file.ts`. **Nothing below is fitted any longer**: every number Word drew
// falls out of these.
const CAMBRIA_MATH: Partial<MathConstants> = {
  axisHeight: 585,
  fractionRuleThickness: 133,
  fractionNumeratorShiftUp: 1200,
  fractionNumeratorDisplayStyleShiftUp: 1550,
  fractionDenominatorShiftDown: 1030,
  fractionDenominatorDisplayStyleShiftDown: 1370,
  fractionNumeratorGapMin: 133,
  fractionDenominatorGapMin: 133,
  fractionNumDisplayStyleGapMin: 260,
  fractionDenomDisplayStyleGapMin: 260,
  delimitedSubFormulaMinHeight: 3000,
  mathLeading: 300,
  scriptPercentScaleDown: 73,
  scriptScriptPercentScaleDown: 60,
};

// **Word snapped every size in the probe onto the 0.24pt grid**: a run stating 11pt was
// drawn at 11.04, one stating 10 at 10.08, one stating 20 at 19.92 and the shrunken
// halves at 7.92, while the 12pt body came out exact because 12 is already a multiple.
// So a case computes at the size Word drew rather than the size the file stated.
const ELEVEN = 11.04;
const TEN = 10.08;
const TWENTY = 19.92;
const SHRUNK = 7.92;

// The ink of the letters the halves were set in, read off Cambria Math itself. Word
// draws a maths run in the Mathematical Italic block, where `gralm` reaches 1430 up on
// its `l` and 447 down on its `g`, and `presk` the same 1430 on its `k` and 430 down on
// its `p`.
const ASCENDER = 1430;
const NUMERATOR_DESCENDER = -447;
const DESCENDER = -430;
// `gr`, the half that showed a fraction is measured off ink rather than off the face's
// own ascent: no ascender in it at all.
const X_HEIGHT = 973;
// `kesta`, which has an ascender and all but no descender.
const KESTA_DESCENDER = -17;

// Cambria Math's own ladder for `(`, every rung of it, in font units.
const LADDER = [
  { measurement: 1898, advance: 850, top: 1445, bottom: -452 },
  { measurement: 2475, advance: 920, top: 1822, bottom: -652 },
  { measurement: 3379, advance: 1004, top: 2275, bottom: -1103 },
  { measurement: 4047, advance: 1012, top: 2608, bottom: -1438 },
  { measurement: 5223, advance: 1175, top: 3197, bottom: -2025 },
  { measurement: 6053, advance: 1248, top: 3611, bottom: -2441 },
  { measurement: 7613, advance: 1396, top: 4392, bottom: -3220 },
  { measurement: 8881, advance: 1502, top: 5025, bottom: -3855 },
];

// A rung is named by the character whose glyph the face draws for it, so each one is a
// character of the fixture with a box and an advance of its own.
const RUNGS = ["A", "B", "C", "D", "E", "F", "G", "H"];

const OPENING = "(".codePointAt(0) ?? 0;
const CLOSING = ")".codePointAt(0) ?? 0;

// A face carrying what Cambria Math carries, built in memory and read back by the
// reader the layout uses. **Nothing here stands in for the reader**: the constants go
// through the MATH table's own bytes, the ink through `glyf` and the ladder through the
// variants table.
function faceHolding(fixture: Partial<FontFixture> = {}): MathFace {
  const rungAdvances = Object.fromEntries(
    LADDER.map((rung, index) => [RUNGS[index] ?? "", rung.advance]),
  );
  const rungBoxes = Object.fromEntries(
    LADDER.map((rung, index) => [
      RUNGS[index] ?? "",
      { left: 0, right: rung.advance, top: rung.top, bottom: rung.bottom },
    ]),
  );
  const ladder = LADDER.map((rung, index) => ({
    character: RUNGS[index] ?? "",
    measurement: rung.measurement,
  }));

  const built = buildSfnt({
    unitsPerEm: UNITS_PER_EM,
    ascender: 1946,
    descender: -455,
    lineGap: 0,
    advances: { l: 500, r: 700, g: 1000, " ": 600, "(": 850, ")": 850, ...rungAdvances },
    boxes: {
      l: { left: 0, bottom: 0, right: 500, top: ASCENDER },
      r: { left: 0, bottom: 0, right: 700, top: X_HEIGHT },
      g: { left: 0, bottom: NUMERATOR_DESCENDER, right: 1000, top: X_HEIGHT },
      "(": { left: 0, bottom: -452, right: 850, top: 1445 },
      ")": { left: 0, bottom: -452, right: 850, top: 1445 },
      ...rungBoxes,
    },
    math: { constants: CAMBRIA_MATH, tallerVariants: { "(": ladder, ")": ladder } },
    ...fixture,
  });

  const face = mathFace(readFontFile(built));
  if (face === null) throw new Error("the face carries no mathematics");
  return face;
}

const FACE = faceHolding();

const inPoints = (units: number, sizePt: number): number => (units * sizePt) / UNITS_PER_EM;

// A half of a fraction, at the size it is drawn: the ink of the letters Word set, and
// the width Word drew them at.
const halfOf = (sizePt: number, widthPt: number, top = ASCENDER, bottom = DESCENDER): MathBox => ({
  widthPt,
  ascentPt: inPoints(top, sizePt),
  descentPt: -inPoints(bottom, sizePt),
  // A half is text, and text reaches its own edges.
  insetPt: 0,
});

// `gralm` over `presk`, which is the fraction of every case but one, at the widths Word
// drew them.
const halvesOf = (sizePt: number, numeratorWidth: number, denominatorWidth: number) => ({
  numerator: halfOf(sizePt, numeratorWidth, ASCENDER, NUMERATOR_DESCENDER),
  denominator: halfOf(sizePt, denominatorWidth, ASCENDER, DESCENDER),
});

const fractionOf = (sizePt: number, setting: MathSetting) =>
  fractionBox({ ...halvesOf(sizePt, 30.817, 28.002), sizePt, setting, face: FACE });

describe("a face that can set an equation", () => {
  it("reads its constants, its outlines and its ladder through the file itself", () => {
    expect(FACE.unitsPerEm).toBe(UNITS_PER_EM);
    expect(FACE.constants.axisHeight).toBe(585);
    expect(FACE.constants.fractionRuleThickness).toBe(133);
    expect(FACE.constants.delimitedSubFormulaMinHeight).toBe(3000);
    expect(FACE.inkOf("l".codePointAt(0) ?? 0)?.top).toBe(ASCENDER);
    expect(FACE.tallerVariantsOf(OPENING).map((each) => each.measurement)).toStrictEqual(
      LADDER.map((rung) => rung.measurement),
    );
  });

  // Nearly every face states no MATH table at all, and a fraction set off another
  // face's constants is a plausible-looking page rather than Word's.
  it("answers nothing for a face carrying no mathematics", () => {
    const bare = buildSfnt({
      unitsPerEm: UNITS_PER_EM,
      ascender: 1946,
      descender: -455,
      lineGap: 0,
      advances: { l: 500 },
    });
    expect(mathFace(readFontFile(bare))).toBeNull();
  });

  it("answers nothing for a face whose outlines it cannot read", () => {
    const unread = buildSfnt({
      unitsPerEm: UNITS_PER_EM,
      ascender: 1946,
      descender: -455,
      lineGap: 0,
      advances: { l: 500 },
      math: { constants: CAMBRIA_MATH },
    });
    expect(mathFace(readFontFile(unread))).toBeNull();
  });
});

// Page 2 of `equation-probe.pdf`: the numerator's baseline at 60.000, the denominator's
// at 75.840, the bar's top at 64.800 and 0.7200 thick.
describe("a fraction standing alone in its paragraph", () => {
  const fraction = fractionOf(ELEVEN, "display");

  it("puts the two baselines where Word put them", () => {
    expectAsWordDrewIt(fraction.numerator.baselinePt - fraction.denominator.baselinePt, 15.84);
  });

  it("puts the bar where Word put it, and draws it as thick", () => {
    expectAsWordDrewIt(fraction.numerator.baselinePt - fraction.bar.topPt, 4.8);
    expectAsWordDrewIt(fraction.bar.topPt - fraction.denominator.baselinePt, 11.04);
    expectAsWordDrewIt(fraction.bar.thicknessPt, 0.72);
  });

  it("leaves the gaps Word left round the bar", () => {
    const numeratorInkBottom = fraction.numerator.baselinePt - fraction.numerator.descentPt;
    const denominatorInkTop = fraction.denominator.baselinePt + fraction.denominator.ascentPt;
    expectAsWordDrewIt(numeratorInkBottom - fraction.bar.topPt, 2.39);
    expectAsWordDrewIt(fraction.bar.topPt - fraction.bar.thicknessPt - denominatorInkTop, 2.66);
  });

  // Word's ink ran from the numerator's top at 52.291 to the denominator's foot at
  // 78.158, which is 25.867 measured off the baselines the pdf gives and the ink
  // Cambria states for those letters.
  it("covers the ink Word covered", () => {
    expectAsWordDrewIt(fraction.ascentPt + fraction.descentPt, 25.867);
  });

  // **A line holding a fraction is its ink and the face's own leading**: the paragraph
  // came out 27.36 against the box's 25.77, and `mathLeading` is 1.62 at 11.04.
  it("stands its line the face's own leading above its ink", () => {
    expectAsWordDrewIt(fraction.ascentPt + fraction.descentPt + mathLeadingPt(ELEVEN, FACE), 27.36);
  });

  // Word drew the bar 30.960 wide against the numerator's own 30.817, and nothing
  // measured explains the 0.143 between them.
  // The box stands 1.08pt out past the bar on each side at 11pt, which is what the
  // bar of a fraction standing above this one spans. It keeps to the em: the same
  // nested fraction at 20pt put the outer bar 1.92 out on each side.
  it("draws the bar as wide as the wider half, and centres the halves on one another", () => {
    expect(fraction.bar.widthPt).toBe(30.817);
    expectAsWordDrewIt(fraction.bar.widthPt, 30.96);
    expectAsWordDrewIt(fraction.widthPt - fraction.bar.widthPt, 2.16);
    const centre = (placed: { readonly leftPt: number; readonly widthPt: number }): number =>
      placed.leftPt + placed.widthPt / 2;
    expect(centre(fraction.numerator)).toBeCloseTo(centre(fraction.denominator), 10);
    expect(centre(fraction.numerator)).toBeCloseTo(centre(fraction.bar), 10);
  });
});

// Pages 7 and 8, where the same fraction was stated at 10 and 20pt and drawn at 10.08
// and 19.92.
describe("the same fraction at three sizes", () => {
  const measured = [
    { sizePt: TEN, widths: [28.053, 25.48], baselines: 14.4, barBelow: 4.56, thickness: 0.72 },
    { sizePt: ELEVEN, widths: [30.817, 28.002], baselines: 15.84, barBelow: 4.8, thickness: 0.72 },
    { sizePt: TWENTY, widths: [55.9, 50.828], baselines: 28.56, barBelow: 8.64, thickness: 1.2 },
  ];

  for (const each of measured) {
    it(`stands where Word stood it at ${String(each.sizePt)}pt`, () => {
      const fraction = fractionBox({
        ...halvesOf(each.sizePt, each.widths[0] ?? 0, each.widths[1] ?? 0),
        sizePt: each.sizePt,
        setting: "display",
        face: FACE,
      });
      expectAsWordDrewIt(
        fraction.numerator.baselinePt - fraction.denominator.baselinePt,
        each.baselines,
      );
      expectAsWordDrewIt(fraction.numerator.baselinePt - fraction.bar.topPt, each.barBelow);
      expectAsWordDrewIt(fraction.bar.thicknessPt, each.thickness);
    });
  }
});

// **The bar is the face's rule thickness snapped to the nearest 0.24pt.** The same
// fraction at seven sizes against Cambria's 133/2048: four rounded up and three down,
// so it is neither a floor nor a ceiling, and it is the first length in this project
// to say the grid belongs to Word's own arithmetic and not to the pdf.
describe("the thickness of the bar", () => {
  const measured = [
    { sizePt: 9, asked: 0.584, drawn: 0.48 },
    { sizePt: 10, asked: 0.649, drawn: 0.72 },
    { sizePt: 11, asked: 0.714, drawn: 0.72 },
    { sizePt: 13, asked: 0.844, drawn: 0.96 },
    { sizePt: 15, asked: 0.974, drawn: 0.96 },
    { sizePt: 17, asked: 1.104, drawn: 1.2 },
    { sizePt: 20, asked: 1.299, drawn: 1.2 },
  ];

  for (const each of measured) {
    it(`draws it ${each.drawn.toFixed(2)}pt at ${String(each.sizePt)}pt, where the face asks ${each.asked.toFixed(3)}`, () => {
      const fraction = fractionBox({
        ...halvesOf(each.sizePt, 30, 28),
        sizePt: each.sizePt,
        setting: "display",
        face: FACE,
      });

      expect(fraction.bar.thicknessPt).toBeCloseTo(each.drawn, 10);
    });
  }
});

// Page 6: the text's baseline at 64.080, the numerator's at 57.600 and drawn at 7.92,
// the denominator's at 69.600, the bar's top at 60.480 and still 0.7200 thick.
describe("a fraction beside ordinary text", () => {
  const fraction = fractionBox({
    ...halvesOf(SHRUNK, 24.736, 22.273),
    sizePt: ELEVEN,
    setting: "text",
    face: FACE,
  });

  it("shrinks the halves by the face's own percentage and no further", () => {
    expect(FACE.constants.scriptPercentScaleDown).toBe(73);
    expectAsWordDrewIt(scriptSizePt(ELEVEN, FACE), SHRUNK);
  });

  it("keeps the bar at the size the run stated", () => {
    expectAsWordDrewIt(fraction.bar.thicknessPt, 0.72);
  });

  // Against the text's own baseline, which is the fraction's own when it is set inline.
  it("hangs the halves and the bar where Word hung them", () => {
    expectAsWordDrewIt(fraction.numerator.baselinePt, 6.48);
    expectAsWordDrewIt(-fraction.denominator.baselinePt, 5.52);
    expectAsWordDrewIt(fraction.bar.topPt, 3.6);
    expectAsWordDrewIt(fraction.numerator.baselinePt - fraction.denominator.baselinePt, 12.0);
  });
});

// Page 13: the inner numerator's baseline at 60.000, the inner denominator's at 72.240,
// the inner bar's top at 63.120, the outer bar's at 75.840 and the outer denominator's
// baseline at 86.880.
describe("a fraction standing in another fraction's numerator", () => {
  const inner = fractionBox({
    ...halvesOf(ELEVEN, 30.817, 28.002),
    sizePt: ELEVEN,
    setting: "text",
    face: FACE,
  });
  const outer = fractionBox({
    numerator: inner,
    denominator: halfOf(ELEVEN, 27.094, ASCENDER, KESTA_DESCENDER),
    sizePt: ELEVEN,
    setting: "display",
    face: FACE,
  });

  // Both of the inner fraction's gaps are the face's own least, which is its rule
  // thickness: 133 units is 0.717pt at 11.04, and Word's ink came back 0.710 and 0.691.
  it("holds the inner fraction to the least gap the face states", () => {
    const inkBottom = inner.numerator.baselinePt - inner.numerator.descentPt;
    const inkTop = inner.denominator.baselinePt + inner.denominator.ascentPt;
    expectAsWordDrewIt(inkBottom - inner.bar.topPt, 0.71);
    expectAsWordDrewIt(inner.bar.topPt - inner.bar.thicknessPt - inkTop, 0.691);
    expectAsWordDrewIt(inner.numerator.baselinePt - inner.denominator.baselinePt, 12.24);
  });

  // The inner fraction is deep enough that the outer one cannot hold to its own shift,
  // so the least gap is what puts it. Word's own numbers give 12.856 between the two
  // baselines, read off the outer denominator and the inner numerator.
  it("pushes the outer numerator up until its least gap is left", () => {
    expectAsWordDrewIt(outer.numerator.baselinePt, 12.856);
    const inkBottom = outer.numerator.baselinePt - outer.numerator.descentPt;
    expectAsWordDrewIt(inkBottom - outer.bar.topPt, 1.4);
  });

  it("leaves the outer denominator on the face's own shift, which nothing pushed", () => {
    expectAsWordDrewIt(-outer.denominator.baselinePt, 7.44);
  });

  // Word's ink ran 52.291 to 86.972, which is 34.681.
  it("covers the ink Word covered", () => {
    expectAsWordDrewIt(outer.ascentPt + outer.descentPt, 34.681);
  });

  // The paragraph came out 36.24 against the box's 34.69.
  it("stands its line the face's own leading above its ink", () => {
    expectAsWordDrewIt(outer.ascentPt + outer.descentPt + mathLeadingPt(ELEVEN, FACE), 36.24);
  });

  // Word drew the outer bar 33.120 wide where the inner fraction is 30.817 across and
  // its own bar 30.960. **Nothing measured says where the other two points come from**,
  // and a fraction standing in a half is the only place the width of one shows.
  // **The bar above spans the box below, not the text below.** Word drew the outer
  // bar 33.120 wide where the inner fraction's own bar is 30.960, which is 2.16 more,
  // and the same 2.16 came back over a fraction nearly three times as wide.
  it("draws the outer bar to the inner fraction's whole box", () => {
    expect(outer.bar.widthPt).toBe(inner.widthPt);
    expectAsWordDrewIt(outer.bar.widthPt, 33.12);
    expectAsWordDrewIt(outer.bar.widthPt - inner.bar.widthPt, 2.16);
  });

  // At 20pt Word put the inner bar 56.40 wide and the outer 60.24, which is 1.92 out
  // on each side where 11pt gave 1.08. A fixed length would have given 1.08 again.
  it("stands out of the bar by a share of the em rather than a length", () => {
    const wide = fractionBox({
      ...halvesOf(TWENTY, 55.9, 50.828),
      sizePt: TWENTY,
      setting: "display",
      face: FACE,
    });

    expectAsWordDrewIt((wide.widthPt - wide.bar.widthPt) / 2, 1.92);
  });
});

// Page 9: `gr` over `preskadlim`, the pair that showed a fraction's height is measured
// off ink. The numerator's baseline came back at 57.360 and the denominator's at 73.200.
describe("a fraction whose numerator reaches no higher than the x-height", () => {
  const fraction = fractionBox({
    numerator: halfOf(ELEVEN, 11.985, X_HEIGHT, NUMERATOR_DESCENDER),
    denominator: halfOf(ELEVEN, 56.695, ASCENDER, DESCENDER),
    sizePt: ELEVEN,
    setting: "display",
    face: FACE,
  });

  it("stands the halves where Word stood them all the same", () => {
    expectAsWordDrewIt(fraction.numerator.baselinePt - fraction.denominator.baselinePt, 15.84);
    expectAsWordDrewIt(fraction.numerator.baselinePt - fraction.bar.topPt, 4.8);
  });

  // Word's ink ran 52.115 to 75.518, which is 23.403 against the 25.867 of the fraction
  // with an `l` in its numerator. **That difference is the whole reason the ink is
  // read**: nothing else about the two fractions differs.
  it("covers less ink for holding smaller letters", () => {
    expectAsWordDrewIt(fraction.ascentPt + fraction.descentPt, 23.403);
    const taller = fractionOf(ELEVEN, "display");
    expectAsWordDrewIt(
      taller.ascentPt + taller.descentPt - (fraction.ascentPt + fraction.descentPt),
      2.464,
    );
  });

  // Word's centres came back 305.891 and 305.924, a thirtieth of a point apart.
  it("centres the narrow half over the wide one", () => {
    const centre = (placed: { readonly leftPt: number; readonly widthPt: number }): number =>
      placed.leftPt + placed.widthPt / 2;
    expect(centre(fraction.numerator)).toBeCloseTo(centre(fraction.denominator), 10);
    expectAsWordDrewIt(fraction.bar.widthPt, 56.88);
  });

  // The paragraph came out 24.72 against the box's 23.30, which is the same leading
  // over a fraction two and a half points shorter.
  it("stands its line the face's own leading above its ink", () => {
    expectAsWordDrewIt(fraction.ascentPt + fraction.descentPt + mathLeadingPt(ELEVEN, FACE), 24.72);
  });
});

describe("a delimiter", () => {
  const fraction = fractionOf(ELEVEN, "display");
  const axisPt = inPoints(585, ELEVEN);

  const round = (content: MathBox, grows: boolean) =>
    delimiterBox({
      opening: OPENING,
      closing: CLOSING,
      content,
      sizePt: ELEVEN,
      grows,
      face: FACE,
    });

  // Page 10. Cambria's ladder reaches 10.23, 13.34, 18.22, 21.82 and 28.16pt at 11.04,
  // and the content stands 25.82 about the axis, so the fourth rung is the tallest that
  // does not overhang it. Word drew that one.
  it("grows to the rung Word grew it to, round a fraction", () => {
    const box = round(fraction, true);
    expect(box.setAsASubFormula).toBe(true);
    expect(box.opening?.variant?.measurement).toBe(4047);
    expectAsWordDrewIt((box.opening?.ascentPt ?? 0) + (box.opening?.descentPt ?? 0), 21.816);
    expectAsWordDrewIt(box.opening?.widthPt ?? 0, 5.454);
    expect(box.grownShort).toBe(false);
  });

  // Every rung above the first has its own ink centred on the axis to the unit, so the
  // grown parenthesis wants no shift at all; Word drew it one grid step below the
  // fraction's baseline.
  it("hangs what it grew on the axis rather than on the baseline", () => {
    const opening = round(fraction, true).opening;
    if (opening === null) throw new Error("no opening");
    expect(opening.baselinePt + (opening.ascentPt - opening.descentPt) / 2).toBeCloseTo(axisPt, 9);
    expectAsWordDrewIt(opening.baselinePt, -0.24);
  });

  // Page 12. Growth turned off gives the first rung, which is the character's own
  // glyph, and Word hung it 0.480 above the fraction's baseline: the first rung is the
  // one Cambria does not centre on the axis, by 88 units, which is 0.477pt at 11.04.
  it("draws the character's own glyph where the file turns growth off, hung on the axis", () => {
    const box = round(fraction, false);
    expect(box.opening?.variant?.measurement).toBe(1898);
    expectAsWordDrewIt((box.opening?.ascentPt ?? 0) + (box.opening?.descentPt ?? 0), 10.232);
    expectAsWordDrewIt(box.opening?.baselinePt ?? 0, 0.48);
    expectAsWordDrewIt(box.opening?.widthPt ?? 0, 4.582);
  });

  // Page 11. A parenthesis round a run came back in one string with the run, at the
  // run's own baseline: the content stands 10.12pt where the face asks for 16.17 before
  // it will set one as a sub-formula at all.
  it("draws a delimiter round a run as ordinary text, on the baseline", () => {
    const run = halfOf(ELEVEN, 30.817, ASCENDER, NUMERATOR_DESCENDER);
    const box = round(run, true);
    expect(box.setAsASubFormula).toBe(false);
    expect(box.opening?.variant).toBeNull();
    expect(box.opening?.baselinePt).toBe(0);
    expectAsWordDrewIt((box.opening?.ascentPt ?? 0) + (box.opening?.descentPt ?? 0), 10.232);
    expectAsWordDrewIt(box.opening?.widthPt ?? 0, 4.582);
  });

  /**
   * **A bracket stands against the content's ink, not against the box it advances
   * by**, which for a fraction are 1.08pt apart at 11pt.
   *
   * Measured on 2026-08-14 off Word's own pdf, over the four delimiters of the two
   * probes: page 10's grown parenthesis is drawn at 285.12 and advances 5.39, page
   * 12's ungrown one at 285.84 advancing 4.67, and **both end at 290.51**, which is
   * where the numerator starts. The bar begins at 290.40 and the box the fraction
   * advances by at 289.32, a whole point further left again. The shallow fraction and
   * the fraction of a fraction in the second probe give the same 1.08 to within a
   * fifth of a point.
   */
  it("lays a bracket against the content's ink rather than against its box", () => {
    const box = round(fraction, true);
    expect(box.opening?.leftPt).toBe(0);
    expect(box.content.leftPt + fraction.insetPt).toBeCloseTo(box.opening?.widthPt ?? 0, 10);
    expect(box.closing?.leftPt).toBeCloseTo(
      box.content.leftPt + fraction.widthPt - fraction.insetPt,
      10,
    );
    expect(box.widthPt).toBeCloseTo(
      (box.opening?.widthPt ?? 0) +
        fraction.widthPt +
        (box.closing?.widthPt ?? 0) -
        2 * fraction.insetPt,
      10,
    );
  });

  // A run keeps no room outside its own letters, so a bracket round one stands where
  // it always did.
  it("lays the content flush where it keeps no room of its own", () => {
    const run = halfOf(ELEVEN, 30.817, ASCENDER, NUMERATOR_DESCENDER);
    const box = round(run, true);
    expect(box.content.leftPt).toBe(box.opening?.widthPt);
    expect(box.widthPt).toBeCloseTo(
      (box.opening?.widthPt ?? 0) + run.widthPt + (box.closing?.widthPt ?? 0),
      10,
    );
  });

  // The line does not grow for a grown delimiter: page 10's paragraph is page 2's to
  // the hundredth, because a rung that does not overhang cannot reach past the ink it
  // stands round.
  it("takes no more room up or down than what it stands round", () => {
    const box = round(fraction, true);
    expect(box.ascentPt).toBeLessThanOrEqual(fraction.ascentPt);
    expect(box.descentPt).toBeLessThanOrEqual(fraction.descentPt);
  });

  it("draws nothing at an end the file leaves empty", () => {
    const box = delimiterBox({
      opening: null,
      closing: CLOSING,
      content: fraction,
      sizePt: ELEVEN,
      grows: true,
      face: FACE,
    });
    expect(box.opening).toBeNull();
    expect(box.content.leftPt).toBe(0);
  });

  // Cambria's ladder tops out at 47.9pt of reach at 11.04, and Word fills what is left
  // by assembling the pieces it states beside it: three of them, overlapping by 200
  // units. Nothing here draws one, so the tallest rung stands in and the box says it
  // was drawn short.
  it("says so where the content stands taller than the whole ladder", () => {
    const box = round({ widthPt: 20, ascentPt: 40, descentPt: 30, insetPt: 0 }, true);
    expect(box.grownShort).toBe(true);
    expect(box.opening?.variant?.measurement).toBe(8881);
  });
});

describe("the box a run of an equation covers", () => {
  it("takes its width from the advances and its ink from the letters", () => {
    const box = textBox("lrg", ELEVEN, FACE);
    expect(box.widthPt).toBeCloseTo(inPoints(2200, ELEVEN), 10);
    expectAsWordDrewIt(box.ascentPt, inPoints(ASCENDER, ELEVEN));
    expectAsWordDrewIt(box.descentPt, inPoints(447, ELEVEN));
  });

  // **This is why a fraction's height needs the glyph boxes and not the face's
  // ascent**, and page 9 against page 2 is the proof of it.
  it("stands lower for letters that reach less high", () => {
    const tall = textBox("lrg", ELEVEN, FACE);
    const short = textBox("rg", ELEVEN, FACE);
    expectAsWordDrewIt(tall.ascentPt - short.ascentPt, inPoints(ASCENDER - X_HEIGHT, ELEVEN));
  });

  it("takes a character that draws nothing for width without ink", () => {
    const box = textBox(" ", ELEVEN, FACE);
    expect(box.ascentPt).toBe(0);
    expect(box.descentPt).toBe(0);
    expect(box.widthPt).toBeCloseTo(inPoints(600, ELEVEN), 10);
  });

  it("answers an empty box for a character the face does not carry", () => {
    expect(textBox("א", ELEVEN, FACE)).toStrictEqual({
      widthPt: 0,
      ascentPt: 0,
      descentPt: 0,
      insetPt: 0,
    });
  });
});

describe("what the geometry asks of the face", () => {
  // Every constant named here is one the geometry reads, and the list is what a face
  // has to state for a fraction or a delimiter to be set at all.
  it("reads thirteen constants and no others", () => {
    const asked = new Set<string>();
    const stated = new Map<string, number>(Object.entries(FACE.constants));
    const watched = new Proxy(FACE.constants, {
      get: (_constants, name: string): number => {
        asked.add(name);
        return stated.get(name) ?? 0;
      },
    });

    const face = { ...FACE, constants: watched };
    const halves = { ...halvesOf(ELEVEN, 10, 10), sizePt: ELEVEN, face };
    fractionBox({ ...halves, setting: "display" });
    fractionBox({ ...halves, setting: "text" });
    scriptSizePt(ELEVEN, face);
    mathLeadingPt(ELEVEN, face);
    delimiterBox({
      opening: OPENING,
      closing: CLOSING,
      content: halfOf(ELEVEN, 10),
      sizePt: ELEVEN,
      grows: true,
      face,
    });

    expect([...asked].sort()).toStrictEqual([
      "axisHeight",
      "delimitedSubFormulaMinHeight",
      "fractionDenomDisplayStyleGapMin",
      "fractionDenominatorDisplayStyleShiftDown",
      "fractionDenominatorGapMin",
      "fractionDenominatorShiftDown",
      "fractionNumDisplayStyleGapMin",
      "fractionNumeratorDisplayStyleShiftUp",
      "fractionNumeratorGapMin",
      "fractionNumeratorShiftUp",
      "fractionRuleThickness",
      "mathLeading",
      "scriptPercentScaleDown",
    ]);
  });
});

// **What a renderer is handed, and the whole of it.** Everything else here answers
// where a thing stands; this says what the things are, so that nothing downstream has
// to know a fraction from a delimiter.
describe("what a set equation comes to on the page", () => {
  const MARK: ParagraphMark = {
    font: { kind: "named", name: "Cambria Math" },
    fontSizePt: ELEVEN,
    bold: false,
    italic: true,
    underline: false,
    raisePt: 0,
    lineSizePt: ELEVEN,
    lineRaisePt: 0,
    color: null,
    characterSpacingPt: 0,
    characterScale: 1,
    kernFromHalfPoints: null,
    highlight: null,
    capitals: "none",
  };

  const run = (text: string, widthPt: number, top = ASCENDER, bottom = DESCENDER): SetMath => ({
    kind: "run",
    text,
    mark: MARK,
    sizePt: ELEVEN,
    box: halfOf(ELEVEN, widthPt, top, bottom),
  });

  // Page 2 of the first probe, the fraction Word drew alone in its paragraph.
  const numerator = run("gralm", 30.817, ASCENDER, NUMERATOR_DESCENDER);
  const denominator = run("presk", 28.002);
  const fraction: SetMath = {
    kind: "fraction",
    mark: MARK,
    box: fractionBox({
      numerator: numerator.box,
      denominator: denominator.box,
      sizePt: ELEVEN,
      setting: "display",
      face: FACE,
    }),
    numerator: [numerator],
    denominator: [denominator],
  };

  // Word put the numerator's baseline at 60.000, the bar's top at 64.800, the
  // denominator's baseline at 75.840 and the bar's left at 290.400. The fraction's own
  // baseline stands 8.400 below the numerator's, and its box stands 1.08 left of its
  // bar.
  const AT = { leftPt: 290.4 - 1.08, baselinePt: 68.4 };

  it("hands out a fraction as two runs and a fill, in the order they are painted", () => {
    const drawn = mathPrimitivesOf([fraction], AT);

    expect(drawn.map((each) => each.kind)).toStrictEqual(["text", "fill", "text"]);
  });

  it("puts each of them where Word drew it", () => {
    const [top, bar, bottom] = mathPrimitivesOf([fraction], AT);
    if (top?.kind !== "text" || bar?.kind !== "fill" || bottom?.kind !== "text") {
      throw new Error("not the three expected");
    }

    expectAsWordDrewIt(top.baselinePt, 60.0);
    expectAsWordDrewIt(bar.topPt, 64.8);
    expectAsWordDrewIt(bottom.baselinePt, 75.84);
    expectAsWordDrewIt(bar.leftPt, 290.4);
    expectAsWordDrewIt(bar.widthPt, 30.96);
    expectAsWordDrewIt(bar.heightPt, 0.72);
    expectAsWordDrewIt(top.leftPt, 290.513);
    expectAsWordDrewIt(bottom.leftPt, 291.852);
  });

  // The page runs down and everything else in the file measures up, so the one place
  // the sign turns over is here.
  it("turns the rise of a half into a fall down the page", () => {
    const [top] = mathPrimitivesOf([fraction], AT);
    if (top?.kind !== "text") throw new Error("no numerator");

    expect(top.baselinePt).toBeLessThan(AT.baselinePt);
    expect(fraction.box.numerator.baselinePt).toBeGreaterThan(0);
  });

  it("carries the mark each thing is drawn in", () => {
    for (const each of mathPrimitivesOf([fraction], AT)) expect(each.mark).toBe(MARK);
  });

  // A delimiter that grew is a glyph of the face's own with no character to name it;
  // one that did not is the character itself and goes down as text.
  it("hands out a grown delimiter as a glyph and an ungrown one as text", () => {
    const round = (grows: boolean): SetMath => ({
      kind: "delimiter",
      mark: MARK,
      sizePt: ELEVEN,
      box: delimiterBox({
        opening: OPENING,
        closing: CLOSING,
        content: fraction.box,
        sizePt: ELEVEN,
        grows,
        face: FACE,
      }),
      content: [fraction],
    });

    const grown = mathPrimitivesOf([round(true)], AT);
    expect(grown.map((each) => each.kind)).toStrictEqual([
      "glyph",
      "text",
      "fill",
      "text",
      "glyph",
    ]);
    const opening = grown[0];
    if (opening?.kind !== "glyph") throw new Error("no opening");
    expect(opening.glyph).toBe(FACE.tallerVariantsOf(OPENING)[3]?.glyph);

    const plain = mathPrimitivesOf([round(false)], AT);
    expect(plain.map((each) => each.kind)).toStrictEqual(["text", "text", "fill", "text", "text"]);
  });

  it("lays one piece after another along the line", () => {
    const first = run("a", 10);
    const second = run("b", 20);
    const drawn = mathPrimitivesOf([first, second], AT);

    expect(drawn.map((each) => (each.kind === "text" ? each.leftPt : 0))).toStrictEqual([
      AT.leftPt,
      AT.leftPt + 10,
    ]);
  });

  it("hands out nothing for an equation holding nothing", () => {
    expect(mathPrimitivesOf([], AT)).toStrictEqual([]);
  });
});

// **Which of the two sets of constants each structure is set with**, which is the one
// thing about a nested fraction that the geometry cannot be asked one box at a time.
describe("setMath choosing the setting", () => {
  const MARK: ParagraphMark = {
    font: { kind: "named", name: "Cambria Math" },
    fontSizePt: ELEVEN,
    bold: false,
    italic: true,
    underline: false,
    raisePt: 0,
    lineSizePt: ELEVEN,
    lineRaisePt: 0,
    color: null,
    characterSpacingPt: 0,
    characterScale: 1,
    kernFromHalfPoints: null,
    highlight: null,
    capitals: "none",
  };

  const half = (text: string): MarkedMath => ({ kind: "run", text, mark: MARK });

  const over = (numerator: string, denominator: string): MarkedMath => ({
    kind: "fraction",
    mark: MARK,
    numerator: [half(numerator)],
    denominator: [half(denominator)],
  });

  const set = (pieces: readonly MarkedMath[]): readonly SetMath[] => {
    const setting = setMath(pieces, {
      sizePt: ELEVEN,
      halfSizePt: ELEVEN,
      setting: "display",
      face: FACE,
      measure: (text, _mark, sizePt) => textBox(text, sizePt, FACE),
    });
    if (setting === null) throw new Error("nothing was set");
    return setting;
  };

  const apart = (piece: SetMath | undefined): number => {
    if (piece?.kind !== "fraction") throw new Error("no fraction");
    return piece.box.numerator.baselinePt - piece.box.denominator.baselinePt;
  };

  // Page 13 of the first probe: the outer pair's baselines stood 26.88 apart and the
  // inner pair's 12.24, which is the text shifts against the display ones.
  it("sets a fraction inside another in the text constants", () => {
    const outer = set([
      { kind: "fraction", mark: MARK, numerator: [over("lrg", "lrg")], denominator: [half("lrg")] },
    ])[0];
    if (outer?.kind !== "fraction") throw new Error("no fraction");

    expectAsWordDrewIt(apart(outer.numerator[0]), 12.24);
  });

  // Page 10 of the first probe: the fraction inside the delimiter kept the display
  // pair's 15.84, the same as the one standing alone.
  it("passes the setting through a delimiter untouched", () => {
    const round: MarkedMath = {
      kind: "delimiter",
      mark: MARK,
      opening: OPENING,
      closing: CLOSING,
      grows: true,
      content: [over("gralm", "presk")],
    };
    const held = set([round])[0];
    if (held?.kind !== "delimiter") throw new Error("no delimiter");

    expectAsWordDrewIt(apart(held.content[0]), 15.84);
  });
});

/**
 * **Word spaces its operators, and the case is read off the left edge rather than off
 * the row.**
 *
 * Cases F to L of `equation-content-probe`, three repeats each, read out of Word's own
 * pdf on 2026-08-14. Every one of them stands alone in its paragraph, so Word centres it
 * on the body's own centre of 306.00 and **twice the distance from the left edge Word
 * drew to that centre is what the row advanced by**. That reading is exact where the
 * row's own reported width is not: it stops at the last letter's ink and leaves out the
 * correction standing after it.
 *
 * The face here carries Cambria Math's own advances and italic corrections for the
 * seven characters the cases hold, read off the file on 2026-08-14. The ink is invented,
 * since none of these asks a question about a height.
 */
describe("the spacing Word puts round an operator", () => {
  // Cambria Math's own, read off the file on 2026-08-14 and 2026-08-16. The face states
  // a correction for the four letters and for none of the rest.
  const ADVANCES = {
    "𝑎": 1141, "𝑏": 1104, "𝑐": 942, "𝑑": 1187,
    "−": 1530, "+": 1530, "×": 1463, "÷": 1530, "⋅": 578,
    "=": 1530, "<": 1534, "≤": 1534, "≠": 1530, "≈": 1507,
    ",": 420, ".": 420, ":": 540,
    "/": 1004, "(": 850, ")": 850, "2": 1134, " ": 451,
  }; // prettier-ignore
  const CORRECTIONS = { "𝑎": 50, "𝑏": 45, "𝑐": 65, "𝑑": 65 };

  const SPACED = faceHolding({
    cmapFormat: 12,
    advances: ADVANCES,
    boxes: Object.fromEntries(
      Object.entries(ADVANCES).map(([character, advance]) => [
        character,
        { left: 0, bottom: 0, right: advance, top: X_HEIGHT },
      ]),
    ),
    math: { constants: CAMBRIA_MATH, italicCorrections: CORRECTIONS },
  });

  const STATED = 11;

  const MARK: ParagraphMark = {
    font: { kind: "named", name: "Cambria Math" },
    fontSizePt: STATED,
    bold: false,
    italic: true,
    underline: false,
    raisePt: 0,
    lineSizePt: STATED,
    lineRaisePt: 0,
    color: null,
    characterSpacingPt: 0,
    characterScale: 1,
    kernFromHalfPoints: null,
    highlight: null,
    capitals: "none",
  };

  const set = (pieces: readonly MarkedMath[]): readonly SetMath[] => {
    const setting = setMath(pieces, {
      sizePt: STATED,
      halfSizePt: STATED,
      setting: "display",
      face: SPACED,
      measure: (text, _mark, sizePt) => textBox(text, sizePt, SPACED),
    });
    if (setting === null) throw new Error("nothing was set");
    return setting;
  };

  const rowPt = (text: string): number =>
    mathRowOf(set([{ kind: "run", text, mark: MARK }])).widthPt;

  // Word's own left edges, and the row each of them says was laid out: 306.00 less the
  // edge, twice over. The last decimal is the pdf's own and not a rule, so a case holds
  // to the hundredth.
  const asWordLaidItOut = (actual: number, leftPt: number): void => {
    expect(actual).toBeCloseTo(2 * (306 - leftPt), 2);
  };

  it("advances two letters by their own advances and the correction after the last", () => {
    asWordLaidItOut(rowPt("𝑎𝑏"), 299.85);
  });

  it("leaves 4/18 of the em on each side of an operation", () => {
    asWordLaidItOut(rowPt("𝑎−𝑏"), 293.163);
    asWordLaidItOut(rowPt("𝑎×𝑏"), 293.342);
  });

  it("leaves 5/18 of the em on each side of a relation", () => {
    asWordLaidItOut(rowPt("𝑎=𝑏"), 292.551);
  });

  // Case K against case G: the same expression with a space either side of the operator
  // in the file came out exactly two space advances wider, so Word's own spacing stands
  // on top of the file's rather than instead of it.
  it("stands on top of the spaces the file states", () => {
    asWordLaidItOut(rowPt("𝑎 − 𝑏"), 290.74);
    expect(rowPt("𝑎 − 𝑏") - rowPt("𝑎−𝑏")).toBeCloseTo((2 * 451 * STATED) / UNITS_PER_EM, 6);
  });

  // The italic correction is the face's own for the character standing before the gap,
  // which is why `ab` and `cd` differ by more than their advances: 45 units against 65.
  it("takes the correction from the character it stands after", () => {
    expect(rowPt("𝑐𝑑")).toBeCloseTo(((942 + 1187 + 65) * STATED) / UNITS_PER_EM, 6);
    expect(rowPt("𝑎𝑏")).toBeCloseTo(((1141 + 1104 + 45) * STATED) / UNITS_PER_EM, 6);
  });

  /**
   * The twenty cases of `equation-spacing-probe`, read out of Word's own pdf on
   * 2026-08-16, three repeats each. **Every number here is what Word laid the row out
   * as**, read off the left edge of a numerator it centred on 306.00, and every one of
   * them falls out of the rules above and the face's own advances to a thousandth.
   *
   * The row Word drew is what the case holds, so a case that disagrees names the rule it
   * disagrees with and no arithmetic of this file's own stands in between.
   */
  const asWordDrewTheRow = (text: string, rowPt_: number): void => {
    expect(rowPt(text)).toBeCloseTo(rowPt_, 2);
  };

  it("spaces a plus, a division sign and a dot operator as operations", () => {
    asWordDrewTheRow("𝑎+𝑏", 25.675);
    asWordDrewTheRow("𝑎÷𝑏", 25.675);
    asWordDrewTheRow("𝑎⋅𝑏", 20.561);
  });

  it("spaces the four inequalities asked as relations", () => {
    asWordDrewTheRow("𝑎<𝑏", 26.919);
    asWordDrewTheRow("𝑎≤𝑏", 26.919);
    asWordDrewTheRow("𝑎≠𝑏", 26.897);
    asWordDrewTheRow("𝑎≈𝑏", 26.774);
  });

  // Neither of the two spacings above: 1.834 at 11pt is 3/18 of the em, where an
  // operation is 4/18 either side and a relation 5/18.
  it("gives punctuation a third spacing of its own", () => {
    asWordDrewTheRow("𝑎,𝑏", 16.658);
    asWordDrewTheRow("𝑎.𝑏", 16.658);
    asWordDrewTheRow("𝑎:𝑏", 17.302);
  });

  // **What Word left alone is a measurement too.** Both of these came back at their own
  // advances and their corrections, so neither character is in the table.
  it("leaves a solidus and a bracket unspaced", () => {
    asWordDrewTheRow("𝑎/𝑏", 17.961);
    asWordDrewTheRow("𝑎(𝑏)", 21.699);
  });

  // **A letter or a digit is what closes a correction, and nothing else is.** `𝑎2` is
  // the two advances with `𝑎`'s 50 units nowhere in it; `2𝑎` keeps the same 50 at the
  // end of the row; and `𝑎(𝑏)` above keeps both letters' corrections in front of a
  // bracket, which the face states no correction for any more than it does for a digit.
  it("closes a correction with a digit and not with a bracket", () => {
    asWordDrewTheRow("𝑎2", 12.219);
    asWordDrewTheRow("2𝑎", 12.488);
  });

  // **An operation opening a row is the sign of what follows it and takes nothing**, and
  // so is one standing after a relation, which is the configuration a real document is
  // full of. **One closing a row keeps its left gap**, where TeX demotes it: `𝑎−` came
  // back 2.444 wider than its own advances, which is one 4/18 and not two. So Word looks
  // leftwards to decide and TeX looks both ways.
  it("spaces an operation that closes a row and not one that opens it", () => {
    asWordDrewTheRow("−𝑏", 14.389);
    asWordDrewTheRow("𝑎−", 17.059);
    asWordDrewTheRow("𝑎=−𝑏", 35.115);
  });

  // Case L, the same expression in a numerator: Word filled a bar 25.680 wide, which is
  // the row above snapped onto its own 0.24 grid.
  it("spans a fraction's bar with the half the spacing widened", () => {
    const fraction = set([
      {
        kind: "fraction",
        mark: MARK,
        numerator: [{ kind: "run", text: "𝑎−𝑏", mark: MARK }],
        denominator: [{ kind: "run", text: "𝑐𝑑", mark: MARK }],
      },
    ])[0];
    if (fraction?.kind !== "fraction") throw new Error("no fraction");
    expectAsWordDrewIt(fraction.box.bar.widthPt, 25.68);
  });

  // A gap is room and nothing else: the letters either side of it are drawn where the
  // row puts them, and nothing at all is drawn in between.
  it("draws nothing for the room it leaves", () => {
    const drawn = mathPrimitivesOf(set([{ kind: "run", text: "𝑎−𝑏", mark: MARK }]), {
      leftPt: 0,
      baselinePt: 0,
    });
    expect(drawn.map((each) => each.kind)).toStrictEqual(["text", "text", "text"]);
    expect(drawn[1]?.leftPt).toBeCloseTo(
      ((1141 + 50) * STATED) / UNITS_PER_EM + (4 / 18) * STATED,
      6,
    );
  });
});
