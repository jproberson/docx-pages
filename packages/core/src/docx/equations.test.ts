import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument, WORDPROCESSING_NS } from "../testing/build-docx.js";
import {
  DEFAULT_MATH_FONT,
  equationsIn,
  markedAsMath,
  mathStyleOf,
  readEquation,
  readMathFont,
  runsAlone,
  type Equation,
  type EquationDelimiter,
  type EquationFraction,
  type EquationPiece,
  type EquationRun,
  type MathStyle,
} from "./equations.js";
import { readUnhonoured } from "./fidelity.js";
import { openDocx, partXml, type DocxPackage } from "./package.js";
import type { ParagraphMark } from "./styles.js";
import { firstNamed, type XmlElement } from "./xml.js";

const MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const SECTION = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;

const packageOf = (body: string, parts: Readonly<Record<string, string>> = {}): DocxPackage =>
  openDocx(buildDocx({ "word/document.xml": wordDocument(`${body}${SECTION}`), ...parts }));

// The one equation a body holds, found as the layout would reach it.
function equationIn(body: string): XmlElement {
  const declared = body.replace(/<m:oMath(Para)?/, `<m:oMath$1 xmlns:m="${MATH_NS}"`);
  const found = equationsIn(partXml(packageOf(declared), "word/document.xml"));
  expect(found).toHaveLength(1);
  const [only] = found;
  if (only === undefined) throw new Error("no equation");
  return only;
}

const equationOf = (math: string): Equation => readEquation(equationIn(`<w:p>${math}</w:p>`));

const contentOf = (equation: Equation): readonly EquationPiece[] =>
  equation.kind === "read" ? equation.content : [];

// The text of every run in something, at whatever depth it stands, so a case can say
// what was read without writing the tree out again.
function textIn(pieces: readonly EquationPiece[]): readonly string[] {
  return pieces.flatMap((piece) => {
    if (piece.kind === "run") return [piece.text];
    if (piece.kind === "break") return [];
    if (piece.kind === "fraction")
      return [...textIn(piece.numerator), ...textIn(piece.denominator)];
    return piece.parts.flatMap(textIn);
  });
}

const textOf = (equation: Equation): readonly string[] => textIn(contentOf(equation));

// What the reader made of the one run in an equation, or the refusal instead of it.
function firstRunOf(equation: Equation): EquationRun {
  const [only] = contentOf(equation);
  if (equation.kind === "refused") throw new Error(equation.unreadable.join(" "));
  if (only === undefined || only.kind !== "run") throw new Error("no run");
  return only;
}

// The one fraction or the one delimiter an equation holds, and nothing else.
function onlyPieceOf(equation: Equation): EquationPiece {
  if (equation.kind === "refused") throw new Error(equation.unreadable.join(" "));
  expect(equation.content).toHaveLength(1);
  const [only] = equation.content;
  if (only === undefined) throw new Error("nothing read");
  return only;
}

function fractionOf(equation: Equation): EquationFraction {
  const only = onlyPieceOf(equation);
  if (only.kind !== "fraction") throw new Error(`a ${only.kind}`);
  return only;
}

function delimiterOf(equation: Equation): EquationDelimiter {
  const only = onlyPieceOf(equation);
  if (only.kind !== "delimiter") throw new Error(`a ${only.kind}`);
  return only;
}

const mathRun = (text: string, properties = "", math = ""): string =>
  `<m:r>${math}${properties}<m:t>${text}</m:t></m:r>`;

const FRACTION =
  `<m:f><m:fPr><m:ctrlPr/></m:fPr>` +
  `<m:num>${mathRun("one")}</m:num><m:den>${mathRun("two")}</m:den></m:f>`;

describe("readEquation", () => {
  it("reads an equation whose content is runs of text, in the order it holds them", () => {
    const equation = equationOf(`<m:oMath>${mathRun("alpha")}${mathRun("beta")}</m:oMath>`);
    expect(equation.kind).toBe("read");
    expect(textOf(equation)).toStrictEqual(["alpha", "beta"]);
  });

  // Word hands the reader the `m:r` rather than a copy of what it says, so the run's
  // own `w:rPr` is resolved by the cascade that resolves every other run's.
  it("answers with the run itself, holding the properties the file wrote on it", () => {
    const properties = `<w:rPr><w:sz w:val="20"/></w:rPr>`;
    const run = firstRunOf(equationOf(`<m:oMath>${mathRun("gamma", properties)}</m:oMath>`));
    expect(run.element.name).toBe("r");
    expect(run.element.namespace).toBe(MATH_NS);
    expect(firstNamed(run.element, WORDPROCESSING_NS, "rPr")).not.toBeNull();
  });

  it("draws a run stating no style of its own slanted, which is what Word does", () => {
    expect(firstRunOf(equationOf(`<m:oMath>${mathRun("delta")}</m:oMath>`)).style).toBe("italic");
  });

  it("reads the style a run does state, and every value of it", () => {
    const styled = (value: string): MathStyle | null => {
      const stated = `<m:rPr><m:sty m:val="${value}"/></m:rPr>`;
      return firstRunOf(equationOf(`<m:oMath>${mathRun("epsilon", "", stated)}</m:oMath>`)).style;
    };
    expect(styled("p")).toBe("plain");
    expect(styled("b")).toBe("bold");
    expect(styled("i")).toBe("italic");
    expect(styled("bi")).toBe("bold-italic");
  });

  // `m:nor` is ordinary text standing inside an equation, which Word styles as an
  // equation in no way at all.
  it("takes a run marked as ordinary text for one the equation styles nowhere", () => {
    const ordinary = `<m:rPr><m:nor/></m:rPr>`;
    expect(
      firstRunOf(equationOf(`<m:oMath>${mathRun("zeta", "", ordinary)}</m:oMath>`)).style,
    ).toBeNull();
  });

  it("keeps the space an equation asks to keep, and drops the one it does not", () => {
    const asked = `<m:r><m:t xml:space="preserve">eta </m:t></m:r>`;
    expect(textOf(equationOf(`<m:oMath>${asked}</m:oMath>`))).toStrictEqual(["eta "]);
    expect(textOf(equationOf(`<m:oMath><m:r><m:t> eta </m:t></m:r></m:oMath>`))).toStrictEqual([
      "eta",
    ]);
  });

  // An equation holding nothing draws nothing and costs the page nothing, so there is
  // nothing to refuse in it.
  it("reads an equation holding nothing as no content at all", () => {
    expect(equationOf(`<m:oMath></m:oMath>`)).toStrictEqual({ kind: "read", content: [] });
  });

  it("passes over a bookmark, which draws nothing wherever it stands", () => {
    const bookmarked = `<w:bookmarkStart w:id="1" w:name="a"/>${mathRun("theta")}<w:bookmarkEnd w:id="1"/>`;
    expect(textOf(equationOf(`<m:oMath>${bookmarked}</m:oMath>`))).toStrictEqual(["theta"]);
  });

  it("refuses a script and an array, which are set in ways nothing here reads", () => {
    const refusalOf = (math: string): unknown => equationOf(`<m:oMath>${math}</m:oMath>`);
    expect(refusalOf(`<m:sSup><m:e>${mathRun("k")}</m:e><m:sup/></m:sSup>`)).toStrictEqual({
      kind: "refused",
      unreadable: ["m:sSup"],
    });
    expect(refusalOf(`<m:eqArr><m:e>${mathRun("l")}</m:e></m:eqArr>`)).toStrictEqual({
      kind: "refused",
      unreadable: ["m:eqArr"],
    });
  });

  // A run that remaps its characters into another alphabet draws none of the ones the
  // file spells, so what it holds cannot be read off the text at all.
  it("refuses a run whose properties ask for more than a style", () => {
    const scripted = `<m:rPr><m:scr m:val="double-struck"/></m:rPr>`;
    expect(equationOf(`<m:oMath>${mathRun("mu", "", scripted)}</m:oMath>`)).toStrictEqual({
      kind: "refused",
      unreadable: ["m:scr"],
    });
  });

  // A refusal is what the report reads, so it names each thing once however many of
  // them stand in the equation, and reads the same way twice.
  it("names everything in the way once, in one order", () => {
    const cluttered = `<m:rad/>${FRACTION}<m:nary/><m:rad/>`;
    expect(equationOf(`<m:oMath>${cluttered}</m:oMath>`)).toStrictEqual({
      kind: "refused",
      unreadable: ["m:nary", "m:rad"],
    });
  });

  // A refusal anywhere refuses the whole equation: half a fraction drawn where Word
  // draws a whole one is the plausible-looking page, and worse than none.
  it("refuses the whole equation for something deep inside a fraction", () => {
    const buried = `<m:f><m:num>${mathRun("one")}</m:num><m:den><m:rad/></m:den></m:f>`;
    expect(equationOf(`<m:oMath>${buried}</m:oMath>`)).toStrictEqual({
      kind: "refused",
      unreadable: ["m:rad"],
    });
  });

  it("finds the equation inside the wrapper Word writes round one standing alone", () => {
    const alone = `<m:oMathPara><m:oMath>${mathRun("nu")}</m:oMath></m:oMathPara>`;
    expect(textOf(readEquation(equationIn(`<w:p>${alone}</w:p>`)))).toStrictEqual(["nu"]);
  });
});

// Three documents of one template end an equation with a run holding nothing but a
// break, which ends the line the equation stands on.
describe("a break inside an equation", () => {
  const kindsIn = (pieces: readonly EquationPiece[]): readonly string[] =>
    pieces.map((piece) => piece.kind);

  it("reads a run holding nothing but a break, and answers with the run it stands in", () => {
    const broken = `${mathRun("a")}<m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:br/></m:r>`;
    const content = contentOf(equationOf(`<m:oMath>${broken}</m:oMath>`));
    expect(kindsIn(content)).toStrictEqual(["run", "break"]);
    const ending = content[1];
    expect(ending?.kind === "break" && ending.element.name).toBe("r");
  });

  // A break in the middle of a run ends the line between the two halves of its text,
  // which is what reading the properties before the content is for.
  it("ends the line where the break stands rather than after the run", () => {
    const broken = `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>a</m:t><w:br/><m:t>b</m:t></m:r>`;
    const content = contentOf(equationOf(`<m:oMath>${broken}</m:oMath>`));
    expect(kindsIn(content)).toStrictEqual(["run", "break", "run"]);
    expect(textIn(content)).toStrictEqual(["a", "b"]);
    for (const piece of content) if (piece.kind === "run") expect(piece.style).toBe("plain");
  });

  // Where a page or a column starts inside an equation is a place nothing has
  // measured, so it is refused rather than read as an ending line.
  it("refuses a break that starts a page or a column", () => {
    for (const type of ["page", "column"]) {
      const broken = `<m:r><w:br w:type="${type}"/></m:r>`;
      expect(equationOf(`<m:oMath>${broken}</m:oMath>`)).toStrictEqual({
        kind: "refused",
        unreadable: ["w:br"],
      });
    }
  });
});

describe("a fraction", () => {
  it("reads its halves, each holding its own content", () => {
    const fraction = fractionOf(equationOf(`<m:oMath>${FRACTION}</m:oMath>`));
    expect(textIn(fraction.numerator)).toStrictEqual(["one"]);
    expect(textIn(fraction.denominator)).toStrictEqual(["two"]);
  });

  // The bar is drawn from the properties the file writes for it, and a `m:ctrlPr` holds
  // its `w:rPr` where a run holds one, so the cascade marks it unchanged.
  it("answers with the properties the file writes for its bar", () => {
    const fraction = fractionOf(equationOf(`<m:oMath>${FRACTION}</m:oMath>`));
    expect(fraction.control?.name).toBe("ctrlPr");
    const bare = `<m:f><m:num>${mathRun("one")}</m:num><m:den>${mathRun("two")}</m:den></m:f>`;
    expect(fractionOf(equationOf(`<m:oMath>${bare}</m:oMath>`)).control).toBeNull();
  });

  it("reads one whose halves are empty as two halves holding nothing", () => {
    const empty = `<m:f><m:num/><m:den/></m:f>`;
    const fraction = fractionOf(equationOf(`<m:oMath>${empty}</m:oMath>`));
    expect(fraction.numerator).toStrictEqual([]);
    expect(fraction.denominator).toStrictEqual([]);
  });

  // A fraction is stacked over a bar unless the file says otherwise. The three others
  // it can say are set differently enough that reading one as a stack draws the wrong
  // thing, and no document met here states any of them.
  it("reads the stacked kind, stated or not, and refuses the three others", () => {
    const typed = (type: string): string =>
      `<m:f><m:fPr><m:type m:val="${type}"/></m:fPr>` +
      `<m:num>${mathRun("one")}</m:num><m:den>${mathRun("two")}</m:den></m:f>`;
    expect(textOf(equationOf(`<m:oMath>${typed("bar")}</m:oMath>`))).toStrictEqual(["one", "two"]);
    for (const type of ["skw", "lin", "noBar"]) {
      expect(equationOf(`<m:oMath>${typed(type)}</m:oMath>`)).toStrictEqual({
        kind: "refused",
        unreadable: ["m:type"],
      });
    }
  });

  it("refuses a property of a fraction it does not read", () => {
    const sized = `<m:f><m:fPr><m:ctrlPr/></m:fPr><m:num><m:argPr/></m:num><m:den/></m:f>`;
    expect(equationOf(`<m:oMath>${sized}</m:oMath>`)).toStrictEqual({
      kind: "refused",
      unreadable: ["m:argPr"],
    });
  });
});

describe("a delimiter", () => {
  const DELIMITED = `<m:d><m:dPr><m:ctrlPr/></m:dPr><m:e>${mathRun("iota")}</m:e></m:d>`;

  it("reads its parts, each holding its own content", () => {
    const delimiter = delimiterOf(equationOf(`<m:oMath>${DELIMITED}</m:oMath>`));
    expect(delimiter.parts).toHaveLength(1);
    expect(delimiter.parts.map(textIn)).toStrictEqual([["iota"]]);
    expect(delimiter.control?.name).toBe("ctrlPr");
  });

  // Word draws a delimiter stating none of its characters in parentheses, and divides
  // one part from the next with a bar.
  it("draws a delimiter stating nothing in parentheses", () => {
    const delimiter = delimiterOf(equationOf(`<m:oMath>${DELIMITED}</m:oMath>`));
    expect(delimiter.opening).toBe("(");
    expect(delimiter.closing).toBe(")");
    expect(delimiter.separator).toBe("|");
  });

  it("reads the characters a delimiter does state", () => {
    const stated =
      `<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/><m:sepChr m:val=","/></m:dPr>` +
      `<m:e>${mathRun("iota")}</m:e><m:e>${mathRun("kappa")}</m:e></m:d>`;
    const delimiter = delimiterOf(equationOf(`<m:oMath>${stated}</m:oMath>`));
    expect(delimiter.opening).toBe("[");
    expect(delimiter.closing).toBe("]");
    expect(delimiter.separator).toBe(",");
    expect(delimiter.parts.map(textIn)).toStrictEqual([["iota"], ["kappa"]]);
  });

  // An empty character is how a bracket open at one end is written, and Word draws
  // nothing at all on that side rather than a character of some other kind.
  it("draws nothing where a delimiter states an empty character", () => {
    const open = `<m:d><m:dPr><m:begChr m:val=""/><m:endChr/></m:dPr><m:e/></m:d>`;
    const delimiter = delimiterOf(equationOf(`<m:oMath>${open}</m:oMath>`));
    expect(delimiter.opening).toBeNull();
    expect(delimiter.closing).toBeNull();
  });

  // Whether a delimiter stating nothing stretches to its content is unmeasured, so the
  // reader says what the file says and leaves the answer to whatever measured it.
  it("says whether the file asked it to grow, and says nothing where the file does not", () => {
    const growing = (stated: string): boolean | null =>
      delimiterOf(equationOf(`<m:oMath><m:d><m:dPr>${stated}</m:dPr><m:e/></m:d></m:oMath>`)).grows;
    expect(growing(`<m:grow/>`)).toBe(true);
    expect(growing(`<m:grow m:val="1"/>`)).toBe(true);
    expect(growing(`<m:grow m:val="0"/>`)).toBe(false);
    expect(growing("")).toBeNull();
  });

  it("refuses a property of a delimiter it does not read", () => {
    const shaped = `<m:d><m:dPr><m:shp m:val="match"/></m:dPr><m:e/></m:d>`;
    expect(equationOf(`<m:oMath>${shaped}</m:oMath>`)).toStrictEqual({
      kind: "refused",
      unreadable: ["m:shp"],
    });
  });

  // The shape 22 of the equations met in real documents are: a fraction whose halves
  // hold runs and parenthesised groups, which the one walk reads at any depth.
  it("reads a fraction inside a delimiter inside a fraction", () => {
    const inner = `<m:f><m:num>${mathRun("one")}</m:num><m:den>${mathRun("two")}</m:den></m:f>`;
    const nested =
      `<m:f><m:num>${mathRun("a")}<m:d><m:e>${inner}</m:e></m:d></m:num>` +
      `<m:den>${mathRun("b")}</m:den></m:f>`;
    const fraction = fractionOf(equationOf(`<m:oMath>${nested}</m:oMath>`));
    expect(textIn(fraction.numerator)).toStrictEqual(["a", "one", "two"]);
    expect(textIn(fraction.denominator)).toStrictEqual(["b"]);
    const [, group] = fraction.numerator;
    expect(group?.kind).toBe("delimiter");
  });
});

// What the layout draws where a paragraph's own runs go, and what still needs setting.
describe("runsAlone", () => {
  it("answers the runs of an equation that is nothing else", () => {
    const runs = runsAlone(equationOf(`<m:oMath>${mathRun("a")}${mathRun("b")}</m:oMath>`));
    expect(runs?.map((run) => run.text)).toStrictEqual(["a", "b"]);
  });

  it("answers nothing for an equation holding something that has to be set", () => {
    expect(runsAlone(equationOf(`<m:oMath>${FRACTION}</m:oMath>`))).toBeNull();
    const delimited = `<m:d><m:e>${mathRun("iota")}</m:e></m:d>`;
    expect(runsAlone(equationOf(`<m:oMath>${delimited}</m:oMath>`))).toBeNull();
    expect(runsAlone(equationOf(`<m:oMath><m:rad/></m:oMath>`))).toBeNull();
    // A break is read, and where a line ends inside an equation is unmeasured, so it
    // is not one of the ones drawn where a paragraph's own runs are.
    expect(runsAlone(equationOf(`<m:oMath><m:r><w:br/></m:r></m:oMath>`))).toBeNull();
  });

  it("answers no runs for an equation holding nothing, which draws nothing", () => {
    expect(runsAlone(equationOf(`<m:oMath/>`))).toStrictEqual([]);
  });
});

// The layout marks a run where it meets one, so what it asks of a run on its own has
// to be what reading the whole equation said about it.
describe("mathStyleOf", () => {
  const styleAskedOfTheRunAlone = (math: string): MathStyle | null =>
    mathStyleOf(firstRunOf(equationOf(`<m:oMath>${math}</m:oMath>`)).element);

  it("answers what reading the equation answered, for every style a run states", () => {
    expect(styleAskedOfTheRunAlone(mathRun("omicron"))).toBe("italic");
    const stated = (value: string): string =>
      mathRun("pi", "", `<m:rPr><m:sty m:val="${value}"/></m:rPr>`);
    expect(styleAskedOfTheRunAlone(stated("p"))).toBe("plain");
    expect(styleAskedOfTheRunAlone(stated("bi"))).toBe("bold-italic");
    expect(styleAskedOfTheRunAlone(mathRun("rho", "", `<m:rPr><m:nor/></m:rPr>`))).toBeNull();
  });
});

describe("readMathFont", () => {
  const settingsHolding = (mathPr: string): string =>
    `<?xml version="1.0"?><w:settings xmlns:w="${WORDPROCESSING_NS}" xmlns:m="${MATH_NS}">
      ${mathPr}</w:settings>`;

  it("answers Cambria Math for a document that states no face of its own", () => {
    expect(readMathFont(packageOf(`<w:p/>`))).toBe(DEFAULT_MATH_FONT);
    expect(readMathFont(packageOf(`<w:p/>`, { "word/settings.xml": settingsHolding("") }))).toBe(
      DEFAULT_MATH_FONT,
    );
  });

  it("answers the face a document does state", () => {
    const stated = `<m:mathPr><m:mathFont m:val="Latin Modern Math"/></m:mathPr>`;
    expect(
      readMathFont(packageOf(`<w:p/>`, { "word/settings.xml": settingsHolding(stated) })),
    ).toBe("Latin Modern Math");
  });
});

describe("markedAsMath", () => {
  const MARK: ParagraphMark = {
    font: { kind: "unresolved" },
    fontSizePt: 11,
    bold: false,
    italic: false,
    underline: false,
    raisePt: 0,
    lineSizePt: 11,
    lineRaisePt: 0,
    color: null,
    characterSpacingPt: 0,
    characterScale: 1,
    kernFromHalfPoints: null,
  };

  it("sets an equation in the document's own face where the run names none", () => {
    expect(markedAsMath(MARK, "italic", "Cambria Math").font).toStrictEqual({
      kind: "named",
      name: "Cambria Math",
    });
  });

  // Word writes the face onto the run as well as into the settings, and what the run
  // says is what it is drawn in.
  it("leaves the face a run does name alone", () => {
    const named = { ...MARK, font: { kind: "named", name: "Arial" } as const };
    expect(markedAsMath(named, "italic", "Cambria Math").font).toStrictEqual({
      kind: "named",
      name: "Arial",
    });
  });

  it("slants and weights what the style states, and nothing it does not", () => {
    expect(markedAsMath(MARK, "italic", DEFAULT_MATH_FONT)).toMatchObject({
      bold: false,
      italic: true,
    });
    expect(markedAsMath(MARK, "bold-italic", DEFAULT_MATH_FONT)).toMatchObject({
      bold: true,
      italic: true,
    });
    expect(markedAsMath(MARK, "plain", DEFAULT_MATH_FONT)).toMatchObject({
      bold: false,
      italic: false,
    });
  });

  it("leaves ordinary text standing in an equation exactly as it was", () => {
    expect(markedAsMath(MARK, null, DEFAULT_MATH_FONT)).toStrictEqual(MARK);
  });
});

describe("what the report says about an equation", () => {
  const kindsOf = (math: string): readonly string[] => {
    const declared = math.replace("<m:oMath", `<m:oMath xmlns:m="${MATH_NS}"`);
    return readUnhonoured(packageOf(`<w:p>${declared}</w:p>`)).map((entry) => entry.kind);
  };

  it("says nothing about an equation of runs, which is drawn where a run is drawn", () => {
    expect(kindsOf(`<m:oMath>${mathRun("xi")}</m:oMath>`)).toStrictEqual([]);
  });

  it("still names the one it cannot read", () => {
    expect(kindsOf(`<m:oMath><m:rad/></m:oMath>`)).toStrictEqual(["equation"]);
  });

  // **Read is not drawn.** The shape of a fraction and of a delimiter is read now and
  // the geometry of neither is measured, so the page is still missing them and the
  // report still says so. This is the line that changes the day the layout sets one.
  it("names a fraction and a delimiter, whose shape is read and whose height is not", () => {
    expect(kindsOf(`<m:oMath>${FRACTION}</m:oMath>`)).toStrictEqual(["equation"]);
    const delimited = `<m:d><m:e>${mathRun("iota")}</m:e></m:d>`;
    expect(kindsOf(`<m:oMath>${delimited}</m:oMath>`)).toStrictEqual(["equation"]);
  });
});
