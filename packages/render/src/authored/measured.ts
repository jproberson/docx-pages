import { existsSync, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

// What Word said about each authored document. These are Word's own answers,
// committed because the documents they describe were written here rather than
// found: nothing in either is anyone's collateral.
//
// Word rounds a paragraph's position to whole points and a shape's size to a
// twentieth of one, so pin against them no finer than that.
export const PARAGRAPH_TOLERANCE_PT = 0.5;
export const SHAPE_TOLERANCE_PT = 0.06;

// A paragraph inside a table is the one exception to reading these as the
// paragraph's own place: Word answers there for the row that holds it, and its
// horizontal answer, measured from the cell rather than from the text column, is
// nought whatever the cell holds. `answers.ts` says which is which.
export type MeasuredParagraph = {
  // Word numbers paragraphs from one; ours from zero.
  readonly index: number;
  readonly page: number;
  readonly topPt: number;
  // Where the paragraph's own content starts, measured from the left of the text
  // column rather than of the page.
  readonly leftPt: number;
};

// Where a character sits along its line, from the left of the text column. This is
// what says where a tab landed: the line's own start cannot.
export type MeasuredCharacter = {
  readonly paragraph: number;
  readonly index: number;
  readonly leftPt: number;
};

export type MeasuredShape = {
  readonly name: string;
  readonly widthPt: number;
  readonly heightPt: number;
};

export type MeasuredDocument = {
  readonly paragraphs: readonly MeasuredParagraph[];
  readonly characters: readonly MeasuredCharacter[];
  readonly shapes: readonly MeasuredShape[];
};

export type Measured = {
  readonly documents: Readonly<Record<string, MeasuredDocument>>;
};

// Found from this module rather than from the working directory. Read against the
// working directory the file is there only for a run started at the repository
// root, and a run started anywhere else used to lose every answer in it without
// saying so, which is a whole suite passing over nothing.
export const MEASURED_PATH = fileURLToPath(new URL("measured.json", import.meta.url));

// The file is written by the script beside it, so it is read for what it should
// hold rather than guarded against every shape it could: a field that is not a
// number is a measurement that never happened, and reading it as one would pin the
// layout against nothing.
const fields = (value: unknown): Record<string, unknown> =>
  typeof value === "object" && value !== null ? { ...value } : {};

const number = (value: unknown, where: string): number => {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${MEASURED_PATH}: expected a number at ${where}`);
  }
  return value;
};

const list = (value: unknown): readonly unknown[] => (Array.isArray(value) ? value : []);

const readParagraph = (value: unknown, where: string): MeasuredParagraph => {
  const each = fields(value);
  return {
    index: number(each["index"], `${where}.index`),
    page: number(each["page"], `${where}.page`),
    topPt: number(each["topPt"], `${where}.topPt`),
    leftPt: number(each["leftPt"], `${where}.leftPt`),
  };
};

const readCharacter = (value: unknown, where: string): MeasuredCharacter => {
  const each = fields(value);
  return {
    paragraph: number(each["paragraph"], `${where}.paragraph`),
    index: number(each["index"], `${where}.index`),
    leftPt: number(each["leftPt"], `${where}.leftPt`),
  };
};

const readShape = (value: unknown, where: string): MeasuredShape => {
  const each = fields(value);
  const name = each["name"];
  return {
    name: typeof name === "string" ? name : "",
    widthPt: number(each["widthPt"], `${where}.widthPt`),
    heightPt: number(each["heightPt"], `${where}.heightPt`),
  };
};

function readDocument(value: unknown, where: string): MeasuredDocument {
  const each = fields(value);
  return {
    paragraphs: list(each["paragraphs"]).map((one, at) =>
      readParagraph(one, `${where}.paragraphs[${String(at)}]`),
    ),
    characters: list(each["characters"]).map((one, at) =>
      readCharacter(one, `${where}.characters[${String(at)}]`),
    ),
    shapes: list(each["shapes"]).map((one, at) => readShape(one, `${where}.shapes[${String(at)}]`)),
  };
}

// **This file is committed**, unlike the reference manifest, so it is never
// legitimately missing: the suite goes quiet over a machine without Word's Calibri
// and over nothing else. Answering an empty measurement for a file that is not
// there is how a green run comes to mean almost nothing ran.
export function readMeasured(path: string = MEASURED_PATH): Measured {
  if (!existsSync(path)) throw new Error(`${path}: no measurements to read`);

  const parsed: unknown = JSON.parse(readFileSync(path, "utf8"));
  const documents: Record<string, MeasuredDocument> = {};
  for (const [id, each] of Object.entries(fields(fields(parsed)["documents"]))) {
    documents[id] = readDocument(each, `documents.${id}`);
  }
  return { documents };
}
