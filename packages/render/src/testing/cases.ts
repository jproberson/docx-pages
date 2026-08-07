import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";

import {
  NO_ADVANCES,
  DocxPagesError,
  readFontFile,
  type FaceRequest,
  type FontMetrics,
  type SuppliedFace,
} from "@docx-pages/core";

// The reference documents and the geometry Word produced from them stay outside
// the repo. Point this at a manifest describing them; without one the reference
// suites report nothing to run.
const MANIFEST_PATH = resolve(
  process.env["DOCX_PAGES_REFERENCE_MANIFEST"] ?? "samples/reference-cases.json",
);

const AT = "render/testing/cases.readManifest";

export type ReferenceFont = {
  readonly name: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly filePath: string | null;
  readonly fileFormat: string | null;
  readonly metrics: FontMetrics;
};

export type ParagraphTop = {
  readonly index: number;
  readonly topPt: number;
};

export type FloatOrigin = {
  readonly index: number;
  readonly leftPt: number;
  readonly topPt: number;
};

// Either coordinate may be left out when Word's value for it is not yet explained.
export type InlineOrigin = {
  readonly index: number;
  readonly leftPt: number | null;
  readonly topPt: number | null;
};

export type PointRect = {
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

export type ReferenceCase = {
  readonly id: string;
  readonly documentPath: string;
  readonly renderedPath: string | null;
  readonly tolerancePt: number;
  readonly bodyTopPt: number | null;
  readonly headerTopsPt: readonly ParagraphTop[];
  readonly bodyTopsPt: readonly ParagraphTop[];
  readonly headerFloatCount: number | null;
  readonly leastBodyFloatCount: number | null;
  readonly floatsPt: readonly FloatOrigin[];
  readonly inlinesPt: readonly InlineOrigin[];
  readonly disjointFloatPairs: readonly (readonly [number, number])[];
  // Left out, nobody measured Word's own output for images and there is nothing to
  // compare against. Written out empty, Word drew none, which is as much an answer
  // as any other: it says this document must draw none either.
  readonly renderedImagesPt: readonly PointRect[] | null;
  readonly renderedPageIndexes: readonly number[] | null;
  // How many pictures this project places that Word's own output draws nothing
  // for, and how many Word drew that nothing here stands on. Left out, both are
  // none: a picture drawn out of nowhere, or one silently dropped, cannot pass
  // unnoticed. Where a picture is drawn in the wrong place there is no count at
  // all, since the two are paired and that is never allowed.
  readonly picturesWordDidNotDraw: number;
  readonly picturesWeDidNotDraw: number;
  // How many laid-out text lines are expected to land where Word drew the same
  // line, within textTolerancePt. Neither the text nor its position is recorded
  // here; both are read from the document and Word's own output at run time.
  readonly textLinesMatched: number | null;
  readonly textLinesPlaced: number | null;
  // How many of the runs a line is made of are expected to be found in Word's own
  // output and to start where Word started them. A line matches by the text of
  // the whole of it; a run is pinned one face at a time inside that.
  readonly textRunsMatched: number | null;
  readonly textRunsPlaced: number | null;
  // How many times a picture in a format no browser draws is met over the whole
  // document, EMF and WMF being what Word writes when a chart is pasted in. One in
  // a header is met once per page. They are marked rather than drawn, and counted
  // here so a new one cannot slip in unnoticed.
  readonly unrenderablePictures: number;
  // How many drawings the document holds of a kind this project does not draw at
  // all, a chart being the one met so far. They are marked as unknown rather than
  // drawn, and counted here so a new one cannot slip in unnoticed.
  readonly unknownDrawings: number;
  // How many blocks of colour and runs of text the metafiles in the document are
  // expected to draw, every one of them where Word drew the same shape from the
  // same recording. Left out, the document has no metafile worth playing.
  readonly metafileFills: number | null;
  readonly metafileRuns: number | null;
  // How many list numbers are expected to be identified against Word's own output
  // and to sit where Word drew them.
  readonly numbersMatched: number | null;
  readonly numbersPlaced: number | null;
  readonly textTolerancePt: number;
  // Everything in the document this project is expected to pass over, by the name
  // the fidelity report gives it, in the order it reports them. Left out, the
  // document is expected to report nothing at all, which is the whole point of the
  // list: a gap cannot be introduced quietly.
  readonly unhonoured: readonly string[];
};

// Font files the reader has to cope with, whether or not any reference document
// is laid out in them.
export type ReferenceFontFile = {
  readonly filePath: string;
  readonly fileFormat: string | null;
  readonly metrics: FontMetrics;
};

// The face a document was authored in, as against the one Word had to fall back
// on when it rendered the reference pdf. Everything about it is read from the file
// itself, since no substitute's metrics stand in for it.
export type AuthoredFace = {
  readonly name: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly filePath: string;
};

export type ReferenceManifest = {
  readonly fonts: readonly ReferenceFont[];
  readonly authoredFonts: readonly AuthoredFace[];
  readonly fontFiles: readonly ReferenceFontFile[];
  readonly cases: readonly ReferenceCase[];
};

const EMPTY: ReferenceManifest = { fonts: [], authoredFonts: [], fontFiles: [], cases: [] };

// `where` carries the manifest path as its root, so it locates the fault on its own.
const invalid = (message: string, where: string): DocxPagesError =>
  new DocxPagesError({ code: "reference-manifest-invalid", message, at: AT, context: { where } });

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

function record(value: unknown, where: string): Record<string, unknown> {
  if (!isRecord(value)) throw invalid("expected an object", where);
  return value;
}

function entries(source: Record<string, unknown>, key: string, where: string): readonly unknown[] {
  const value = source[key];
  if (value === undefined) return [];
  if (!Array.isArray(value)) throw invalid("expected an array", `${where}.${key}`);
  return value;
}

function number(source: Record<string, unknown>, key: string, where: string): number {
  const value = source[key];
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw invalid("expected a finite number", `${where}.${key}`);
  }
  return value;
}

function optionalNumber(
  source: Record<string, unknown>,
  key: string,
  where: string,
): number | null {
  return source[key] === undefined ? null : number(source, key, where);
}

function text(source: Record<string, unknown>, key: string, where: string): string {
  const value = source[key];
  if (typeof value !== "string" || value === "") {
    throw invalid("expected a non-empty string", `${where}.${key}`);
  }
  return value;
}

function optionalText(source: Record<string, unknown>, key: string, where: string): string | null {
  return source[key] === undefined ? null : text(source, key, where);
}

function flag(source: Record<string, unknown>, key: string, where: string): boolean {
  const value = source[key];
  if (value === undefined) return false;
  if (typeof value !== "boolean") throw invalid("expected a boolean", `${where}.${key}`);
  return value;
}

const readMetrics = (value: unknown, where: string): FontMetrics => {
  const source = record(value, where);
  return {
    unitsPerEm: number(source, "unitsPerEm", where),
    ascender: number(source, "ascender", where),
    descender: number(source, "descender", where),
    lineGap: number(source, "lineGap", where),
  };
};

const readFont = (value: unknown, where: string): ReferenceFont => {
  const source = record(value, where);
  return {
    name: text(source, "name", where),
    bold: flag(source, "bold", where),
    italic: flag(source, "italic", where),
    filePath: optionalText(source, "filePath", where),
    fileFormat: optionalText(source, "fileFormat", where),
    metrics: readMetrics(source["metrics"], `${where}.metrics`),
  };
};

const readAuthoredFace = (value: unknown, where: string): AuthoredFace => {
  const source = record(value, where);
  return {
    name: text(source, "name", where),
    bold: flag(source, "bold", where),
    italic: flag(source, "italic", where),
    filePath: text(source, "filePath", where),
  };
};

const readFontFileEntry = (value: unknown, where: string): ReferenceFontFile => {
  const source = record(value, where);
  return {
    filePath: text(source, "filePath", where),
    fileFormat: optionalText(source, "fileFormat", where),
    metrics: readMetrics(source["metrics"], `${where}.metrics`),
  };
};

const readParagraphTop = (value: unknown, where: string): ParagraphTop => {
  const source = record(value, where);
  return { index: number(source, "index", where), topPt: number(source, "topPt", where) };
};

const readFloatOrigin = (value: unknown, where: string): FloatOrigin => {
  const source = record(value, where);
  return {
    index: number(source, "index", where),
    leftPt: number(source, "leftPt", where),
    topPt: number(source, "topPt", where),
  };
};

const readInlineOrigin = (value: unknown, where: string): InlineOrigin => {
  const source = record(value, where);
  return {
    index: number(source, "index", where),
    leftPt: optionalNumber(source, "leftPt", where),
    topPt: optionalNumber(source, "topPt", where),
  };
};

const readRect = (value: unknown, where: string): PointRect => {
  const source = record(value, where);
  return {
    leftPt: number(source, "leftPt", where),
    topPt: number(source, "topPt", where),
    widthPt: number(source, "widthPt", where),
    heightPt: number(source, "heightPt", where),
  };
};

function readPair(value: unknown, where: string): readonly [number, number] {
  const entries: readonly unknown[] = Array.isArray(value) ? value : [];
  const [first, second] = entries;
  if (entries.length !== 2 || typeof first !== "number" || typeof second !== "number") {
    throw invalid("expected a pair of float indexes", where);
  }
  return [first, second];
}

function readIndex(value: unknown, where: string): number {
  if (typeof value !== "number" || !Number.isInteger(value)) {
    throw invalid("expected an integer", where);
  }
  return value;
}

function readCase(value: unknown, at: number, root: string): ReferenceCase {
  const where = `${root}#cases[${String(at)}]`;
  const source = record(value, where);
  const list = <T>(key: string, read: (entry: unknown, where: string) => T): readonly T[] =>
    entries(source, key, where).map((entry, index) =>
      read(entry, `${where}.${key}[${String(index)}]`),
    );

  // A list nobody wrote is a measurement nobody took, which nothing can be asserted
  // against; one written out empty is a measurement that found nothing. Collapsing
  // the two would let a document quietly stop being checked at all.
  const measuredList = <T>(
    key: string,
    read: (entry: unknown, where: string) => T,
  ): readonly T[] | null => (source[key] === undefined ? null : list(key, read));

  return {
    id: text(source, "id", where),
    documentPath: text(source, "documentPath", where),
    renderedPath: optionalText(source, "renderedPath", where),
    tolerancePt: optionalNumber(source, "tolerancePt", where) ?? 0.5,
    bodyTopPt: optionalNumber(source, "bodyTopPt", where),
    headerTopsPt: list("headerTopsPt", readParagraphTop),
    bodyTopsPt: list("bodyTopsPt", readParagraphTop),
    headerFloatCount: optionalNumber(source, "headerFloatCount", where),
    leastBodyFloatCount: optionalNumber(source, "leastBodyFloatCount", where),
    floatsPt: list("floatsPt", readFloatOrigin),
    inlinesPt: list("inlinesPt", readInlineOrigin),
    disjointFloatPairs: list("disjointFloatPairs", readPair),
    renderedImagesPt: measuredList("renderedImagesPt", readRect),
    renderedPageIndexes: measuredList("renderedPageIndexes", readIndex),
    picturesWordDidNotDraw: optionalNumber(source, "picturesWordDidNotDraw", where) ?? 0,
    picturesWeDidNotDraw: optionalNumber(source, "picturesWeDidNotDraw", where) ?? 0,
    textLinesMatched: optionalNumber(source, "textLinesMatched", where),
    textLinesPlaced: optionalNumber(source, "textLinesPlaced", where),
    textRunsMatched: optionalNumber(source, "textRunsMatched", where),
    textRunsPlaced: optionalNumber(source, "textRunsPlaced", where),
    unrenderablePictures: optionalNumber(source, "unrenderablePictures", where) ?? 0,
    unknownDrawings: optionalNumber(source, "unknownDrawings", where) ?? 0,
    metafileFills: optionalNumber(source, "metafileFills", where),
    metafileRuns: optionalNumber(source, "metafileRuns", where),
    numbersMatched: optionalNumber(source, "numbersMatched", where),
    numbersPlaced: optionalNumber(source, "numbersPlaced", where),
    textTolerancePt: optionalNumber(source, "textTolerancePt", where) ?? 1,
    unhonoured: entries(source, "unhonoured", where).map((each, at) => {
      if (typeof each !== "string")
        throw invalid("expected a string", `${where}.unhonoured[${String(at)}]`);
      return each;
    }),
  };
}

export function readReferenceManifest(path: string = MANIFEST_PATH): ReferenceManifest {
  if (!existsSync(path)) return EMPTY;

  let parsed: unknown;
  try {
    parsed = JSON.parse(readFileSync(path, "utf8"));
  } catch (error: unknown) {
    throw new DocxPagesError({
      code: "reference-manifest-unreadable",
      message: "the reference manifest is not readable json",
      at: AT,
      context: { path },
      cause: error,
    });
  }

  const source = record(parsed, path);
  return {
    fonts: entries(source, "fonts", path).map((entry, at) =>
      readFont(entry, `${path}#fonts[${String(at)}]`),
    ),
    authoredFonts: entries(source, "authoredFonts", path).map((entry, at) =>
      readAuthoredFace(entry, `${path}#authoredFonts[${String(at)}]`),
    ),
    fontFiles: entries(source, "fontFiles", path).map((entry, at) =>
      readFontFileEntry(entry, `${path}#fontFiles[${String(at)}]`),
    ),
    cases: entries(source, "cases", path).map((entry, at) => readCase(entry, at, path)),
  };
}

export const referenceFonts = (): readonly ReferenceFont[] => readReferenceManifest().fonts;

export const referenceFontFiles = (): readonly ReferenceFontFile[] =>
  readReferenceManifest().fontFiles;

export const referenceCases = (): readonly ReferenceCase[] =>
  readReferenceManifest().cases.filter((each) => existsSync(each.documentPath));

// The manifest's metrics stay authoritative for vertical geometry; the font file
// is read for the advances and for what kind of face it is, neither of which a
// manifest could carry.
function faceOf(font: ReferenceFont): SuppliedFace {
  const path = font.filePath;
  const style = { name: font.name, bold: font.bold, italic: font.italic };
  if (path === null || !existsSync(path)) {
    return { ...style, metrics: font.metrics, advances: NO_ADVANCES };
  }
  const read = readFontFile(new Uint8Array(readFileSync(path)));
  return {
    ...style,
    metrics: font.metrics,
    advances: read.advances,
    sansSerif: read.sansSerif,
  };
}

export const suppliedFaces = (): readonly SuppliedFace[] => referenceFonts().map(faceOf);

export const authoredFonts = (): readonly AuthoredFace[] =>
  readReferenceManifest().authoredFonts.filter((each) => existsSync(each.filePath));

// The authored file is the whole answer for the face it names: its own metrics as
// well as its own advances, since the manifest's are the substitute's.
function authoredFaceOf(face: AuthoredFace): SuppliedFace {
  const read = readFontFile(new Uint8Array(readFileSync(face.filePath)));
  return {
    name: face.name,
    bold: face.bold,
    italic: face.italic,
    metrics: read.metrics,
    advances: read.advances,
    sansSerif: read.sansSerif,
  };
}

const styleKey = (face: FaceRequest): string =>
  `${face.name.toLowerCase()}|${String(face.bold)}|${String(face.italic)}`;

// Every face the reference documents need, with the authored ones in place of the
// substitutes Word fell back on.
export function authoredFaces(): readonly SuppliedFace[] {
  const authored = authoredFonts().map(authoredFaceOf);
  const taken = new Set(authored.map(styleKey));
  return [...authored, ...suppliedFaces().filter((each) => !taken.has(styleKey(each)))];
}
