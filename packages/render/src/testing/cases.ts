import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";

import {
  NO_ADVANCES,
  OnePagerError,
  readFontFile,
  type FontMetrics,
  type SuppliedFace,
} from "@onepager/core";

// The reference documents and the geometry Word produced from them stay outside
// the repo. Point this at a manifest describing them; without one the reference
// suites report nothing to run.
const MANIFEST_PATH = resolve(
  process.env["ONEPAGER_REFERENCE_MANIFEST"] ?? "samples/reference-cases.json",
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
  readonly renderedImagesPt: readonly PointRect[];
  readonly renderedPageIndexes: readonly number[];
};

export type ReferenceManifest = {
  readonly fonts: readonly ReferenceFont[];
  readonly cases: readonly ReferenceCase[];
};

const EMPTY: ReferenceManifest = { fonts: [], cases: [] };

// `where` carries the manifest path as its root, so it locates the fault on its own.
const invalid = (message: string, where: string): OnePagerError =>
  new OnePagerError({ code: "reference-manifest-invalid", message, at: AT, context: { where } });

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
    renderedImagesPt: list("renderedImagesPt", readRect),
    renderedPageIndexes: list("renderedPageIndexes", readIndex),
  };
}

export function readReferenceManifest(path: string = MANIFEST_PATH): ReferenceManifest {
  if (!existsSync(path)) return EMPTY;

  let parsed: unknown;
  try {
    parsed = JSON.parse(readFileSync(path, "utf8"));
  } catch (error: unknown) {
    throw new OnePagerError({
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
    cases: entries(source, "cases", path).map((entry, at) => readCase(entry, at, path)),
  };
}

export const referenceFonts = (): readonly ReferenceFont[] => readReferenceManifest().fonts;

export const referenceCases = (): readonly ReferenceCase[] =>
  readReferenceManifest().cases.filter((each) => existsSync(each.documentPath));

// The manifest's metrics stay authoritative for vertical geometry; the font file
// is read only for the advances, which no manifest could carry.
function faceOf(font: ReferenceFont): SuppliedFace {
  const path = font.filePath;
  const style = { name: font.name, bold: font.bold, italic: font.italic };
  if (path === null || !existsSync(path)) {
    return { ...style, metrics: font.metrics, advances: NO_ADVANCES };
  }
  return {
    ...style,
    metrics: font.metrics,
    advances: readFontFile(new Uint8Array(readFileSync(path))).advances,
  };
}

export const suppliedFaces = (): readonly SuppliedFace[] => referenceFonts().map(faceOf);
