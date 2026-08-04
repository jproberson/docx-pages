import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

import { isOnePagerError } from "@onepager/core";

import { readReferenceManifest } from "./cases.js";

const directory = mkdtempSync(resolve(tmpdir(), "onepager-cases-"));
let written = 0;

const manifestOf = (json: string): string => {
  const path = resolve(directory, `manifest-${String(written++)}.json`);
  writeFileSync(path, json);
  return path;
};

const codeOf = (run: () => unknown): string => {
  try {
    run();
  } catch (error: unknown) {
    if (isOnePagerError(error)) return error.code;
    throw error;
  }
  throw new Error("expected a failure");
};

const ONE_CASE = `{"cases":[{"id":"a","documentPath":"/tmp/a.docx"}]}`;

describe("readReferenceManifest", () => {
  it("reports nothing to run when no manifest exists", () => {
    expect(readReferenceManifest(resolve(directory, "absent.json"))).toStrictEqual({
      fonts: [],
      cases: [],
    });
  });

  it("fills every optional expectation in so a sparse case still describes itself", () => {
    const [only] = readReferenceManifest(manifestOf(ONE_CASE)).cases;
    expect(only).toStrictEqual({
      id: "a",
      documentPath: "/tmp/a.docx",
      renderedPath: null,
      tolerancePt: 0.5,
      bodyTopPt: null,
      headerTopsPt: [],
      bodyTopsPt: [],
      headerFloatCount: null,
      leastBodyFloatCount: null,
      floatsPt: [],
      inlinesPt: [],
      disjointFloatPairs: [],
      renderedImagesPt: [],
      renderedPageIndexes: [],
    });
  });

  it("reads the expectations a case does declare", () => {
    const json = `{"cases":[{"id":"a","documentPath":"/tmp/a.docx","renderedPath":"/tmp/a.pdf",
      "tolerancePt":0.1,"bodyTopPt":86.75,"bodyTopsPt":[{"index":3,"topPt":120.5}],
      "floatsPt":[{"index":0,"leftPt":25.5,"topPt":21.75}],"disjointFloatPairs":[[1,2]],
      "renderedPageIndexes":[0,1]}]}`;
    const [only] = readReferenceManifest(manifestOf(json)).cases;
    expect(only?.tolerancePt).toBe(0.1);
    expect(only?.bodyTopsPt).toStrictEqual([{ index: 3, topPt: 120.5 }]);
    expect(only?.disjointFloatPairs).toStrictEqual([[1, 2]]);
    expect(only?.renderedPageIndexes).toStrictEqual([0, 1]);
  });

  it("reads a font's metrics alongside where its file lives", () => {
    const json = `{"fonts":[{"name":"Meridian Sans","filePath":"/tmp/m.woff","fileFormat":"woff",
      "metrics":{"unitsPerEm":1000,"ascender":800,"descender":-200,"lineGap":0}},
      {"name":"Meridian Sans","bold":true,"filePath":"/tmp/m-bold.woff","fileFormat":"woff",
      "metrics":{"unitsPerEm":1000,"ascender":800,"descender":-200,"lineGap":0}}]}`;
    const metrics = { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 };

    expect(readReferenceManifest(manifestOf(json)).fonts).toStrictEqual([
      {
        name: "Meridian Sans",
        bold: false,
        italic: false,
        filePath: "/tmp/m.woff",
        fileFormat: "woff",
        metrics,
      },
      {
        name: "Meridian Sans",
        bold: true,
        italic: false,
        filePath: "/tmp/m-bold.woff",
        fileFormat: "woff",
        metrics,
      },
    ]);
  });

  it("refuses a manifest that is not json", () => {
    expect(codeOf(() => readReferenceManifest(manifestOf(`{`)))).toBe(
      "reference-manifest-unreadable",
    );
  });

  it("refuses a case that names no document", () => {
    expect(codeOf(() => readReferenceManifest(manifestOf(`{"cases":[{"id":"a"}]}`)))).toBe(
      "reference-manifest-invalid",
    );
  });

  it("says which manifest, and where inside it, a bad expectation sits", () => {
    const json = `{"cases":[{"id":"a","documentPath":"/tmp/a.docx","bodyTopsPt":[{"index":1}]}]}`;
    const path = manifestOf(json);
    try {
      readReferenceManifest(path);
    } catch (error: unknown) {
      if (!isOnePagerError(error)) throw error;
      expect(error.context["where"]).toBe(`${path}#cases[0].bodyTopsPt[0].topPt`);
      return;
    }
    throw new Error("expected a failure");
  });
});
