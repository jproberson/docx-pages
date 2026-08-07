import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { caseOf, gatherDocuments } from "./documents.js";

let root = "";
const at = (...parts: string[]): string => join(root, ...parts);

const put = (path: string, contents = "not a real document"): string => {
  mkdirSync(resolve(path, ".."), { recursive: true });
  writeFileSync(path, contents);
  return path;
};

beforeAll(() => {
  root = mkdtempSync(join(tmpdir(), "docx-pages-check-"));
});

afterAll(() => {
  rmSync(root, { recursive: true, force: true });
});

describe("gatherDocuments", () => {
  it("names a document for its own file, punctuation and case laid aside", () => {
    const into = at("named");
    const gathered = gatherDocuments(into, [put(at("from", "Quarterly Notes (Final).docx"))]);

    expect(gathered).toStrictEqual([
      { id: "quarterly-notes-final", path: resolve(into, "quarterly-notes-final.docx") },
    ]);
  });

  // The whole reason a gathered document is renamed at all: two directories may
  // each hold a `report.docx`, and the second drawing over the first would be a
  // page silently answering for the wrong document.
  it("keeps two documents of the same name apart", () => {
    const into = at("apart");
    const gathered = gatherDocuments(into, [
      put(at("one", "Report.docx"), "the first"),
      put(at("two", "report.docx"), "the second"),
    ]);

    expect(gathered.map((each) => each.id)).toStrictEqual(["report", "report-2"]);
    expect(new Set(gathered.map((each) => each.path)).size).toBe(2);
  });

  it("copies the document in rather than reading it where it lies", () => {
    const into = at("copied");
    const source = put(at("elsewhere", "kept.docx"), "the bytes");
    const [gathered] = gatherDocuments(into, [source]);

    expect(gathered?.path).toBe(resolve(into, "kept.docx"));
    expect(gathered?.path).not.toBe(resolve(source));
  });

  it("names a file it cannot find rather than drawing nothing and saying nothing", () => {
    expect(() => gatherDocuments(at("missing"), [at("nowhere", "absent.docx")])).toThrow(
      /no such document/,
    );
  });

  it("still names a document whose file name is punctuation alone", () => {
    const into = at("unnamed");
    const gathered = gatherDocuments(into, [put(at("odd", "---.docx"))]);

    expect(gathered.map((each) => each.id)).toStrictEqual(["document"]);
  });
});

describe("caseOf", () => {
  it("takes Word's pdf as the page to draw beside, where Word left one", () => {
    const path = put(at("withPdf", "asked.docx"));
    put(at("withPdf", "asked.pdf"));

    expect(caseOf("asked", path).renderedPath).toBe(resolve(at("withPdf", "asked.pdf")));
  });

  it("has no page to draw beside where Word left none", () => {
    const path = put(at("noPdf", "asked.docx"));

    expect(caseOf("asked", path).renderedPath).toBeNull();
  });

  // Nothing is measured for a document handed over on the command line, so every
  // number a reference case carries has to say so rather than reading as zero
  // agreement with Word.
  it("pins no measurement of its own", () => {
    const each = caseOf("asked", put(at("empty", "asked.docx")));

    expect(each.textLinesMatched).toBeNull();
    expect(each.bodyTopPt).toBeNull();
    expect(each.unhonoured).toStrictEqual([]);
  });
});
