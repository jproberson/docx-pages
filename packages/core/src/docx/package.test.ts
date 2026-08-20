import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { isDocxPagesError, type DocxPagesError } from "../errors.js";
import { MAIN_DOCUMENT_PART, openDocx, partXml } from "./package.js";

const packageHolding = (documentXml: string) =>
  openDocx(buildDocx({ "word/document.xml": documentXml }));

const thrownBy = (read: () => unknown): DocxPagesError => {
  try {
    read();
  } catch (error: unknown) {
    if (isDocxPagesError(error)) return error;
    throw error;
  }
  throw new Error("expected a DocxPagesError");
};

describe("partXml", () => {
  it("reads the root of a part that is whole", () => {
    expect(partXml(packageHolding(wordDocument("<w:p/>")), MAIN_DOCUMENT_PART).name).toBe(
      "document",
    );
  });

  // A part cut off in the middle of a tag is what the parser throws over rather than
  // answers nothing for, and a caller telling the documents it refuses from faults of
  // its own has only the error to tell them apart by.
  it("refuses a part cut off in the middle of a tag rather than letting the parser throw", () => {
    const thrown = thrownBy(() => partXml(packageHolding("<w:document"), MAIN_DOCUMENT_PART));
    expect(thrown.code).toBe("docx-malformed");
    expect(thrown.context["part"]).toBe(MAIN_DOCUMENT_PART);
    expect(thrown.cause).toBeInstanceOf(Error);
  });

  it("refuses a part holding no element at all", () => {
    const thrown = thrownBy(() => partXml(packageHolding(""), MAIN_DOCUMENT_PART));
    expect(thrown.code).toBe("docx-malformed");
    expect(thrown.context["part"]).toBe(MAIN_DOCUMENT_PART);
  });

  it("names the part a package does not carry", () => {
    const thrown = thrownBy(() =>
      partXml(packageHolding(wordDocument("<w:p/>")), "word/styles.xml"),
    );
    expect(thrown.code).toBe("docx-malformed");
    expect(thrown.at).toBe("core/docx/package.partText");
  });
});
