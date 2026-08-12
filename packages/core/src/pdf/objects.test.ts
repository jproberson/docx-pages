import { strFromU8, unzlibSync } from "fflate";
import { describe, expect, it } from "vitest";

import { type DocxPagesError, isDocxPagesError } from "../errors.js";

import {
  formatPdfNumber,
  pdfArray,
  pdfDictionary,
  pdfHexString,
  pdfName,
  pdfNumber,
  pdfObjects,
  pdfStream,
  pdfString,
  type PdfValue,
} from "./objects.js";

const caught = (run: () => unknown): DocxPagesError => {
  try {
    run();
  } catch (error: unknown) {
    if (isDocxPagesError(error)) return error;
    throw error;
  }
  throw new Error("expected a DocxPagesError");
};

// The syntax is latin1, so the file reads back one character to a byte.
const textOf = (bytes: Uint8Array): string => strFromU8(bytes, true);

// One object holding the value under test, read back out of the whole file.
function written(value: PdfValue): string {
  const objects = pdfObjects();
  const root = objects.add(value);
  const text = textOf(objects.bytes({ root }));
  const body = /1 0 obj\n([\s\S]*)\nendobj/.exec(text)?.[1];
  if (body === undefined) throw new Error(`no object found in ${text}`);
  return body;
}

describe("a pdf number", () => {
  it("writes a whole number without a point", () => {
    expect(formatPdfNumber(72)).toBe("72");
    expect(formatPdfNumber(-72)).toBe("-72");
    expect(formatPdfNumber(0)).toBe("0");
  });

  it("writes a fraction to as many places as it has, and no trailing zeros", () => {
    expect(formatPdfNumber(11.5)).toBe("11.5");
    expect(formatPdfNumber(1 / 3)).toBe("0.333333");
  });

  // A pdf real carries no exponent, and `String(1e-7)` is one. A reader meeting
  // it stops on a syntax error rather than drawing the page nearly right.
  it("writes a number near zero out in full rather than as an exponent", () => {
    expect(formatPdfNumber(1e-7)).toBe("0");
    expect(formatPdfNumber(0.0000015)).toBe("0.000002");
  });

  it("never writes a negative zero, which is a sign nothing here means", () => {
    expect(formatPdfNumber(-1e-9)).toBe("0");
    expect(formatPdfNumber(-0)).toBe("0");
  });

  it("refuses a number a pdf has no way to hold", () => {
    expect(caught(() => formatPdfNumber(Number.NaN)).code).toBe("pdf-number-unwritable");
    expect(caught(() => formatPdfNumber(Number.POSITIVE_INFINITY)).code).toBe(
      "pdf-number-unwritable",
    );
  });

  // Past a pdf's own limit on a whole number JavaScript writes an exponent, which
  // is a syntax error to a reader. Refused rather than written as one: no page
  // reaches out here, so a number that does is a fault to be told about.
  it("refuses a number past what a pdf can hold rather than writing an exponent", () => {
    expect(caught(() => formatPdfNumber(1e21)).code).toBe("pdf-number-unwritable");
    expect(formatPdfNumber(2147483647)).toBe("2147483647");
  });
});

describe("a pdf name", () => {
  it("writes an ordinary name as itself", () => {
    expect(written(pdfName("FontFile2"))).toBe("/FontFile2");
  });

  // A face is named by the document, which may call it anything at all.
  it("escapes what a name may not carry, so a face's own name survives it", () => {
    expect(written(pdfName("Times New Roman"))).toBe("/Times#20New#20Roman");
    expect(written(pdfName("a/b#c"))).toBe("/a#2fb#23c");
  });
});

describe("a pdf string", () => {
  it("escapes the brackets that would end it early", () => {
    expect(written(pdfString("a (b) c"))).toBe("(a \\(b\\) c)");
  });

  it("escapes the backslash that does the escaping", () => {
    expect(written(pdfString("a\\b"))).toBe("(a\\\\b)");
  });

  it("writes bytes as hex, which is what a glyph number needs", () => {
    expect(written(pdfHexString(Uint8Array.from([0, 36, 255])))).toBe("<0024ff>");
  });
});

describe("a pdf dictionary", () => {
  it("writes its entries in the order they were stated", () => {
    const value = pdfDictionary({
      Type: pdfName("Page"),
      MediaBox: pdfArray([0, 0, 612, 792].map(pdfNumber)),
    });

    expect(written(value)).toBe("<</Type /Page/MediaBox [0 0 612 792]>>");
  });

  // Which is how a dictionary carries the parts of itself a document did not ask
  // for: an absent outline is no outline rather than a null one.
  it("leaves out an entry stating no value", () => {
    expect(written(pdfDictionary({ Type: pdfName("Page"), Rotate: undefined }))).toBe(
      "<</Type /Page>>",
    );
  });
});

describe("a pdf stream", () => {
  it("deflates its bytes and says so, and states the length it wrote", () => {
    const objects = pdfObjects();
    objects.add(pdfStream({ Type: pdfName("XObject") }, new Uint8Array(200)));
    const text = textOf(objects.bytes({ root: { kind: "reference", number: 1 } }));

    expect(text).toContain("/Filter /FlateDecode");
    const length = /\/Length (\d+)/.exec(text)?.[1];
    expect(Number(length)).toBeLessThan(200);
  });

  it("round-trips the bytes it was given", () => {
    const content = Uint8Array.from({ length: 300 }, (_, at) => at % 256);
    const objects = pdfObjects();
    objects.add(pdfStream({}, content));
    const bytes = objects.bytes({ root: { kind: "reference", number: 1 } });

    const text = textOf(bytes);
    const from = text.indexOf("stream\n") + "stream\n".length;
    const to = text.indexOf("\nendstream");
    expect([...unzlibSync(bytes.subarray(from, to))]).toStrictEqual([...content]);
  });

  // A jpeg is already compressed and states its own filter; deflating it again
  // would grow the file and hide the filter the reader has to see.
  it("passes bytes through untouched where it is told not to deflate", () => {
    const objects = pdfObjects();
    objects.add(pdfStream({ Filter: pdfName("DCTDecode") }, Uint8Array.from([1, 2, 3]), false));
    const text = textOf(objects.bytes({ root: { kind: "reference", number: 1 } }));

    expect(text).toContain("/Filter /DCTDecode");
    expect(text).not.toContain("FlateDecode");
    expect(text).toContain("/Length 3");
  });
});

describe("a written document", () => {
  it("opens with the header and a binary comment, so nothing takes it for text", () => {
    const objects = pdfObjects();
    const bytes = objects.bytes({ root: objects.add(pdfDictionary({})) });

    expect(textOf(bytes.subarray(0, 9))).toBe("%PDF-1.7\n");
    expect([...bytes.subarray(9, 15)]).toStrictEqual([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]);
  });

  it("ends with the trailer, the table's own offset, and the marker", () => {
    const objects = pdfObjects();
    const root = objects.add(pdfDictionary({ Type: pdfName("Catalog") }));
    const info = objects.add(pdfDictionary({ Producer: pdfString("docx-pages") }));
    const text = textOf(objects.bytes({ root, info }));

    expect(text).toContain("/Size 3/Root 1 0 R/Info 2 0 R");
    expect(text.endsWith("%%EOF\n")).toBe(true);

    const startedAt = Number(/startxref\n(\d+)/.exec(text)?.[1]);
    expect(text.slice(startedAt, startedAt + 5)).toBe("xref\n");
  });

  // A reader seeks to an object by the offset in this table, so an entry that is
  // out by a byte is a document that will not open.
  it("points every entry in the table at the object it names", () => {
    const objects = pdfObjects();
    const first = objects.add(pdfDictionary({ Type: pdfName("Catalog") }));
    objects.add(pdfStream({}, new Uint8Array(64)));
    objects.add(pdfString("last"));
    const text = textOf(objects.bytes({ root: first }));

    const table = /xref\n0 4\n([\s\S]{80})/.exec(text)?.[1] ?? "";
    expect(table.slice(0, 20)).toBe("0000000000 65535 f \n");

    for (const number of [1, 2, 3]) {
      const offset = Number(table.slice(number * 20, number * 20 + 10));
      expect(text.slice(offset, offset + `${String(number)} 0 obj`.length)).toBe(
        `${String(number)} 0 obj`,
      );
    }
  });

  // The circular half of a pdf: a page names the tree it hangs off and the tree
  // names the page, so one of the two is referred to before it is written.
  it("lets an object be referred to before it is filled in", () => {
    const objects = pdfObjects();
    const tree = objects.reserve();
    const page = objects.add(pdfDictionary({ Type: pdfName("Page"), Parent: tree }));
    objects.put(tree, pdfDictionary({ Type: pdfName("Pages"), Kids: pdfArray([page]) }));

    const text = textOf(objects.bytes({ root: tree }));
    expect(text).toContain("/Type /Page/Parent 1 0 R");
    expect(text).toContain("/Type /Pages/Kids [2 0 R]");
  });

  it("refuses to write a document holding an object nobody filled in", () => {
    const objects = pdfObjects();
    const root = objects.add(pdfDictionary({}));
    objects.reserve();

    const error = caught(() => objects.bytes({ root }));
    expect(error.code).toBe("pdf-object-unwritten");
    expect(error.context["object"]).toBe(2);
  });

  it("refuses to fill the same object in twice", () => {
    const objects = pdfObjects();
    const at = objects.reserve();
    objects.put(at, pdfDictionary({}));

    expect(
      caught(() => {
        objects.put(at, pdfDictionary({}));
      }).code,
    ).toBe("pdf-object-written-twice");
  });
});
