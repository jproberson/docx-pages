import { describe, expect, it } from "vitest";

import type { Substitution } from "@docx-pages/core";

import { facesPaintedWith } from "./docx-document.js";
import type { DocxFont } from "./docx-document.js";

// Which faces a file of the page has to carry, which is a different list from the
// one the caller handed in and is the only place the screen and the file could
// come apart.
//
// The bytes themselves are never read here, so a byte apiece is enough to tell one
// face from another.

const font = (name: string, byte: number, cut: Partial<DocxFont> = {}): DocxFont => ({
  name,
  bytes: new Uint8Array([byte]),
  ...cut,
});

const stoodIn = (asked: string, used: string): Substitution => ({
  requested: { name: asked, bold: false, italic: false },
  used: { name: used, bold: false, italic: false },
});

const names = (faces: readonly { readonly name: string }[]): readonly string[] =>
  faces.map((face) => face.name);

const byteOf = (faces: readonly { readonly bytes: Uint8Array }[], at: number): number | undefined =>
  faces[at]?.bytes[0];

describe("facesPaintedWith", () => {
  it("carries a face the caller supplied under the name it was supplied as", () => {
    const painted = facesPaintedWith([font("Calibri", 1)], [], []);

    expect(names(painted)).toStrictEqual(["Calibri"]);
    expect(byteOf(painted, 0)).toBe(1);
  });

  // The rule the whole function exists for. The layout measured `Cambria` and drew
  // it in Carlito's widths, and the file has to say `Cambria` for the same reason
  // the browser is offered it under that name.
  it("carries a face stood in for under the name the document asked for", () => {
    const painted = facesPaintedWith(
      [],
      [stoodIn("Cambria", "Carlito")],
      [font("Carlito", 7), font("Tinos", 8)],
    );

    expect(names(painted)).toStrictEqual(["Cambria", "Carlito", "Tinos"]);
    // Under the asked-for name, but drawn out of the bytes that stood in.
    expect(byteOf(painted, 0)).toBe(7);
  });

  it("keeps the weight and the slope the document asked for, not the stand-in's", () => {
    const painted = facesPaintedWith(
      [],
      [
        {
          requested: { name: "Cambria", bold: true, italic: false },
          used: { name: "Carlito", bold: false, italic: false },
        },
      ],
      [font("Carlito", 7)],
    );

    expect(painted[0]).toMatchObject({ name: "Cambria", bold: true, italic: false });
  });

  // Nothing can be carried for a stand-in whose bytes were never handed over. The
  // writer refuses such a document, which is louder than writing it in a face
  // nothing measured, and is not something to paper over here.
  it("carries nothing for a stand-in whose bytes nobody handed over", () => {
    const painted = facesPaintedWith([], [stoodIn("Cambria", "Carlito")], []);

    expect(painted).toStrictEqual([]);
  });

  it("passes over a substitution for a face the document never named", () => {
    const painted = facesPaintedWith([], [stoodIn("", "Carlito")], [font("Carlito", 7)]);

    expect(names(painted)).toStrictEqual(["Carlito"]);
  });

  // Both the writer and the browser take the first face that answers a name, so a
  // face supplied for exactness has to stand ahead of a default of the same name or
  // the page is drawn in the wrong one.
  it("puts a supplied face ahead of a default that shares its name", () => {
    const painted = facesPaintedWith([font("Calibri", 1)], [], [font("Calibri", 2)]);

    expect(names(painted)).toStrictEqual(["Calibri", "Calibri"]);
    expect(byteOf(painted, 0)).toBe(1);
  });

  it("takes a face to be upright and unslanted where the caller said neither", () => {
    const painted = facesPaintedWith([font("Calibri", 1)], [], []);

    expect(painted[0]).toMatchObject({ bold: false, italic: false });
  });

  it("carries the cut the caller stated", () => {
    const painted = facesPaintedWith([font("Calibri", 1, { bold: true, italic: true })], [], []);

    expect(painted[0]).toMatchObject({ bold: true, italic: true });
  });
});
