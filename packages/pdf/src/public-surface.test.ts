import { describe, expect, it } from "vitest";

import * as pdf from "./index.js";

// What `@docx-pages/pdf` promises anyone who installs it, pinned by name, for the
// same reason core's list is: dropping an export typechecks perfectly inside this
// repository and breaks every consumer outside it.
//
// Two names, and it should stay about two. Everything this package knows how to do
// is decided by the layout it is handed; there is no second way to draw a page and
// no knob here that moves anything on one.
const PDF_SURFACE: readonly string[] = ["pdfOfDocx", "writePdf"];

describe("the public surface of @docx-pages/pdf", () => {
  it("exports what it promises, and nothing has quietly gone away", () => {
    const found = Object.keys(pdf)
      .filter((name) => name !== "default")
      .sort();

    expect(found).toStrictEqual([...PDF_SURFACE].sort());
  });
});
