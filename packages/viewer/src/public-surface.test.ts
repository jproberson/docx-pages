import { describe, expect, it } from "vitest";

import * as viewer from "./index.js";
import * as pack from "./pack.js";

// What `@docx-pages/viewer` promises anyone who installs it. The same argument as
// core's list, and a much shorter one: a component that draws a `.docx` from its
// bytes alone, a component that draws a laid-out page, a component that draws one
// page of it, the three helpers a caller needs to hand it pictures, and the one
// that hands a written page to the browser as a file.
const VIEWER_SURFACE: readonly string[] = [
  "Document",
  "DocxDocument",
  "Page",
  "downloadPdf",
  "drawablesOf",
  "imageDataUrl",
  "imageResolver",
];

describe("the public surface of @docx-pages/viewer", () => {
  it("exports what it promises, and nothing has quietly gone away", () => {
    const found = Object.keys(viewer)
      .filter((name) => name !== "default")
      .sort();

    expect(found).toStrictEqual([...VIEWER_SURFACE].sort());
  });
});

describe("the public surface of @docx-pages/viewer/pack", () => {
  it("offers the same component with the pack already behind it", () => {
    const found = Object.keys(pack)
      .filter((name) => name !== "default")
      .sort();

    expect(found).toStrictEqual(["DocxDocument"]);
  });
});
