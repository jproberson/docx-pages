import { describe, expect, it } from "vitest";

import * as viewer from "./index.js";

// What `@docx-pages/viewer` promises anyone who installs it. The same argument as
// core's list, and a much shorter one: a component that draws a laid-out page, a
// component that draws one page of it, and the three helpers a caller needs to
// hand it pictures.
const VIEWER_SURFACE: readonly string[] = [
  "Document",
  "Page",
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
