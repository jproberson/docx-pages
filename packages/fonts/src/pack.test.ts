import { describe, expect, it } from "vitest";

import { bestEffortMetrics, layOutDocument, openDocx } from "@docx-pages/core";
import { buildDocx, wordDocument } from "@docx-pages/core/testing";

import { METRIC_TWINS, PACK_FACES } from "./index.js";
import { defaultFacesFromDisk } from "./node.js";

const SECTION = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// Declares its default face as any document Word saved would, so the paragraph
// mark resolves the same way the run does.
const naming = (face: string): Uint8Array =>
  buildDocx({
    "word/document.xml": wordDocument(
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="${face}" w:hAnsi="${face}"/><w:sz w:val="24"/></w:rPr><w:t>Measured words</w:t></w:r></w:p>${SECTION}`,
    ),
    "word/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${face}" w:hAnsi="${face}"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
  });

describe("the default pack", () => {
  it("reads every face it ships, and each knows its own widths", async () => {
    const defaults = await defaultFacesFromDisk();

    expect(defaults.faces).toHaveLength(PACK_FACES.length);
    for (const face of defaults.faces) {
      expect(face.advances.kind, face.name).toBe("advances");
    }
  });

  it("lays out a document naming Calibri with nothing supplied at all, and says so", async () => {
    const defaults = await defaultFacesFromDisk();
    const faces = bestEffortMetrics([], defaults);
    const laid = layOutDocument(openDocx(naming("Calibri")), faces);

    expect(laid.kind).toBe("laid-out");
    expect(faces.substitutions()).toStrictEqual([
      {
        requested: { name: "Calibri", bold: false, italic: false },
        used: { name: "Carlito", bold: false, italic: false },
      },
    ]);
    if (laid.kind === "laid-out") {
      expect(laid.unhonoured.map((entry) => entry.kind)).toContain("substituted-face");
    }
  });

  it("answers for every name the twin table promises", async () => {
    const defaults = await defaultFacesFromDisk();

    for (const [named, twin] of Object.entries(METRIC_TWINS)) {
      const faces = bestEffortMetrics([], defaults);
      const found = faces.metricsFor({ name: named, bold: false, italic: false });
      expect(found.kind, named).toBe("found");
      expect(
        faces.substitutions().map((each) => each.used.name),
        named,
      ).toStrictEqual([twin]);
    }
  });
});
