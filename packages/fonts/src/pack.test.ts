import { describe, expect, it } from "vitest";

import { bestEffortMetrics, layOutDocument, openDocx } from "@docx-pages/core";
import { buildDocx, wordDocument } from "@docx-pages/core/testing";

import { METRIC_TWINS, PACK_FACES, readPack } from "./index.js";
import { defaultFacesFromDisk, readFromDisk } from "./node.js";

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

  // A page is drawn as well as measured, and both need the same file: the browser
  // is offered the face the layout measured, and a pdf carries the ones it draws
  // in. So the bytes come back beside the metrics rather than being read twice.
  it("hands back the file each face was read out of, beside what was read from it", async () => {
    const pack = await readPack(readFromDisk);

    expect(pack.bytes.map((each) => each.name)).toStrictEqual(
      pack.defaults.faces.map((face) => face.name),
    );
    for (const face of pack.bytes) {
      expect(face.bytes.byteLength, face.name).toBeGreaterThan(0);
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

  it("answers an unknown sans name with Open Sans, marked as the sans it is", async () => {
    const defaults = await defaultFacesFromDisk();
    const faces = bestEffortMetrics([], defaults, new Map([["housemade sans", "sans-serif"]]));

    const found = faces.metricsFor({ name: "Housemade Sans", bold: false, italic: false });
    expect(found.kind).toBe("found");
    expect(faces.substitutions().map((each) => each.used.name)).toStrictEqual(["Open Sans"]);
    // The file's own PANOSE bytes misstate the face; the pack corrects them, so
    // a character borrowed for it goes through the sans chain.
    expect(defaults.faces.find((face) => face.name === "Open Sans")?.sansSerif).toBe(true);
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
