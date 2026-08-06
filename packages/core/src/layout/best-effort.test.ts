import { describe, expect, it } from "vitest";

import { readFaceShapes } from "../docx/font-table.js";
import { openDocx } from "../docx/package.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { buildFace } from "../testing/build-font.js";
import { bestEffortMetrics, type FaceDefaults } from "./best-effort.js";
import { layOutDocument } from "./document.js";
import type { SuppliedFace } from "./font-metrics.js";

const METRICS = { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 };

// The pack the defaults stand on, built like real files so that each face knows
// its own missing-glyph advance.
const pack = (name: string, sansSerif: boolean): SuppliedFace =>
  buildFace({ name, metrics: METRICS, sansSerif, notdefAdvance: 350 });

const DEFAULTS: FaceDefaults = {
  faces: [pack("Twin Sans", true), pack("Twin Serif", false), pack("Twin Mono", false)],
  twins: { "meridian text": "Twin Serif" },
  sansSerif: "Twin Sans",
  serif: "Twin Serif",
  monospace: "Twin Mono",
};

const ask = (name: string) => ({ name, bold: false, italic: false });

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const SECTION = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;

describe("bestEffortMetrics", () => {
  it("answers with the twin whose widths match the name", () => {
    const faces = bestEffortMetrics([], DEFAULTS);
    const found = faces.metricsFor(ask("Meridian Text"));

    expect(found.kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Meridian Text"), used: ask("Twin Serif") },
    ]);
  });

  it("answers a classified name with the default of its shape", () => {
    const shapes = new Map([["quill display", "serif" as const]]);
    const faces = bestEffortMetrics([], DEFAULTS, shapes);

    expect(faces.metricsFor(ask("Quill Display")).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Quill Display"), used: ask("Twin Serif") },
    ]);
  });

  it("answers a name nothing classifies with the sans default", () => {
    const faces = bestEffortMetrics([], DEFAULTS);

    expect(faces.metricsFor(ask("Unheard Of")).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Unheard Of"), used: ask("Twin Sans") },
    ]);
  });

  it("prefers a face the caller supplied over every default", () => {
    const supplied = buildFace({ name: "Meridian Text", metrics: METRICS });
    const faces = bestEffortMetrics([supplied], DEFAULTS);

    expect(faces.metricsFor(ask("Meridian Text")).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([]);
  });

  it("answers a character nothing maps with the missing-glyph box, and says so", () => {
    const faces = bestEffortMetrics([], DEFAULTS);
    const found = faces.metricsFor(ask("Unheard Of"));
    if (found.kind !== "found" || found.advances.kind !== "advances") throw new Error("unusable");

    // A code point no fixture maps: the box answers through `elsewhere`, at the
    // advance the pack face declares for it.
    expect(found.advances.advanceFor(0x2603)).toBeNull();
    expect(found.elsewhere?.(0x2603)).toStrictEqual({ metrics: METRICS, advance: 350 });
    expect(faces.missingGlyphs()).toStrictEqual([{ face: ask("Unheard Of"), codePoint: 0x2603 }]);
  });

  it("lays out a document whose cascade names no face at all, loudly", () => {
    // No styles part and no run properties: the cascade has nothing to resolve.
    const bytes = buildDocx({
      "word/document.xml": wordDocument(`<w:p><w:r><w:t>abc</w:t></w:r></w:p>${SECTION}`),
    });
    const faces = bestEffortMetrics([], DEFAULTS);
    const laid = layOutDocument(openDocx(bytes), faces);

    expect(laid.kind).toBe("laid-out");
    expect(faces.substitutions()).toStrictEqual([{ requested: ask(""), used: ask("Twin Sans") }]);
  });

  it("folds every box and stand-in into what the layout reports", () => {
    const document = `<w:p><w:r><w:rPr><w:rFonts w:ascii="Unheard Of" w:hAnsi="Unheard Of"/></w:rPr><w:t xml:space="preserve">a☃</w:t></w:r></w:p>${SECTION}`;
    const bytes = buildDocx({
      "word/document.xml": wordDocument(document),
      "word/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Twin Sans" w:hAnsi="Twin Sans"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    });
    const faces = bestEffortMetrics([], DEFAULTS);
    const laid = layOutDocument(openDocx(bytes), faces);
    if (laid.kind !== "laid-out") throw new Error(JSON.stringify(laid.blocker));

    const kinds = laid.unhonoured.map((entry) => entry.kind);
    expect(kinds).toContain("substituted-face");
    expect(kinds).toContain("missing-glyph");
  });

  it("reads the document's own classification for the shape", () => {
    const bytes = buildDocx({
      "word/document.xml": wordDocument(`<w:p><w:r><w:t>a</w:t></w:r></w:p>${SECTION}`),
      "word/fontTable.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="${W_NS}"><w:font w:name="Quill Display"><w:family w:val="roman"/></w:font></w:fonts>`,
    });
    const shapes = readFaceShapes(openDocx(bytes));
    const faces = bestEffortMetrics([], DEFAULTS, shapes);

    expect(faces.metricsFor(ask("Quill Display")).kind).toBe("found");
    expect(faces.substitutions()).toStrictEqual([
      { requested: ask("Quill Display"), used: ask("Twin Serif") },
    ]);
  });
});
