import { describe, expect, it } from "vitest";

import { openDocx } from "../docx/package.js";
import { buildDocx, WORDPROCESSING_NS } from "../testing/build-docx.js";
import { buildFace } from "../testing/build-font.js";
import { bestEffortMetrics, type FaceDefaults } from "./best-effort.js";
import { layOutDocument, type LaidOutDocument } from "./document.js";

const METRICS = { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 };
const TWIN = buildFace({ name: "Twin Sans", metrics: METRICS, sansSerif: true });

const DEFAULTS: FaceDefaults = {
  faces: [TWIN],
  twins: {},
  sansSerif: "Twin Sans",
  serif: "Twin Sans",
  monospace: "Twin Sans",
  lastResort: "Twin Sans",
};

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// A page keeping whatever it is told between its edge and its text, and between
// its edge and the room it holds for a header.
const section = (topTwips: number, headerTwips: number, header: boolean): string =>
  `<w:sectPr>` +
  (header ? `<w:headerReference w:type="default" r:id="rId1"/>` : "") +
  `<w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="${String(topTwips)}" w:right="720" w:bottom="720" w:left="720"` +
  ` w:header="${String(headerTwips)}" w:footer="720" w:gutter="0"/>` +
  `</w:sectPr>`;

const HEADER_PARTS: Readonly<Record<string, string>> = {
  "word/header1.xml": `<?xml version="1.0"?>
    <w:hdr xmlns:w="${WORDPROCESSING_NS}"><w:p><w:r><w:t>above</w:t></w:r></w:p></w:hdr>`,
  "word/_rels/document.xml.rels": `<?xml version="1.0"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Target="header1.xml" Type="${R_NS}/header"/>
    </Relationships>`,
};

function laidOut(topTwips: number, headerTwips: number, header: boolean): LaidOutDocument {
  const body = `<w:p><w:r><w:t>first</w:t></w:r></w:p>${section(topTwips, headerTwips, header)}`;
  // Written out rather than taken from `wordDocument`, which declares no `r` prefix
  // for a header to be referenced through.
  const bytes = buildDocx({
    "word/document.xml":
      `<?xml version="1.0"?><w:document xmlns:w="${WORDPROCESSING_NS}" xmlns:r="${R_NS}">` +
      `<w:body>${body}</w:body></w:document>`,
    ...(header ? HEADER_PARTS : {}),
  });
  const laid = layOutDocument(openDocx(bytes), bestEffortMetrics([], DEFAULTS));
  if (laid.kind !== "laid-out") throw new Error(`blocked: ${laid.blocker.kind}`);
  return laid;
}

describe("where a page starts its body", () => {
  // The room a header would have taken is not taken by a header that is not there,
  // which is what the footer has always said on its own side.
  it("starts at the top margin where the page draws no header", () => {
    const laid = laidOut(20, 720, false);

    expect(laid.bodyTopPt).toBe(1);
    expect(laid.pages[0]?.body[0]?.topPt).toBe(1);
  });

  it("starts under a header that reaches past the top margin", () => {
    const laid = laidOut(20, 720, true);

    expect(laid.bodyTopPt).toBe(laid.headerTopPt + (laid.pages[0]?.headerHeightPt ?? 0));
    expect(laid.bodyTopPt).toBeGreaterThan(36);
  });

  it("leaves the body at the top margin where the header stops above it", () => {
    const laid = laidOut(5760, 720, true);

    expect(laid.bodyTopPt).toBe(288);
  });
});
