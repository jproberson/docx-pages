import { describe, expect, it } from "vitest";

import { openDocx } from "../docx/package.js";
import { buildDocx, WORDPROCESSING_NS } from "../testing/build-docx.js";
import { buildFace } from "../testing/build-font.js";
import { bestEffortMetrics, type FaceDefaults } from "./best-effort.js";
import { layOutDocument } from "./document.js";
import { drawablesOf, type HighlightPaint } from "./drawables.js";

// A face whose line is exactly the size it is set at, so a height in a case below is
// the rule under test and not the face's own leading.
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

const SECTION = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"
    w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

const STYLES = `<?xml version="1.0"?>
  <w:styles xmlns:w="${WORDPROCESSING_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Twin Sans" w:hAnsi="Twin Sans"/><w:sz w:val="24"/>
  </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`;

function highlightsOf(body: string): readonly HighlightPaint[] {
  const pkg = openDocx(
    buildDocx({
      "word/document.xml": `<?xml version="1.0"?><w:document xmlns:w="${WORDPROCESSING_NS}">
        <w:body>${body}${SECTION}</w:body></w:document>`,
      "word/styles.xml": STYLES,
    }),
  );
  const laid = layOutDocument(pkg, bestEffortMetrics([], DEFAULTS));
  if (laid.kind !== "laid-out") throw new Error(laid.blocker.kind);
  const page = laid.pages[0];
  if (page === undefined) throw new Error("expected a page");
  return drawablesOf(laid, page).flatMap((drawable) =>
    drawable.kind === "paint" ? drawable.highlights : [],
  );
}

const run = (text: string, properties = ""): string =>
  `<w:r><w:rPr>${properties}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;

const YELLOW = `<w:highlight w:val="yellow"/>`;
const size = (halfPoints: number): string => `<w:sz w:val="${String(halfPoints)}"/>`;

describe("what a highlight paints", () => {
  it("paints the run's own advance in the colour the name stands for", () => {
    const [painted] = highlightsOf(`<w:p>${run("one ")}${run("two", YELLOW)}</w:p>`);
    expect(painted?.color).toBe("#ffff00");
    expect(painted?.leftPt).toBeCloseTo(36 + 4 * 12 * 0.5, 5);
    expect(painted?.widthPt).toBeCloseTo(3 * 12 * 0.5, 5);
  });

  it("paints nothing for a run that turns an inherited highlight off", () => {
    const off = `<w:highlight w:val="none"/>`;
    expect(highlightsOf(`<w:p>${run("one", off)}</w:p>`)).toStrictEqual([]);
  });

  // Word paints the line rather than the run: a 12pt run on a line holding 24pt text
  // came out 29.28pt tall, the same as the line of 24pt throughout.
  it("is as tall as the line, not as the run standing on it", () => {
    const body = `<w:p>${run("big ", size(48))}${run("small", YELLOW)}</w:p>`;
    const [painted] = highlightsOf(body);
    expect(painted?.heightPt).toBeCloseTo(24, 5);
  });

  it("is as tall as a raised run's own line", () => {
    const raised = `${YELLOW}<w:vertAlign w:val="superscript"/>`;
    const [painted] = highlightsOf(`<w:p>${run("one ")}${run("two", raised)}</w:p>`);
    expect(painted?.heightPt).toBeCloseTo(12, 5);
  });

  // An exact rule is a slot the text is dropped into and Word paints the whole of it;
  // the room a multiple opens below the text is not painted at all.
  it("fills the slot an exact line rule states", () => {
    const rule = `<w:pPr><w:spacing w:line="480" w:lineRule="exact"/></w:pPr>`;
    const [painted] = highlightsOf(`<w:p>${rule}${run("one", YELLOW)}</w:p>`);
    expect(painted?.heightPt).toBeCloseTo(24, 5);
  });

  it("leaves the room a line multiple opens below the text unpainted", () => {
    const rule = `<w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr>`;
    const [painted] = highlightsOf(`<w:p>${rule}${run("one", YELLOW)}</w:p>`);
    expect(painted?.heightPt).toBeCloseTo(12, 5);
  });

  // Word's own pdf of an empty paragraph whose mark states a highlight holds no fill.
  it("paints nothing for an empty paragraph whose mark is highlighted", () => {
    const mark = `<w:pPr><w:rPr>${YELLOW}</w:rPr></w:pPr>`;
    expect(highlightsOf(`<w:p>${mark}</w:p>`)).toStrictEqual([]);
  });
});
