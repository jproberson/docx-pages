import { describe, expect, it } from "vitest";

import { readBlocks } from "../docx/blocks.js";
import type { DrawingContent } from "../docx/drawing.js";
import { readInlines } from "../docx/inlines.js";
import { openDocx } from "../docx/package.js";
import { NO_THEME } from "../docx/theme.js";
import { readStyleTable } from "../docx/styles.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { lookupFontMetrics } from "./font-metrics.js";
import { measureStack, type ParagraphBox } from "./stack.js";
import { placeInlines } from "./inlines.js";
import type { WrapBand } from "./wrapping.js";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

// The body's own column on a letter page with half-inch margins.
const LEFT_PT = 36;
const WIDTH_PT = 540;

// 180pt by 90pt.
const image = (cx = 2286000, cy = 1143000) =>
  `<w:r><w:drawing><wp:inline xmlns:wp="${WP_NS}">
     <wp:extent cx="${String(cx)}" cy="${String(cy)}"/>
     <wp:docPr id="1" name="Logo"/></wp:inline></w:drawing></w:r>`;

const V_NS = "urn:schemas-microsoft-com:vml";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const legacy = (id: string, width: string, height: string, position = "") =>
  `<w:r><w:pict><v:shape xmlns:v="${V_NS}" type="#_x0000_t75" style="${position}width:${width};height:${height}">
     <v:imagedata xmlns:r="${R_NS}" r:id="${id}"/></v:shape></w:pict></w:r>`;

const contentOf = (
  placed: { readonly drawing: { readonly content: DrawingContent } } | undefined,
) => (placed?.drawing.content.kind === "picture" ? placed.drawing.content.relationshipId : null);

const paragraph = (properties: string, runs: string) =>
  `<w:p>${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}${runs}</w:p>`;

const metricsFor = lookupFontMetrics;

// The drawings are placed the way the document places them: the paragraph is laid
// out first, and each drawing lands where its own line put it.
const place = (body: string, topPt = 100, bands: readonly WrapBand[] = []) => {
  const pkg = openDocx(
    buildDocx({ "word/document.xml": wordDocument(body), "word/styles.xml": STYLES }),
  );
  const blocks = readBlocks(pkg);
  const measured = measureStack({
    blocks,
    styles: readStyleTable(pkg),
    metricsFor,
    part: "word/document.xml",
    originPt: topPt,
    leftPt: LEFT_PT,
    widthPt: WIDTH_PT,
    bandsFor: () => bands,
  });
  if (measured.kind === "blocked") throw new Error(`blocked: ${measured.blocker.kind}`);

  const box = measured.boxes[0];
  const found = blocks.flatMap((block) => (block.kind === "paragraph" ? [block.paragraph] : []))[0];
  if (box === undefined || found === undefined) throw new Error("expected a paragraph");

  return {
    box,
    placed: placeInlines({
      drawings: readInlines(found),
      box,
      resolvePart: () => null,
      theme: NO_THEME,
    }),
  };
};

// A face the builtin metrics answer for, so a paragraph mark can be measured.
const STYLES = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Figure"><w:pPr><w:jc w:val="right"/></w:pPr></w:style>
  </w:styles>`;

const leftsOf = (box: ParagraphBox): readonly number[] => box.lines.map((line) => line.leftPt);

describe("placeInlines", () => {
  it("starts an unaligned drawing at the left margin", () => {
    const { placed } = place(paragraph("", image()));
    expect(placed[0]?.leftPt).toBeCloseTo(36, 6);
    expect(placed[0]?.widthPt).toBeCloseTo(180, 6);
  });

  // A drawing stands on its line's baseline, which for a line holding nothing else
  // puts its top at the paragraph's own.
  it("seats the drawing at its paragraph's top", () => {
    const { placed } = place(paragraph("", image()), 253.87);
    expect(placed[0]?.topPt).toBeCloseTo(253.87, 6);
    expect(placed[0]?.heightPt).toBeCloseTo(90, 6);
  });

  it("ends a right-aligned drawing at the right margin", () => {
    const { placed } = place(paragraph(`<w:jc w:val="right"/>`, image()));
    expect(placed[0]?.leftPt).toBeCloseTo(612 - 36 - 180, 6);
  });

  it("centres between the margins", () => {
    const { placed } = place(paragraph(`<w:jc w:val="center"/>`, image()));
    expect(placed[0]?.leftPt).toBeCloseTo(36 + (540 - 180) / 2, 6);
  });

  it("narrows the line by the paragraph's indents", () => {
    const indented = `<w:jc w:val="right"/><w:ind w:left="720" w:right="1440"/>`;
    const { placed } = place(paragraph(indented, image()));
    expect(placed[0]?.leftPt).toBeCloseTo(612 - 36 - 72 - 180, 6);
  });

  it("runs several drawings along the line rather than stacking them", () => {
    const { placed } = place(paragraph("", `${image()}${image(1143000, 1143000)}`));
    expect(placed.map((each) => each.leftPt)).toStrictEqual([36, 216]);
  });

  it("aligns a run of drawings by their combined width", () => {
    const { placed } = place(paragraph(`<w:jc w:val="right"/>`, `${image()}${image()}`));
    expect(placed[0]?.leftPt).toBeCloseTo(612 - 36 - 360, 6);
  });

  it("takes the alignment a paragraph style sets when the paragraph sets none", () => {
    const { placed } = place(paragraph(`<w:pStyle w:val="Figure"/>`, image()));
    expect(placed[0]?.leftPt).toBeCloseTo(612 - 36 - 180, 6);
  });

  // Measured against Word: an inline picture right-aligned beside a wrapping
  // object ends where that object's band begins, not at the margin behind it.
  it("keeps a drawing out of the band an object beside it wraps", () => {
    const band: WrapBand = { leftPt: 400, rightPt: 600, topPt: 0, bottomPt: 400 };
    const { placed, box } = place(paragraph(`<w:jc w:val="right"/>`, image()), 100, [band]);
    expect(leftsOf(box)).toStrictEqual([220]);
    expect(placed[0]?.leftPt).toBeCloseTo(220, 6);
  });
});

// The form Word wrote before DrawingML and still writes for some of what it
// draws. It stands on the line exactly as a `wp:inline` does, and states its size
// in the css of the shape's `style` rather than in an extent.
describe("placeInlines over a picture written the legacy way", () => {
  it("places one at the size its style states", () => {
    const { placed } = place(paragraph("", legacy("rId7", "180pt", "90pt")));
    expect(placed[0]?.leftPt).toBeCloseTo(36, 6);
    expect(placed[0]?.widthPt).toBeCloseTo(180, 6);
    expect(placed[0]?.heightPt).toBeCloseTo(90, 6);
    expect(contentOf(placed[0])).toBe("rId7");
  });

  it("aligns and seats one as it does a drawing", () => {
    const { placed } = place(
      paragraph(`<w:jc w:val="right"/>`, legacy("rId7", "180pt", "90pt")),
      253.87,
    );
    expect(placed[0]?.leftPt).toBeCloseTo(612 - 36 - 180, 6);
    expect(placed[0]?.topPt).toBeCloseTo(253.87, 6);
  });

  // **The order is the whole of why this is worth a test.** A drawing is matched to
  // its place on the line by counting, so reading every `wp:inline` and then every
  // `w:pict` would hand a paragraph holding both each other's pictures.
  it("keeps the two forms in the order the paragraph writes them", () => {
    const { placed } = place(
      paragraph("", `${legacy("rId7", "72pt", "72pt")}${image()}${legacy("rId9", "36pt", "36pt")}`),
    );
    expect(placed.map((each) => each.widthPt)).toStrictEqual([72, 180, 36]);
    expect(placed.map(contentOf)).toStrictEqual(["rId7", null, "rId9"]);
  });

  // A VML shape carrying `position:absolute` is out of flow like a `wp:anchor`,
  // and a run holding nothing else places nothing on the line at all.
  it("leaves a positioned one off the line", () => {
    const { placed, box } = place(
      paragraph("", `${legacy("rId7", "180pt", "90pt", "position:absolute;")}${image()}`),
    );
    expect(placed.map((each) => each.widthPt)).toStrictEqual([180]);
    expect(placed.map(contentOf)).toStrictEqual([null]);
    expect(box.lines[0]?.line.widthPt).toBeCloseTo(180, 6);
  });

  it("passes over a shape that names no picture", () => {
    const { placed } = place(
      paragraph(
        "",
        `<w:r><w:pict><v:rect xmlns:v="${V_NS}" style="width:180pt;height:90pt"/></w:pict></w:r>`,
      ),
    );
    expect(placed).toStrictEqual([]);
  });
});
