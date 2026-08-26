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

// The same, turned about its own middle: the flow keeps it the room the turn rounds to.
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const turnedImage = (cx: number, cy: number, degrees: number) =>
  `<w:r><w:drawing><wp:inline xmlns:wp="${WP_NS}">
     <wp:extent cx="${String(cx)}" cy="${String(cy)}"/>
     <wp:docPr id="1" name="Logo"/>
     <a:graphic xmlns:a="${A_NS}"><a:graphicData uri="${PIC_NS}">
       <pic:pic xmlns:pic="${PIC_NS}">
         <pic:spPr><a:xfrm rot="${String(degrees * 60000)}"/></pic:spPr>
       </pic:pic>
     </a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const V_NS = "urn:schemas-microsoft-com:vml";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const legacy = (id: string, width: string, height: string, position = "", holder = "pict") =>
  `<w:r><w:${holder}><v:shape xmlns:v="${V_NS}" type="#_x0000_t75" style="${position}width:${width};height:${height}">
     <v:imagedata xmlns:r="${R_NS}" r:id="${id}"/></v:shape></w:${holder}></w:r>`;

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

  // **A paragraph whose lines are told exactly how tall to be cuts off a drawing that
  // hangs above it.** Measured on 2026-08-25 by `exact-line-clip-probe`, six cases read
  // off Word's own pdf: a picture 150 by 100 alone in paragraphs of exactly 10, 20, 40,
  // 60, 100 and 120 showed 8.01, 16.01, 32.01, 48.01, 80.01 and 96.01 of its height,
  // each starting at the paragraph's own top.
  it("cuts a drawing off at the paragraph where the line is told to be shorter", () => {
    const exactly = `<w:spacing w:line="200" w:lineRule="exact"/>`;
    const { placed } = place(paragraph(exactly, image()), 100);
    const cut = placed[0]?.clipTo;
    expect(cut?.topPt).toBeCloseTo(100, 6);
    expect(cut?.heightPt).toBeCloseTo(10, 6);
    // The drawing itself is placed whole, hanging from the baseline as ever.
    expect(placed[0]?.topPt).toBeCloseTo(108 - 90, 6);
    expect(placed[0]?.heightPt).toBeCloseTo(90, 6);
  });

  // The same picture with an empty line above it in the same paragraph showed 18.01,
  // and with twelve above it the whole of its 100: what the drawing hangs over is its
  // paragraph's own room and not the line it stands on.
  it("keeps what a drawing hangs over its own paragraph's earlier lines", () => {
    const exactly = `<w:spacing w:line="200" w:lineRule="exact"/>`;
    const twelve = `<w:r>${"<w:br/>".repeat(12)}</w:r>`;
    expect(place(paragraph(exactly, twelve + image()), 100).placed[0]?.clipTo).toBeNull();
  });

  // The cut is the paragraph's own room, two lines of 10 here, and what shows of the
  // picture is the 18 of it between that room's top and its own bottom on the second
  // baseline. Word drew 18.01.
  it("cuts at the paragraph's top where the drawing hangs past it", () => {
    const exactly = `<w:spacing w:line="200" w:lineRule="exact"/>`;
    const { placed } = place(paragraph(exactly, `<w:r><w:br/></w:r>${image()}`), 100);
    const cut = placed[0]?.clipTo;
    expect(cut?.topPt).toBeCloseTo(100, 6);
    expect(cut?.heightPt).toBeCloseTo(20, 6);
    const drawing = placed[0];
    if (drawing === undefined) throw new Error("expected the drawing");
    expect(drawing.topPt + drawing.heightPt - (cut?.topPt ?? 0)).toBeCloseTo(18, 6);
  });

  // **A drawing on its side is measured by the paint it reaches.** The flow keeps it the
  // room its turn rounds to, so a picture stored 90 wide by 180 tall stands 180 by 90 in
  // a paragraph 90 tall: read the stored box and it looks as though it hangs 90 out of a
  // paragraph it sits inside. `e199f3435eaf` came out cut down its sides for it.
  it("cuts nothing off a turned drawing whose paint stands inside its paragraph", () => {
    const { placed } = place(paragraph("", turnedImage(1143000, 2286000, 90)), 100);
    expect(placed[0]?.clipTo).toBeNull();
    expect(placed[0]?.widthPt).toBeCloseTo(90, 6);
    expect(placed[0]?.heightPt).toBeCloseTo(180, 6);
  });

  it("cuts nothing where the line is tall enough to hold the drawing", () => {
    const exactly = `<w:spacing w:line="2400" w:lineRule="exact"/>`;
    expect(place(paragraph(exactly, image()), 100).placed[0]?.clipTo).toBeNull();
  });

  // A line grown to hold the drawing holds all of it, which is every line stating no
  // rule and every one stating `atLeast`.
  it("cuts nothing where the line measured itself", () => {
    expect(place(paragraph("", image()), 100).placed[0]?.clipTo).toBeNull();
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

  // What Word draws for an embedded spreadsheet or equation is a picture of it
  // written exactly this way, so the container it stands in is the only difference.
  it("draws the picture a w:object stands on the line for what it embeds", () => {
    const { placed } = place(paragraph("", legacy("rId7", "180pt", "90pt", "", "object")));
    expect(placed[0]?.widthPt).toBeCloseTo(180, 6);
    expect(contentOf(placed[0])).toBe("rId7");
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
