import { describe, expect, it } from "vitest";

import { readAnchors, type FloatingAnchor } from "../docx/anchors.js";
import { openDocx } from "../docx/package.js";
import { readParagraphs } from "../docx/paragraphs.js";
import type { SectionGeometry } from "../docx/section.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { placeFloat } from "./floats.js";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

const LETTER: SectionGeometry = {
  widthTwips: 12240,
  heightTwips: 15840,
  margin: {
    topTwips: 720,
    rightTwips: 720,
    bottomTwips: 0,
    leftTwips: 720,
    headerTwips: 432,
    footerTwips: 144,
  },
};

const anchorXml = (options: {
  h: string;
  v: string;
  cx?: number;
  cy?: number;
  wrap?: string;
}) => `<w:p><w:r><w:drawing><wp:anchor xmlns:wp="${WP_NS}" behindDoc="0" relativeHeight="5">
  <wp:extent cx="${String(options.cx ?? 2286000)}" cy="${String(options.cy ?? 904240)}"/>
  ${options.wrap ?? "<wp:wrapNone/>"}
  <wp:docPr id="1" name="Picture 12"/>
  ${options.h}${options.v}
</wp:anchor></w:drawing></w:r></w:p>`;

const offsetH = (emu: number, from = "column") =>
  `<wp:positionH relativeFrom="${from}"><wp:posOffset>${String(emu)}</wp:posOffset></wp:positionH>`;
const offsetV = (emu: number, from = "paragraph") =>
  `<wp:positionV relativeFrom="${from}"><wp:posOffset>${String(emu)}</wp:posOffset></wp:positionV>`;
const alignH = (align: string, from = "margin") =>
  `<wp:positionH relativeFrom="${from}"><wp:align>${align}</wp:align></wp:positionH>`;

const firstAnchor = (body: string): FloatingAnchor => {
  const pkg = openDocx(buildDocx({ "word/document.xml": wordDocument(body) }));
  const paragraph = readParagraphs(pkg)[0];
  if (paragraph === undefined) throw new Error("expected a paragraph");
  const anchor = readAnchors(paragraph)[0];
  if (anchor === undefined) throw new Error("expected an anchor");
  return anchor;
};

const place = (body: string, paragraphTopPt: number, bodyTopPt = 36) =>
  placeFloat({ anchor: firstAnchor(body), page: LETTER, paragraphTopPt, bodyTopPt });

describe("readAnchors", () => {
  it("reads extent, wrap and stacking order", () => {
    const anchor = firstAnchor(anchorXml({ h: offsetH(0), v: offsetV(0) }));
    expect(anchor.name).toBe("Picture 12");
    expect(anchor.widthEmu).toBe(2286000);
    expect(anchor.wrap).toBe("none");
    expect(anchor.behindDoc).toBe(false);
    expect(anchor.relativeHeight).toBe(5);
  });

  it("reads a wrapSquare float", () => {
    const anchor = firstAnchor(
      anchorXml({ h: offsetH(0), v: offsetV(0), wrap: '<wp:wrapSquare wrapText="bothSides"/>' }),
    );
    expect(anchor.wrap).toBe("square");
  });

  it("distinguishes an aligned position from an offset one", () => {
    const anchor = firstAnchor(anchorXml({ h: alignH("right"), v: offsetV(0) }));
    expect(anchor.horizontal).toStrictEqual({ kind: "align", from: "margin", align: "right" });
  });
});

describe("placeFloat", () => {
  it("resolves a column offset against the left margin", () => {
    // The Reference header logo: 5199353 EMU right of the column origin.
    expect(place(anchorXml({ h: offsetH(5199353), v: offsetV(0) }), 21.6).leftPt).toBeCloseTo(
      445.398,
      3,
    );
  });

  it("lets a negative vertical offset escape upward out of its paragraph", () => {
    // Word puts this logo above the header paragraph it is anchored to.
    expect(place(anchorXml({ h: offsetH(0), v: offsetV(-162119) }), 21.6).topPt).toBeCloseTo(
      8.835,
      3,
    );
  });

  it("resolves a paragraph offset against that paragraph's top", () => {
    expect(place(anchorXml({ h: offsetH(0), v: offsetV(1884045) }), 346.286).topPt).toBeCloseTo(
      494.636,
      3,
    );
  });

  it("measures a page-relative offset from the page edge, not the margin", () => {
    expect(place(anchorXml({ h: offsetH(0, "page"), v: offsetV(0) }), 100).leftPt).toBe(0);
  });

  it("aligns right against the far margin, allowing for the float's own width", () => {
    const placed = place(anchorXml({ h: alignH("right"), v: offsetV(0) }), 100);
    expect(placed.leftPt).toBeCloseTo(612 - 36 - 180, 6);
  });

  it("centres within the margins", () => {
    const placed = place(anchorXml({ h: alignH("center"), v: offsetV(0) }), 100);
    expect(placed.leftPt).toBeCloseTo(36 + (540 - 180) / 2, 6);
  });

  it("converts the extent into points", () => {
    const placed = place(anchorXml({ h: offsetH(0), v: offsetV(0) }), 100);
    expect(placed.widthPt).toBeCloseTo(180, 6);
    expect(placed.heightPt).toBeCloseTo(71.2, 3);
  });
});
