import { describe, expect, it } from "vitest";

import { readAnchors, WHOLE_FRAME, type FloatingAnchor, type WrapArea } from "../docx/anchors.js";
import { openDocx } from "../docx/package.js";
import { readParagraphs } from "../docx/blocks.js";
import type { SectionGeometry } from "../docx/section.js";
import { DEFAULT_SETTINGS } from "../docx/settings.js";
import { NO_THEME } from "../docx/theme.js";
import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { placeFloat, UNPAINTED } from "./floats.js";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

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

const DISTANCES = `distT="45720" distB="45720" distL="114300" distR="114300"`;

const anchorXml = (options: {
  h: string;
  v: string;
  cx?: number;
  cy?: number;
  wrap?: string;
  distances?: boolean;
  flip?: string;
}) => `<w:p><w:r><w:drawing><wp:anchor xmlns:wp="${WP_NS}" xmlns:a="${A_NS}"
  behindDoc="0" relativeHeight="5" ${options.distances === true ? DISTANCES : ""}>
  <wp:extent cx="${String(options.cx ?? 2286000)}" cy="${String(options.cy ?? 904240)}"/>
  ${options.wrap ?? "<wp:wrapNone/>"}
  <wp:docPr id="1" name="Logo"/>
  ${options.h}${options.v}
  <a:graphic><a:graphicData><a:xfrm ${options.flip ?? ""}/></a:graphicData></a:graphic>
</wp:anchor></w:drawing></w:r></w:p>`;

// A polygon over the object's own 21600ths, kept off the frame's left edge and
// its foot.
const tightWrap = (left: number, bottom: number) =>
  `<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon>
     <wp:start x="${String(left)}" y="0"/>
     <wp:lineTo x="${String(left)}" y="${String(bottom)}"/>
     <wp:lineTo x="21600" y="${String(bottom)}"/>
     <wp:lineTo x="21600" y="0"/>
     <wp:lineTo x="${String(left)}" y="0"/>
   </wp:wrapPolygon></wp:wrapTight>`;

const offsetH = (emu: number, from = "column") =>
  `<wp:positionH relativeFrom="${from}"><wp:posOffset>${String(emu)}</wp:posOffset></wp:positionH>`;
const offsetV = (emu: number, from = "paragraph") =>
  `<wp:positionV relativeFrom="${from}"><wp:posOffset>${String(emu)}</wp:posOffset></wp:positionV>`;
const alignH = (align: string, from = "margin") =>
  `<wp:positionH relativeFrom="${from}"><wp:align>${align}</wp:align></wp:positionH>`;

// The rectangle a wrap area comes down to, which most of these ask about; the
// corners it was drawn from have their own test.
const edgesOf = (area: WrapArea) => ({
  left: area.left,
  top: area.top,
  right: area.right,
  bottom: area.bottom,
});

const firstAnchor = (body: string): FloatingAnchor => {
  const pkg = openDocx(buildDocx({ "word/document.xml": wordDocument(body) }));
  const paragraph = readParagraphs(pkg)[0];
  if (paragraph === undefined) throw new Error("expected a paragraph");
  const anchor = readAnchors(paragraph)[0];
  if (anchor === undefined) throw new Error("expected an anchor");
  return anchor;
};

// Every position below was measured off a document Word itself wrote, which
// declares the compatibility mode that leaves an object where the flow put it.
const MODERN = { ...DEFAULT_SETTINGS, compatibilityMode: 15 };

const place = (body: string, paragraphTopPt: number, bodyTopPt = 36) =>
  placeFloat({
    anchor: firstAnchor(body),
    page: LETTER,
    paragraphTopPt,
    bodyTopPt,
    marginTopPt: 36,
    resolvePart: () => null,
    theme: NO_THEME,
    settings: MODERN,
  });

describe("readAnchors", () => {
  it("reads extent, wrap and stacking order", () => {
    const anchor = firstAnchor(anchorXml({ h: offsetH(0), v: offsetV(0) }));
    expect(anchor.name).toBe("Logo");
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

  it("reads the side a wrap allows text on, and takes both where it names none", () => {
    const sideOf = (wrap: string) => firstAnchor(anchorXml({ h: offsetH(0), v: offsetV(0), wrap }));
    expect(sideOf('<wp:wrapSquare wrapText="largest"/>').side).toBe("largest");
    expect(sideOf('<wp:wrapSquare wrapText="left"/>').side).toBe("left");
    expect(sideOf("<wp:wrapSquare/>").side).toBe("bothSides");
    expect(sideOf("<wp:wrapTopAndBottom/>").side).toBe("bothSides");
  });

  it("keeps text off the whole of an object wrapped square", () => {
    const anchor = firstAnchor(
      anchorXml({ h: offsetH(0), v: offsetV(0), wrap: '<wp:wrapSquare wrapText="bothSides"/>' }),
    );
    expect(anchor.area).toStrictEqual(WHOLE_FRAME);
  });

  it("keeps text off the rectangle around a tight wrap's polygon", () => {
    const anchor = firstAnchor(
      anchorXml({ h: offsetH(0), v: offsetV(0), wrap: tightWrap(5400, 16200) }),
    );
    expect(anchor.wrap).toBe("tight");
    expect(edgesOf(anchor.area)).toStrictEqual({ left: 0.25, top: 0, right: 1, bottom: 0.75 });
  });

  it("turns that polygon over with an object that was flipped", () => {
    const anchor = firstAnchor(
      anchorXml({ h: offsetH(0), v: offsetV(0), wrap: tightWrap(5400, 16200), flip: 'flipH="1"' }),
    );
    expect(edgesOf(anchor.area)).toStrictEqual({ left: 0, top: 0, right: 0.75, bottom: 0.75 });
  });

  it("keeps every corner of that polygon, turned over with it", () => {
    const anchor = firstAnchor(
      anchorXml({ h: offsetH(0), v: offsetV(0), wrap: tightWrap(5400, 16200), flip: 'flipH="1"' }),
    );
    expect(anchor.area.corners).toStrictEqual([
      { x: 0.75, y: 0 },
      { x: 0.75, y: 0.75 },
      { x: 0, y: 0.75 },
      { x: 0, y: 0 },
      { x: 0.75, y: 0 },
    ]);
  });

  it("takes a tight wrap with no polygon as the whole frame", () => {
    const anchor = firstAnchor(
      anchorXml({ h: offsetH(0), v: offsetV(0), wrap: "<wp:wrapTight/>" }),
    );
    expect(anchor.area).toStrictEqual(WHOLE_FRAME);
  });

  it("reads the distances text is kept off a wrapping object", () => {
    const anchor = firstAnchor(anchorXml({ h: offsetH(0), v: offsetV(0), distances: true }));
    expect(anchor.distances).toStrictEqual({
      topEmu: 45720,
      rightEmu: 114300,
      bottomEmu: 45720,
      leftEmu: 114300,
    });
  });

  it("keeps text against an object that asks for no distance at all", () => {
    const anchor = firstAnchor(anchorXml({ h: offsetH(0), v: offsetV(0) }));
    expect(anchor.distances).toStrictEqual({
      topEmu: 0,
      rightEmu: 0,
      bottomEmu: 0,
      leftEmu: 0,
    });
  });

  it("distinguishes an aligned position from an offset one", () => {
    const anchor = firstAnchor(anchorXml({ h: alignH("right"), v: offsetV(0) }));
    expect(anchor.horizontal).toStrictEqual({ kind: "align", from: "margin", align: "right" });
  });
});

describe("placeFloat", () => {
  it("resolves a column offset against the left margin", () => {
    // A header logo 5199353 EMU right of the column origin.
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

  it("measures a margin-relative offset from where the body's text starts", () => {
    // Word puts this header picture, asked for 55.77pt above the margin, 13.3pt
    // down a page whose header holds the body off until 69.04pt.
    const placed = placeFloat({
      anchor: firstAnchor(anchorXml({ h: offsetH(0), v: offsetV(-708274, "margin") })),
      page: LETTER,
      paragraphTopPt: 21.6,
      bodyTopPt: 21.6,
      marginTopPt: 69.04,
      resolvePart: () => null,
      theme: NO_THEME,
      settings: MODERN,
    });
    expect(placed.topPt).toBeCloseTo(13.27, 2);
  });

  // A document declaring no compatibility mode has Word put its objects on the
  // twip grid, which is what a legacy document's wrapping turns on.
  it("rounds the place a legacy document puts an object to the whole twip", () => {
    const legacy = placeFloat({
      anchor: firstAnchor(anchorXml({ h: offsetH(0), v: offsetV(0) })),
      page: LETTER,
      paragraphTopPt: 494.636,
      bodyTopPt: 36,
      marginTopPt: 36,
      resolvePart: () => null,
      theme: NO_THEME,
    });
    expect(legacy.topPt).toBe(494.65);
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

  it("lands an aligned object on the size it turned out to be, not its stored one", () => {
    const placed = placeFloat({
      anchor: firstAnchor(anchorXml({ h: alignH("right"), v: offsetV(0) })),
      page: LETTER,
      paragraphTopPt: 100,
      bodyTopPt: 36,
      marginTopPt: 36,
      resolvePart: () => null,
      theme: NO_THEME,
      sizePt: { widthPt: 44.5, heightPt: 27.2 },
    });

    expect(placed.leftPt).toBeCloseTo(612 - 36 - 44.5, 6);
    expect(placed.widthPt).toBeCloseTo(44.5, 6);
    expect(placed.heightPt).toBeCloseTo(27.2, 6);
  });

  it("converts the extent into points", () => {
    const placed = place(anchorXml({ h: offsetH(0), v: offsetV(0) }), 100);
    expect(placed.widthPt).toBeCloseTo(180, 6);
    expect(placed.heightPt).toBeCloseTo(71.2, 3);
  });
});

const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const pictureXml = (fill: string) =>
  anchorXml({
    h: offsetH(0),
    v: offsetV(0),
  }).replace(
    "</wp:anchor>",
    `<a:graphic xmlns:a="${A_NS}"><a:graphicData><pic:pic xmlns:pic="${PIC_NS}"
       xmlns:r="${R_NS}"><pic:blipFill>${fill}</pic:blipFill></pic:pic>
     </a:graphicData></a:graphic></wp:anchor>`,
  );

const placeResolving = (body: string, resolvePart: (id: string) => string | null) =>
  placeFloat({
    anchor: firstAnchor(body),
    page: LETTER,
    paragraphTopPt: 100,
    bodyTopPt: 36,
    marginTopPt: 36,
    resolvePart,
    theme: NO_THEME,
  });

describe("placeFloat content", () => {
  it("resolves a picture to the part that holds its bytes", () => {
    const placed = placeResolving(
      pictureXml(`<a:blip r:embed="rId7"/><a:srcRect t="7272"/>`),
      (id) => (id === "rId7" ? "word/media/image1.png" : null),
    );
    expect(placed.content).toStrictEqual({
      kind: "picture",
      part: "word/media/image1.png",
      paint: UNPAINTED,
      crop: { left: 0, top: 0.07272, right: 0, bottom: 0 },
    });
  });

  it("says which relationship it could not resolve rather than dropping the picture", () => {
    const placed = placeResolving(pictureXml(`<a:blip r:embed="rId7"/>`), () => null);
    expect(placed.content).toStrictEqual({ kind: "missing-picture", relationshipId: "rId7" });
  });

  it("carries a shape through as a frame with nothing to fetch", () => {
    const placed = placeResolving(anchorXml({ h: offsetH(0), v: offsetV(0) }), () => null);
    expect(placed.content).toStrictEqual({ kind: "unknown" });
  });
});

// **A flip has to reach whatever draws the object.** It was read for the wrap
// polygon and dropped before anything painted, so `drawablesOf` handed every
// top-level object an unflipped one and both backends drew a connector between the
// wrong corners and a triangle the wrong way up. Nothing could see it but the page.
describe("a flipped object", () => {
  const flipOf = (flip: string) =>
    place(anchorXml({ h: offsetH(0), v: offsetV(0), flip }), 100).flip;

  it("carries what the drawing states through to where it is placed", () => {
    expect(flipOf('flipH="1"')).toStrictEqual({ horizontal: true, vertical: false });
    expect(flipOf('flipV="1"')).toStrictEqual({ horizontal: false, vertical: true });
    expect(flipOf('flipH="1" flipV="1"')).toStrictEqual({ horizontal: true, vertical: true });
  });

  it("says a drawing stating nothing was flipped neither way", () => {
    expect(flipOf("")).toStrictEqual({ horizontal: false, vertical: false });
  });
});
