import { describe, expect, it } from "vitest";

import {
  readDrawingContent,
  readDrawingFlip,
  DEFAULT_TEXT_INSETS,
  NO_CROP,
  NO_PAINT,
} from "./drawing.js";
import { parseXml } from "./xml.js";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const drawing = (inner: string) => {
  const root = parseXml(
    `<wp:anchor xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}"
       xmlns:wps="${WPS_NS}" xmlns:r="${R_NS}" xmlns:w="${W_NS}">
       <a:graphic><a:graphicData>${inner}</a:graphicData></a:graphic></wp:anchor>`,
  );
  if (root === null) throw new Error("unparsed");
  return readDrawingContent(root);
};

const picture = (fill: string) => `<pic:pic><pic:blipFill>${fill}</pic:blipFill></pic:pic>`;

const flipOf = (transform: string) => {
  const root = parseXml(
    `<wp:anchor xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}">
       <a:graphic><a:graphicData><pic:pic><pic:spPr>${transform}</pic:spPr></pic:pic>
       </a:graphicData></a:graphic></wp:anchor>`,
  );
  if (root === null) throw new Error("unparsed");
  return readDrawingFlip(root);
};

describe("readDrawingFlip", () => {
  it("reads an object left the way it was drawn", () => {
    expect(flipOf(`<a:xfrm/>`)).toStrictEqual({ horizontal: false, vertical: false });
    expect(flipOf(``)).toStrictEqual({ horizontal: false, vertical: false });
  });

  it("reads an object turned over", () => {
    expect(flipOf(`<a:xfrm flipH="1"/>`)).toStrictEqual({ horizontal: true, vertical: false });
    expect(flipOf(`<a:xfrm flipV="1"/>`)).toStrictEqual({ horizontal: false, vertical: true });
  });
});

describe("readDrawingContent", () => {
  it("reads a picture's relationship and reports it uncropped", () => {
    expect(drawing(picture(`<a:blip r:embed="rId7"/>`))).toStrictEqual({
      kind: "picture",
      relationshipId: "rId7",
      crop: NO_CROP,
      paint: NO_PAINT,
    });
  });

  it("reads srcRect as fractions of the source bitmap", () => {
    const content = drawing(picture(`<a:blip r:embed="rId7"/><a:srcRect t="7272" b="7272"/>`));
    expect(content).toStrictEqual({
      kind: "picture",
      relationshipId: "rId7",
      crop: { left: 0, top: 0.07272, right: 0, bottom: 0.07272 },
      paint: NO_PAINT,
    });
  });

  it("treats a picture with no relationship as unresolvable rather than guessing", () => {
    expect(drawing(picture(`<a:blip/>`)).kind).toBe("unknown");
  });

  it("tells a text box apart from a plain shape", () => {
    expect(drawing(`<wps:wsp><wps:txbx><w:txbxContent/></wps:txbx></wps:wsp>`).kind).toBe(
      "text-box",
    );
    expect(drawing(`<wps:wsp><wps:spPr/></wps:wsp>`).kind).toBe("shape");
  });

  it("does not mistake a picture inside a text box for the text box's own content", () => {
    const inner = `<wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:drawing>
      ${picture(`<a:blip r:embed="rId9"/>`)}
    </w:drawing></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp>`;
    expect(drawing(inner).kind).toBe("text-box");
  });

  it("reports content it does not recognise instead of failing", () => {
    expect(drawing(`<a:chart/>`).kind).toBe("unknown");
  });
});

const textBox = (bodyPr: string, content = "<w:p/>") =>
  drawing(`<wps:wsp><wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>
    ${bodyPr}</wps:wsp>`);

const bodyOf = (bodyPr: string, content?: string) => {
  const found = textBox(bodyPr, content);
  if (found.kind !== "text-box") throw new Error(found.kind);
  return found.body;
};

describe("a text box's body", () => {
  it("reads the paragraphs it holds, numbered within the box", () => {
    const body = bodyOf(``, `<w:p/><w:p/>`);

    expect(
      body.blocks.map((block) => block.kind === "paragraph" && block.paragraph.index),
    ).toStrictEqual([0, 1]);
  });

  it("takes Word's own insets when the shape states none", () => {
    expect(bodyOf(``).insets).toStrictEqual(DEFAULT_TEXT_INSETS);
  });

  it("reads the insets the shape does state, including a zero one", () => {
    const bodyPr = `<wps:bodyPr lIns="0" tIns="45720" rIns="0" bIns="12700"/>`;

    expect(bodyOf(bodyPr).insets).toStrictEqual({
      leftEmu: 0,
      topEmu: 45720,
      rightEmu: 0,
      bottomEmu: 12700,
    });
  });

  it("reads where the text sits when the box is taller than the text", () => {
    expect(bodyOf(``).anchor).toBe("top");
    expect(bodyOf(`<wps:bodyPr anchor="ctr"/>`).anchor).toBe("center");
    expect(bodyOf(`<wps:bodyPr anchor="b"/>`).anchor).toBe("bottom");
  });

  it("reads a box that refuses to wrap its text", () => {
    expect(bodyOf(``).wraps).toBe(true);
    expect(bodyOf(`<wps:bodyPr wrap="none"/>`).wraps).toBe(false);
  });

  it("reads a box that fits itself to the text it holds", () => {
    expect(bodyOf(``).fitsText).toBe(false);
    expect(bodyOf(`<wps:bodyPr><a:spAutoFit/></wps:bodyPr>`).fitsText).toBe(true);
  });

  it("reads a box whose content is missing as an empty one", () => {
    const found = drawing(`<wps:wsp><wps:txbx/></wps:wsp>`);
    expect(found.kind === "text-box" && found.body.blocks).toStrictEqual([]);
  });
});

const shapePaint = (spPr: string) => {
  const found = drawing(`<wps:wsp><wps:spPr>${spPr}</wps:spPr></wps:wsp>`);
  if (found.kind !== "shape") throw new Error("expected a shape");
  return found.paint;
};

describe("the paint a shape carries", () => {
  it("reads the fill a shape is given, and none from one that declines it", () => {
    expect(
      shapePaint(`<a:solidFill><a:schemeClr val="bg1"><a:lumMod val="95000"/></a:schemeClr>
        </a:solidFill>`).fill,
    ).toStrictEqual({
      base: { kind: "scheme", slot: "bg1" },
      luminanceScale: 0.95,
      luminanceOffset: 0,
    });
    expect(shapePaint(`<a:noFill/>`).fill).toBeNull();
    expect(shapePaint(``)).toStrictEqual(NO_PAINT);
  });

  // A fill inside the outline is the outline's own colour, and saying so is the
  // difference between a hairline rule and a panel of solid grey.
  it("does not take the outline's colour for the shape's fill", () => {
    expect(
      shapePaint(`<a:noFill/><a:ln><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln>`),
    ).toStrictEqual({
      fill: null,
      outline: {
        color: { base: { kind: "literal", hex: "BFBFBF" }, luminanceScale: 1, luminanceOffset: 0 },
        widthPt: 0.75,
      },
      geometry: "rectangle",
    });
  });

  // Word leaves the width off an outline it draws at its own default, which its
  // pdf lays down as three quarters of a point.
  it("measures an outline, falling back on the width Word draws one at", () => {
    const outlined = (attributes: string) =>
      shapePaint(`<a:ln ${attributes}><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`)
        .outline?.widthPt;
    expect(outlined(``)).toBe(0.75);
    expect(outlined(`w="25400"`)).toBe(2);
  });

  it("reads an outline that paints nothing as one that takes room and no colour", () => {
    const found = shapePaint(`<a:ln w="9525"><a:noFill/></a:ln>`).outline;
    expect(found?.widthPt).toBeCloseTo(0.75, 5);
    expect(found?.color).toBeNull();
  });

  it("tells a line apart from every other preset, which are all rectangles here", () => {
    expect(shapePaint(`<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`).geometry).toBe("line");
    expect(shapePaint(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`).geometry).toBe(
      "rectangle",
    );
  });

  it("reads a picture's own outline, which is what frames one in these files", () => {
    const found = drawing(
      `<pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill>
        <pic:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF">
          <a:lumMod val="50000"/></a:srgbClr></a:solidFill></a:ln></pic:spPr></pic:pic>`,
    );
    expect(found.kind === "picture" && found.paint.outline?.color?.luminanceScale).toBe(0.5);
  });
});
