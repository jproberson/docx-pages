import { describe, expect, it } from "vitest";

import { readDrawingContent, NO_CROP } from "./drawing.js";
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

describe("readDrawingContent", () => {
  it("reads a picture's relationship and reports it uncropped", () => {
    expect(drawing(picture(`<a:blip r:embed="rId7"/>`))).toStrictEqual({
      kind: "picture",
      relationshipId: "rId7",
      crop: NO_CROP,
    });
  });

  it("reads srcRect as fractions of the source bitmap", () => {
    const content = drawing(picture(`<a:blip r:embed="rId7"/><a:srcRect t="7272" b="7272"/>`));
    expect(content).toStrictEqual({
      kind: "picture",
      relationshipId: "rId7",
      crop: { left: 0, top: 0.07272, right: 0, bottom: 0.07272 },
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
