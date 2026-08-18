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
        widthStated: false,
      },
      geometry: "rectangle",
      path: null,
    });
  });

  // A path a file draws point by point, kept as shares of the shape's own box so
  // that a group scaling its children scales the outline with them. The space the
  // points are in is the path's own where it names one and the shape's extent where
  // it does not, which 34 of the corpus's 332 custom geometries rely on.
  describe("a path the file draws point by point", () => {
    const custom = (pathList: string, extent = `<a:xfrm><a:ext cx="1000" cy="500"/></a:xfrm>`) =>
      shapePaint(`${extent}<a:custGeom><a:avLst/><a:pathLst>${pathList}</a:pathLst></a:custGeom>`);

    it("reads a triangle out of the space the path states", () => {
      const paint = custom(
        `<a:path w="200" h="100">
           <a:moveTo><a:pt x="100" y="0"/></a:moveTo>
           <a:lnTo><a:pt x="200" y="100"/></a:lnTo>
           <a:lnTo><a:pt x="0" y="100"/></a:lnTo>
           <a:close/>
         </a:path>`,
      );
      expect(paint.geometry).toBe("custom");
      expect(paint.path).toStrictEqual([
        { kind: "move", to: { x: 0.5, y: 0 } },
        { kind: "line", to: { x: 1, y: 1 } },
        { kind: "line", to: { x: 0, y: 1 } },
        { kind: "close" },
      ]);
    });

    it("takes the shape's own extent where the path states no space", () => {
      expect(
        custom(`<a:path><a:moveTo><a:pt x="500" y="250"/></a:moveTo></a:path>`).path,
      ).toStrictEqual([{ kind: "move", to: { x: 0.5, y: 0.5 } }]);
    });

    it("reads the curve, which four corpus documents draw with", () => {
      expect(
        custom(
          `<a:path w="100" h="100"><a:cubicBezTo>
             <a:pt x="0" y="50"/><a:pt x="50" y="100"/><a:pt x="100" y="100"/>
           </a:cubicBezTo></a:path>`,
        ).path,
      ).toStrictEqual([
        {
          kind: "curve",
          first: { x: 0, y: 0.5 },
          second: { x: 0.5, y: 1 },
          to: { x: 1, y: 1 },
        },
      ]);
    });

    it("holds every subpath of a path list, one after another", () => {
      const paint = custom(
        `<a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:close/></a:path>
         <a:path w="100" h="100"><a:moveTo><a:pt x="100" y="100"/></a:moveTo><a:close/></a:path>`,
      );
      expect(paint.path?.length).toBe(4);
    });

    // **A command this cannot play refuses the whole path.** A shape missing one of
    // its sides is a wrong drawing; a shape missing altogether is the gap the report
    // already names, and neither an arc nor a quadratic appears in the corpus.
    it("refuses a path holding an arc rather than dropping the arc", () => {
      const paint = custom(
        `<a:path w="100" h="100">
           <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
           <a:arcTo wR="50" hR="50" stAng="0" swAng="5400000"/>
         </a:path>`,
      );
      expect(paint.geometry).toBe("custom");
      expect(paint.path).toBeNull();
    });

    it("refuses a path whose space is nothing at all", () => {
      expect(
        custom(`<a:path w="0" h="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo></a:path>`, "").path,
      ).toBeNull();
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

  // A box that fits itself to its text grows by a width the file states and by
  // nothing for one it does not, so the two have to be told apart.
  it("says whether the file stated the outline's width or Word's hairline stood in", () => {
    const stated = (attributes: string) =>
      shapePaint(`<a:ln ${attributes}><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`)
        .outline?.widthStated;
    expect(stated(``)).toBe(false);
    expect(stated(`w="25400"`)).toBe(true);
  });

  it("reads an outline that paints nothing as one that takes room and no colour", () => {
    const found = shapePaint(`<a:ln w="9525"><a:noFill/></a:ln>`).outline;
    expect(found?.widthPt).toBeCloseTo(0.75, 5);
    expect(found?.color).toBeNull();
  });

  // An outline filled with a colour stated at no opacity is the other way a file
  // writes a line it does not want drawn, and it still takes its width.
  it("reads an outline stated fully transparent the same as one that paints nothing", () => {
    const found = shapePaint(
      `<a:ln w="1778"><a:solidFill><a:srgbClr val="000000"><a:alpha val="0"/></a:srgbClr>
       </a:solidFill></a:ln>`,
    ).outline;
    expect(found?.widthPt).toBeCloseTo(0.14, 2);
    expect(found?.color).toBeNull();
  });

  it("reads no fill from a shape whose fill is stated fully transparent", () => {
    expect(
      shapePaint(`<a:solidFill><a:srgbClr val="4472C4"><a:alpha val="0"/></a:srgbClr>
        </a:solidFill>`).fill,
    ).toBeNull();
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

const WPG_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

// A group standing in a space of its own, with whatever children are given.
const group = (space: string, children: string) =>
  drawing(
    `<wpg:wgp xmlns:wpg="${WPG_NS}"><wpg:grpSpPr><a:xfrm>` +
      `<a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/>${space}` +
      `</a:xfrm></wpg:grpSpPr>${children}</wpg:wgp>`,
  );

const shapeIn = (transform: string, spPr = "") =>
  `<wps:wsp xmlns:wps="${WPS_NS}"><wps:spPr><a:xfrm>${transform}</a:xfrm>${spPr}</wps:spPr></wps:wsp>`;

const childrenOf = (content: ReturnType<typeof drawing>) => {
  if (content.kind !== "group") throw new Error("expected a group");
  return content.children;
};

describe("a group of shapes", () => {
  // **A picture inside a group is not what the group is.** Asking for the picture
  // first is what drew a photograph's cropped middle where Word draws a diagram of
  // 323 shapes, with the other 322 drawn nowhere at all.
  it("is read as the group rather than as the first picture inside it", () => {
    const inside =
      `<pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="500" cy="500"/></a:xfrm></pic:spPr></pic:pic>`;
    const content = group(
      "",
      `${inside}${shapeIn(`<a:off x="500" y="0"/><a:ext cx="500" cy="500"/>`)}`,
    );

    expect(content.kind).toBe("group");
    expect(childrenOf(content).map((each) => each.content.kind)).toStrictEqual([
      "picture",
      "shape",
    ]);
  });

  it("places a child as the fraction of the group's box it stands in", () => {
    const content = group(
      `<a:chOff x="0" y="0"/><a:chExt cx="1000" cy="500"/>`,
      shapeIn(`<a:off x="250" y="250"/><a:ext cx="500" cy="125"/>`),
    );

    expect(childrenOf(content)[0]).toMatchObject({
      leftFraction: 0.25,
      topFraction: 0.5,
      widthFraction: 0.5,
      heightFraction: 0.25,
    });
  });

  // The child space has an origin of its own as well as a size, and Word writes one
  // that is nowhere near the page for a group anything has been moved inside.
  it("takes the child space's own origin off first", () => {
    const content = group(
      `<a:chOff x="100" y="200"/><a:chExt cx="1000" cy="1000"/>`,
      shapeIn(`<a:off x="600" y="200"/><a:ext cx="100" cy="100"/>`),
    );

    expect(childrenOf(content)[0]).toMatchObject({ leftFraction: 0.5, topFraction: 0 });
  });

  // Word writes no child space for a group nothing has been scaled inside, and then
  // the group's own extent is the space its children stand in.
  it("stands the children in the group's own extent where it states no child space", () => {
    const content = group("", shapeIn(`<a:off x="500" y="0"/><a:ext cx="500" cy="1000"/>`));

    expect(childrenOf(content)[0]).toMatchObject({ leftFraction: 0.5, widthFraction: 0.5 });
  });

  it("reads a group inside a group as a group again", () => {
    const inner =
      `<wpg:grpSp xmlns:wpg="${WPG_NS}"><wpg:grpSpPr><a:xfrm>` +
      `<a:off x="0" y="0"/><a:ext cx="500" cy="1000"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></wpg:grpSpPr>` +
      shapeIn(`<a:off x="50" y="0"/><a:ext cx="50" cy="100"/>`) +
      `</wpg:grpSp>`;
    const nested = childrenOf(group("", inner))[0];

    expect(nested).toMatchObject({ leftFraction: 0, widthFraction: 0.5 });
    expect(childrenOf(nested?.content ?? { kind: "unknown" })[0]).toMatchObject({
      leftFraction: 0.5,
      widthFraction: 0.5,
    });
  });

  it("keeps which way round a child was turned", () => {
    const content = group(
      "",
      shapeIn(`<a:off x="0" y="0"/><a:ext cx="10" cy="10"/>`).replace(
        "<a:xfrm>",
        `<a:xfrm flipV="1">`,
      ),
    );

    expect(childrenOf(content)[0]?.flip).toStrictEqual({ horizontal: false, vertical: true });
  });
});

describe("the preset geometry a shape names", () => {
  it("reads the ones the corpus draws", () => {
    const geometryOf = (preset: string) => shapePaint(`<a:prstGeom prst="${preset}"/>`).geometry;

    expect(geometryOf("ellipse")).toBe("ellipse");
    expect(geometryOf("roundRect")).toBe("rounded-rectangle");
    expect(geometryOf("triangle")).toBe("triangle");
    expect(geometryOf("rect")).toBe("rectangle");
  });

  // A connector is stored as a box with a line across it, exactly as `line` is.
  it("draws a straight connector as the line it is", () => {
    expect(shapePaint(`<a:prstGeom prst="straightConnector1"/>`).geometry).toBe("line");
  });

  it("draws a preset it does not know as the rectangle everything used to be", () => {
    expect(shapePaint(`<a:prstGeom prst="cloudCallout"/>`).geometry).toBe("rectangle");
  });
});
