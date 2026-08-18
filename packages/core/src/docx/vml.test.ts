import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument } from "../testing/build-docx.js";
import { readAnchors, type FloatingAnchor } from "./anchors.js";
import { readParagraphs } from "./blocks.js";
import { DEFAULT_TEXT_INSETS } from "./drawing.js";
import { openDocx } from "./package.js";
import { V_NS } from "./vml.js";

const W10_NS = "urn:schemas-microsoft-com:office:word";

const anchorsOf = (body: string): readonly FloatingAnchor[] =>
  readParagraphs(openDocx(buildDocx({ "word/document.xml": wordDocument(body) }))).flatMap(
    (paragraph) => readAnchors(paragraph),
  );

// One container, one shape, and a paragraph of text inside the box. The corpus
// writes every one of its 70 boxes as a `v:shape` of the text box type.
const held = (shape: string, holder = "pict") =>
  `<w:p><w:r><w:${holder} xmlns:v="${V_NS}" xmlns:w10="${W10_NS}">${shape}` +
  `</w:${holder}></w:r></w:p>`;

const TEXT = `<w:p><w:r><w:t>boxed</w:t></w:r></w:p>`;

const textBox = (style: string, attributes = "", inside = TEXT, box = "") =>
  `<v:shape type="#_x0000_t202" style="${style}" ${attributes}>` +
  `<v:textbox ${box}><w:txbxContent>${inside}` +
  `</w:txbxContent></v:textbox><w10:wrap anchorx="page"/></v:shape>`;

// What the corpus's commonest box states: out of the flow, across from the page's
// own edge and down from the paragraph it is anchored in.
const PLACED = "position:absolute;margin-left:292.75pt;margin-top:14pt;width:323pt;height:129.6pt";

const EMU_PER_POINT = 12700;
const points = (emu: number) => emu / EMU_PER_POINT;

const only = (anchors: readonly FloatingAnchor[]): FloatingAnchor => {
  expect(anchors).toHaveLength(1);
  const [first] = anchors;
  if (first === undefined) throw new Error("no anchor");
  return first;
};

describe("readAnchors over a drawing in the old form", () => {
  it("reads a positioned text box into the anchor a wp:anchor would have given", () => {
    const anchor = only(
      anchorsOf(held(textBox(`${PLACED};mso-position-horizontal-relative:page`))),
    );

    expect(points(anchor.widthEmu)).toBeCloseTo(323, 6);
    expect(points(anchor.heightEmu)).toBeCloseTo(129.6, 6);
    expect(anchor.horizontal.from).toBe("page");
    expect(anchor.horizontal).toStrictEqual({
      kind: "offset",
      from: "page",
      offsetEmu: Math.round(292.75 * EMU_PER_POINT),
    });
    // VML states no vertical origin here, and its own default is the text, which is
    // the paragraph down the page.
    expect(anchor.vertical).toStrictEqual({
      kind: "offset",
      from: "paragraph",
      offsetEmu: Math.round(14 * EMU_PER_POINT),
    });
    expect(anchor.content.kind).toBe("text-box");
    expect(anchor.paragraphIndex).toBe(0);
  });

  it("reads the box's own paragraphs, which are laid out in its frame", () => {
    const two = `${TEXT}${TEXT}`;
    const anchor = only(anchorsOf(held(textBox(PLACED, "", two))));
    expect(anchor.content.kind === "text-box" && anchor.content.body.blocks).toHaveLength(2);
  });

  // Word writes `left:0;text-align:left` beside the margin that is the real
  // position, on 19 of the corpus's 70 boxes.
  it("passes over the css left, which is not where the shape stands", () => {
    const styled = `position:absolute;left:0;text-align:left;margin-left:433.85pt;
      margin-top:56.15pt;width:157.75pt;height:10pt;mso-position-horizontal-relative:page`;
    const anchor = only(anchorsOf(held(textBox(styled.replace(/\s+/g, "")))));
    expect(anchor.horizontal).toStrictEqual({
      kind: "offset",
      from: "page",
      offsetEmu: Math.round(433.85 * EMU_PER_POINT),
    });
  });

  // The nine that state an alignment state a margin off the paper beside it.
  it("takes an alignment over the offset on the same axis", () => {
    const styled = `${PLACED};mso-position-horizontal:right;mso-position-horizontal-relative:page`;
    const anchor = only(anchorsOf(held(textBox(styled))));
    expect(anchor.horizontal).toStrictEqual({ kind: "align", from: "page", align: "right" });
  });

  it("reads mso-position-horizontal:absolute as the offset it stands beside", () => {
    const styled = `${PLACED};mso-position-horizontal:absolute;mso-position-vertical:absolute`;
    expect(only(anchorsOf(held(textBox(styled)))).horizontal.kind).toBe("offset");
  });

  it("names the thing each origin is measured from in the other dialect's words", () => {
    const from = (relative: string) =>
      only(anchorsOf(held(textBox(`${PLACED};mso-position-vertical-relative:${relative}`))))
        .vertical.from;
    expect(from("page")).toBe("page");
    expect(from("margin")).toBe("margin");
    expect(from("text")).toBe("paragraph");
    expect(from("line")).toBe("line");
  });

  it("passes over a box whose origin is a word it does not know", () => {
    expect(
      anchorsOf(held(textBox(`${PLACED};mso-position-vertical-relative:elsewhere`))),
    ).toStrictEqual([]);
  });

  it("reads a length in inches, and a bare nought", () => {
    const styled = "position:absolute;margin-left:0;margin-top:14.15pt;width:3in;height:1in";
    const anchor = only(anchorsOf(held(textBox(styled))));
    expect(points(anchor.widthEmu)).toBeCloseTo(216, 6);
    expect(anchor.horizontal).toStrictEqual({ kind: "offset", from: "column", offsetEmu: 0 });
  });

  // Word stacks its drawings by the size of the z-index and puts the ones below
  // nought behind the text.
  it("reads the z-index as the depth and which side of the text it is drawn on", () => {
    const behind = only(anchorsOf(held(textBox(`${PLACED};z-index:-251657216`))));
    expect(behind.behindDoc).toBe(true);
    expect(behind.relativeHeight).toBe(251657216);

    const inFront = only(anchorsOf(held(textBox(`${PLACED};z-index:251658240`))));
    expect(inFront.behindDoc).toBe(false);
    expect(inFront.relativeHeight).toBe(251658240);
  });

  // Nothing measured says what Word does with `w10:wrap`, and 60 of the corpus's 70
  // boxes sit behind the text where it wraps nothing, so no line is moved for one.
  it("wraps nothing, whatever the shape says about wrapping", () => {
    const anchor = only(anchorsOf(held(textBox(`${PLACED};mso-wrap-distance-left:9pt`))));
    expect(anchor.wrap).toBe("none");
    expect(anchor.distances.leftEmu).toBe(Math.round(9 * EMU_PER_POINT));
  });

  it("reads the same box in a w:object as in a w:pict", () => {
    expect(only(anchorsOf(held(textBox(PLACED), "object"))).content.kind).toBe("text-box");
  });

  describe("what the box does with its own text", () => {
    const bodyOf = (style: string, attributes = "", box = "") => {
      const { content } = only(anchorsOf(held(textBox(style, attributes, TEXT, box))));
      if (content.kind !== "text-box") throw new Error("not a text box");
      return content;
    };

    it("holds it off the walls by Word's own default where the box states none", () => {
      expect(bodyOf(PLACED).body.insets).toStrictEqual(DEFAULT_TEXT_INSETS);
    });

    it("reads the inset the box states, left, top, right and bottom", () => {
      expect(bodyOf(PLACED, "", `inset="0,0,20pt,0"`).body.insets).toStrictEqual({
        leftEmu: 0,
        topEmu: 0,
        rightEmu: Math.round(20 * EMU_PER_POINT),
        bottomEmu: 0,
      });
    });

    it("reads where the text sits in the box, and wraps it unless told not to", () => {
      expect(bodyOf(`${PLACED};v-text-anchor:bottom`).body.anchor).toBe("bottom");
      expect(bodyOf(`${PLACED};v-text-anchor:middle`).body.anchor).toBe("center");
      expect(bodyOf(PLACED).body.anchor).toBe("top");
      expect(bodyOf(PLACED).body.wraps).toBe(true);
      expect(bodyOf(`${PLACED};mso-wrap-style:none`).body.wraps).toBe(false);
    });

    it("reads a box that fits itself to its text, which states it on itself", () => {
      expect(bodyOf(PLACED).body.fitsText).toBe(false);
      expect(bodyOf(PLACED, "", `style="mso-fit-shape-to-text:t"`).body.fitsText).toBe(true);
    });

    // A VML colour carries the palette entry Word remembers beside it.
    it("reads the fill and the outline, with the palette entry left out of the colour", () => {
      const painted = bodyOf(PLACED, `fillcolor="#f2f2f2 [3052]" stroked="f"`).paint;
      expect(painted.fill).toStrictEqual({
        base: { kind: "literal", hex: "f2f2f2" },
        luminanceScale: 1,
        luminanceOffset: 0,
      });
      expect(painted.outline).toBeNull();
      expect(bodyOf(PLACED, `filled="f" stroked="f"`).paint.fill).toBeNull();
    });

    // Unexercised by the corpus, where every box turns the stroke off, and it is
    // what VML says a shape saying nothing is drawn as.
    it("fills a shape white and strokes it black where it says neither", () => {
      const painted = bodyOf(PLACED).paint;
      expect(painted.fill?.base).toStrictEqual({ kind: "literal", hex: "ffffff" });
      expect(painted.outline?.color?.base).toStrictEqual({ kind: "literal", hex: "000000" });
      expect(painted.outline?.widthPt).toBeCloseTo(0.75, 6);
      expect(painted.outline?.widthStated).toBe(false);
    });

    it("reads a stroke the shape states, and says the width was stated", () => {
      const painted = bodyOf(PLACED, `strokeweight=".5pt" strokecolor="#ff0000"`).paint;
      expect(painted.outline?.widthPt).toBeCloseTo(0.5, 6);
      expect(painted.outline?.widthStated).toBe(true);
      expect(painted.outline?.color?.base).toStrictEqual({ kind: "literal", hex: "ff0000" });
    });

    it("leaves a colour it cannot read unpainted rather than putting the default down", () => {
      expect(bodyOf(PLACED, `fillcolor="window"`).paint.fill).toBeNull();
    });
  });

  describe("what it refuses, so that the report goes on naming it", () => {
    it("passes over a box standing in the line rather than out of it", () => {
      expect(anchorsOf(held(textBox("width:323pt;height:129.6pt")))).toStrictEqual([]);
    });

    // 26 items in eight corpus documents state a share beside a size in points, and
    // which of the two Word draws is unmeasured.
    it("passes over a box sized as a share of something else", () => {
      expect(anchorsOf(held(textBox(`${PLACED};mso-width-percent:400`)))).toStrictEqual([]);
      expect(anchorsOf(held(textBox(`${PLACED};mso-height-percent:200`)))).toStrictEqual([]);
      expect(anchorsOf(held(textBox(`${PLACED};mso-width-percent:0`)))).toHaveLength(1);
    });

    it("passes over a box whose size or position it cannot read", () => {
      const noSize = "position:absolute;margin-left:0;margin-top:0;height:10pt";
      const noOffset = "position:absolute;width:10pt;height:10pt";
      const strangeUnit = "position:absolute;margin-left:0;margin-top:0;width:10em;height:10pt";
      expect(anchorsOf(held(textBox(noSize)))).toStrictEqual([]);
      expect(anchorsOf(held(textBox(noOffset)))).toStrictEqual([]);
      expect(anchorsOf(held(textBox(strangeUnit)))).toStrictEqual([]);
    });

    it("passes over a shape that is not a text box at all", () => {
      const line = `<v:line style="position:absolute" from="0,0" to="180pt,0"/>`;
      expect(anchorsOf(held(line))).toStrictEqual([]);
    });

    // The copy Word itself ignores. `paragraphOwnDrawings` drops it before this is
    // asked, which is what keeps the fallback of a DrawingML box from being read a
    // second time.
    it("passes over the fallback twin of a shape read from mc:Choice", () => {
      const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
      const body = `<w:p><w:r><mc:AlternateContent xmlns:mc="${MC_NS}">
        <mc:Fallback><w:pict xmlns:v="${V_NS}" xmlns:w10="${W10_NS}">${textBox(PLACED)}</w:pict>
        </mc:Fallback></mc:AlternateContent></w:r></w:p>`;
      expect(anchorsOf(body)).toStrictEqual([]);
    });
  });

  // A group states its children in a space of its own, `coordsize` and
  // `coordorigin`, which is `a:chExt` and `a:chOff` under another name.
  describe("a group", () => {
    const group = (children: string, attributes = `coordorigin=",-1244" coordsize="12240,1424"`) =>
      held(
        `<v:group style="position:absolute;margin-left:0;margin-top:-62.2pt;width:612pt;height:71.2pt"
           ${attributes}>${children}</v:group>`,
      );

    const inGroup = (style: string) =>
      `<v:shape type="#_x0000_t202" style="${style}"><v:textbox><w:txbxContent>${TEXT}` +
      `</w:txbxContent></v:textbox></v:shape>`;

    it("keeps each child where the group's own units put it", () => {
      const anchor = only(
        anchorsOf(group(inGroup("position:absolute;top:-1245;width:12240;height:1424"))),
      );
      expect(points(anchor.widthEmu)).toBeCloseTo(612, 6);
      if (anchor.content.kind !== "group") throw new Error("not a group");

      const [child] = anchor.content.children;
      expect(child?.leftFraction).toBeCloseTo(0, 6);
      expect(child?.topFraction).toBeCloseTo(-1 / 1424, 6);
      expect(child?.widthFraction).toBeCloseTo(1, 6);
      expect(child?.heightFraction).toBeCloseTo(1, 6);
      expect(child?.content.kind).toBe("text-box");
    });

    it("measures a child from the group's own origin", () => {
      const anchor = only(
        anchorsOf(group(inGroup("position:absolute;left:6120;top:-532;width:6120;height:712"))),
      );
      if (anchor.content.kind !== "group") throw new Error("not a group");
      const [child] = anchor.content.children;
      expect(child?.leftFraction).toBeCloseTo(0.5, 6);
      expect(child?.topFraction).toBeCloseTo(712 / 1424, 6);
      expect(child?.widthFraction).toBeCloseTo(0.5, 6);
    });

    // **A group half read is still drawn.** A box in the right place is worth having
    // whether or not the rule beside it could be read, and what was left out is what
    // the fidelity report goes on naming.
    it("draws the children it can read and leaves the rest out", () => {
      const rule = `<v:rect style="position:absolute;top:-1;width:12240;height:15" fillcolor="#bebebe"/>`;
      const anchor = only(
        anchorsOf(group(rule + inGroup("position:absolute;top:-1245;width:12240;height:1424"))),
      );
      if (anchor.content.kind !== "group") throw new Error("not a group");
      expect(anchor.content.children).toHaveLength(1);
    });

    it("reads a group inside a group, each level in its own space", () => {
      const inner =
        `<v:group style="position:absolute;left:0;top:-1244;width:6120;height:1424"` +
        ` coordorigin="0,0" coordsize="1000,1000">` +
        `${inGroup("position:absolute;left:500;top:0;width:500;height:1000")}</v:group>`;
      const anchor = only(anchorsOf(group(inner)));
      if (anchor.content.kind !== "group") throw new Error("not a group");

      const [child] = anchor.content.children;
      expect(child?.widthFraction).toBeCloseTo(0.5, 6);
      if (child?.content.kind !== "group") throw new Error("not a group inside a group");
      const [inside] = child.content.children;
      expect(inside?.leftFraction).toBeCloseTo(0.5, 6);
      expect(inside?.content.kind).toBe("text-box");
    });

    it("passes over a group that states no space of its own", () => {
      const child = inGroup("position:absolute;top:0;width:100;height:100");
      expect(anchorsOf(group(child, ""))).toStrictEqual([]);
      expect(anchorsOf(group(child, `coordsize="0,0"`))).toStrictEqual([]);
    });

    it("passes over a group holding nothing it can read", () => {
      const rule = `<v:rect style="position:absolute;top:-1;width:12240;height:15"/>`;
      expect(anchorsOf(group(rule))).toStrictEqual([]);
    });
  });
});
