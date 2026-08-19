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

  /**
   * **A width stated as a share of the text frame wins over the length beside it.**
   * Measured against Word on 2026-08-19: a box stating `width:150pt` beside
   * `mso-width-percent:400;mso-width-relative:margin`, standing in a 540pt frame,
   * broke its own eighteen words at 439.7, 428.5 and 411.2 from a left of 236.4,
   * which is 216pt of room and 40.0% of the frame. The share reaches the layout,
   * since the frame it is a share of is the section's.
   */
  it("carries a width stated as a share of the frame beside the length it states", () => {
    const shared = `${PLACED};mso-width-percent:400;mso-width-relative:margin`;
    const anchor = only(anchorsOf(held(textBox(shared))));

    expect(anchor.frameWidthShare).toBeCloseTo(0.4, 6);
    expect(points(anchor.widthEmu)).toBeCloseTo(323, 6);
  });

  // Word writes the declaration out as a nought for a shape stating no share at all,
  // which is what nearly every shape in the corpus does.
  it("takes a share of nought as no share", () => {
    const none = `${PLACED};mso-width-percent:0;mso-width-relative:margin`;
    expect(only(anchorsOf(held(textBox(none)))).frameWidthShare).toBeUndefined();
  });

  // `mso-height-percent:200` stands beside the width's share on all nine corpus
  // shapes and what Word makes of it has never been asked, so the height in points
  // is what is drawn. It costs nothing: all nine fit their shape to their text.
  it("leaves a share stated for the height alone", () => {
    const shared = `${PLACED};mso-height-percent:200;mso-height-relative:margin`;
    const anchor = only(anchorsOf(held(textBox(shared))));

    expect(points(anchor.heightEmu)).toBeCloseTo(129.6, 6);
    expect(anchor.frameWidthShare).toBeUndefined();
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

  /**
   * **A `v:line` states no offset and no size**: its two ends are where it stands,
   * in the same space a `margin-left` is measured in, and the box it is drawn in is
   * the one they span. All nine in the corpus are written exactly so, which is why
   * every one of them was drawn nowhere until 2026-08-19.
   */
  describe("a line out of the flow", () => {
    const line = (attributes: string) =>
      `<v:line style="position:absolute;mso-position-horizontal:absolute;` +
      `mso-position-horizontal-relative:text;mso-position-vertical:absolute;` +
      `mso-position-vertical-relative:text" ${attributes}/>`;

    it("stands in the box its two ends span, and is stroked rather than filled in", () => {
      const anchor = only(anchorsOf(held(line(`from="134pt,128pt" to="430pt,131pt"`))));

      expect(points(anchor.widthEmu)).toBeCloseTo(296, 6);
      expect(points(anchor.heightEmu)).toBeCloseTo(3, 6);
      expect(anchor.horizontal).toStrictEqual({
        kind: "offset",
        from: "column",
        offsetEmu: Math.round(134 * EMU_PER_POINT),
      });
      expect(anchor.vertical).toStrictEqual({
        kind: "offset",
        from: "paragraph",
        offsetEmu: Math.round(128 * EMU_PER_POINT),
      });
      if (anchor.content.kind !== "shape") throw new Error("not a shape");
      expect(anchor.content.paint.geometry).toBe("line");
      expect(anchor.content.paint.outline?.color?.base).toStrictEqual({
        kind: "literal",
        hex: "000000",
      });
    });

    // Which of the box's two diagonals the line runs along is the order of its own
    // ends, turned over again by whatever the style's `flip` says. Eight of the nine
    // in the corpus state `flip:y` beside coordinates that already run down, and
    // what Word makes of the two together cannot be seen on any of them: every one
    // spans 0.6pt or less down over 300pt or more across.
    it("runs along the diagonal its ends put it on", () => {
      const down = only(anchorsOf(held(line(`from="0,0" to="90pt,30pt"`))));
      const up = only(anchorsOf(held(line(`from="0,30pt" to="90pt,0"`))));
      const turned = `from="0,0" to="90pt,30pt" `;

      expect(down.flip).toStrictEqual({ horizontal: false, vertical: false });
      expect(up.flip).toStrictEqual({ horizontal: false, vertical: true });
      expect(up.vertical).toStrictEqual({ kind: "offset", from: "paragraph", offsetEmu: 0 });
      expect(
        only(anchorsOf(held(line(turned).replace("position:absolute", "position:absolute;flip:y"))))
          .flip,
      ).toStrictEqual({ horizontal: false, vertical: true });
    });

    it("is refused where either end is stated in a unit this cannot read", () => {
      expect(anchorsOf(held(line(`from="0,0" to="90em,30pt"`)))).toStrictEqual([]);
      expect(anchorsOf(held(line(`from="0,0"`)))).toStrictEqual([]);
    });
  });

  describe("what it refuses, so that the report goes on naming it", () => {
    it("passes over a box standing in the line rather than out of it", () => {
      expect(anchorsOf(held(textBox("width:323pt;height:129.6pt")))).toStrictEqual([]);
    });

    // Only a share of the margin is measured, and the nine corpus shapes that state
    // one all state that. A share of the page is refused rather than guessed at.
    it("passes over a box whose width is a share of something other than the margin", () => {
      const ofThePage = `${PLACED};mso-width-percent:400;mso-width-relative:page`;
      expect(anchorsOf(held(textBox(ofThePage)))).toStrictEqual([]);
      expect(anchorsOf(held(textBox(`${PLACED};mso-width-percent:400`)))).toStrictEqual([]);
    });

    it("passes over a box whose size or position it cannot read", () => {
      const noSize = "position:absolute;margin-left:0;margin-top:0;height:10pt";
      const noOffset = "position:absolute;width:10pt;height:10pt";
      const strangeUnit = "position:absolute;margin-left:0;margin-top:0;width:10em;height:10pt";
      expect(anchorsOf(held(textBox(noSize)))).toStrictEqual([]);
      expect(anchorsOf(held(textBox(noOffset)))).toStrictEqual([]);
      expect(anchorsOf(held(textBox(strangeUnit)))).toStrictEqual([]);
    });

    // The line left this list on 2026-08-19 and is drawn where its ends put it. What
    // is still passed over is a shape holding neither text nor a picture whose
    // geometry its own element name does not state: what a `v:shape` draws is the
    // `v:shapetype` it names, which is a reading of its own.
    it("passes over a shape that is neither a text box nor a geometry it can name", () => {
      const plain = `<v:shape type="#_x0000_t202" style="${PLACED}"/>`;
      expect(anchorsOf(held(plain))).toStrictEqual([]);
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

    // A group's own children are where the corpus keeps most of its ovals and its
    // rounded rectangles, and a child holding no text was read as nothing at all
    // until 2026-08-19.
    it("reads a child that is paint and a geometry rather than a text box", () => {
      const round = `<v:roundrect style="position:absolute;left:0;top:0;width:12240;height:1424"
         fillcolor="#f1c232" stroked="f"/>`;
      const anchor = only(anchorsOf(group(round)));
      if (anchor.content.kind !== "group") throw new Error("not a group");

      const [child] = anchor.content.children;
      if (child?.content.kind !== "shape") throw new Error("not a shape");
      expect(child.content.paint.geometry).toBe("rounded-rectangle");
      expect(child.content.paint.fill?.base).toStrictEqual({ kind: "literal", hex: "f1c232" });
      expect(child.content.paint.outline).toBeNull();
    });

    // **A shape the style turns is left undrawn rather than drawn straight.** Three
    // corpus documents state a `rotation` on 43 shapes between them, in fractions of
    // a degree, and nothing here reads one.
    it("leaves a child the style turns out of its own box undrawn", () => {
      const turned = `<v:roundrect style="position:absolute;left:0;top:0;width:12240;height:1424;
         rotation:1752415fd" fillcolor="#f1c232"/>`;
      expect(anchorsOf(group(turned))).toStrictEqual([]);
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
      // A `v:shape` draws whatever `v:shapetype` it names, which is unread, so it is
      // the child a group can hold and this cannot draw.
      const unread = `<v:shape type="#_x0000_t32" style="position:absolute;top:-1;width:12240;height:15"/>`;
      const anchor = only(
        anchorsOf(group(unread + inGroup("position:absolute;top:-1245;width:12240;height:1424"))),
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
      const unread = `<v:shape type="#_x0000_t32" style="position:absolute;top:-1;width:12240;height:15"/>`;
      expect(anchorsOf(group(unread))).toStrictEqual([]);
    });
  });
});
