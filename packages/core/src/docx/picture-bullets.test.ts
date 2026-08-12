import { describe, expect, it } from "vitest";

import { drawablesOf } from "../layout/drawables.js";
import { bestEffortMetrics, type FaceDefaults } from "../layout/best-effort.js";
import { layOutDocument, type LaidOutDocument } from "../layout/document.js";
import { buildDocx, WORDPROCESSING_NS } from "../testing/build-docx.js";
import { buildFace } from "../testing/build-font.js";
import { openDocx } from "./package.js";
import { pictureBulletOf } from "./picture-bullets.js";
import { V_NS } from "./vml.js";

// What Word draws for a `w:numPicBullet`, measured on 2026-08-12 and written up in
// `picture-bullets.ts`: the picture once, at the head of the first paragraph of the
// body, standing on that line's baseline with the paragraph's own text after it.

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

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

// Letter, half-inch margins, so the text runs from 36pt across and 36pt down.
const SECTION =
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"` +
  ` w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

const bullet = (style: string, crop = ""): string =>
  `<w:numPicBullet w:numPicBulletId="0"><w:pict>` +
  `<v:shape id="Picture 0" type="#_x0000_t75" style="${style}" o:bullet="t">` +
  `<v:imagedata r:id="rId9"${crop}/></v:shape></w:pict></w:numPicBullet>`;

const numbering = (declared: string): string =>
  `<?xml version="1.0"?><w:numbering xmlns:w="${WORDPROCESSING_NS}" xmlns:r="${R_NS}"` +
  ` xmlns:v="${V_NS}" xmlns:o="urn:schemas-microsoft-com:office:office">${declared}</w:numbering>`;

// The relationship the picture is named through is the numbering part's own, whose
// ids are a different namespace from the document's: `rId9` is the picture here and
// nothing at all over there.
const NUMBERING_RELS = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}">
  <Relationship Id="rId9" Target="media/bullet.png" Type="${R_NS}/image"/></Relationships>`;

type Parts = Readonly<Record<string, string | Uint8Array>>;

const packageOf = (body: string, declared: string | null, extra: Parts = {}): Parts => ({
  "word/document.xml":
    `<?xml version="1.0"?><w:document xmlns:w="${WORDPROCESSING_NS}" xmlns:r="${R_NS}">` +
    `<w:body>${body}${SECTION}</w:body></w:document>`,
  ...(declared === null
    ? {}
    : {
        "word/numbering.xml": numbering(declared),
        "word/_rels/numbering.xml.rels": NUMBERING_RELS,
        "word/media/bullet.png": new Uint8Array([137, 80, 78, 71]),
      }),
  ...extra,
});

const opened = (body: string, declared: string | null, extra: Parts = {}) =>
  openDocx(buildDocx(packageOf(body, declared, extra)));

function laidOut(body: string, declared: string | null): LaidOutDocument {
  const laid = layOutDocument(opened(body, declared), bestEffortMetrics([], DEFAULTS));
  if (laid.kind !== "laid-out") throw new Error(`blocked: ${laid.blocker.kind}`);
  return laid;
}

function picturesOn(laid: LaidOutDocument) {
  const page = laid.pages[0];
  if (page === undefined) throw new Error("the document made no page");
  return drawablesOf(laid, page).flatMap((drawable) =>
    drawable.kind === "object" && drawable.content.kind === "picture"
      ? [
          {
            part: drawable.content.part,
            crop: drawable.content.crop,
            leftPt: drawable.leftPt,
            topPt: drawable.topPt,
            widthPt: drawable.widthPt,
            heightPt: drawable.heightPt,
          },
        ]
      : [],
  );
}

const paragraph = (text: string, properties = ""): string =>
  `<w:p>${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}<w:r><w:t>${text}</w:t></w:r></w:p>`;

describe("the picture a numbering part declares as a bullet", () => {
  it("names it through the numbering part's own relationships", () => {
    const read = pictureBulletOf(opened(paragraph("alpha"), bullet("width:36pt;height:24pt")));
    expect(read?.part).toBe("word/media/bullet.png");
  });

  it("finds none where the document has no numbering part", () => {
    expect(pictureBulletOf(opened(paragraph("alpha"), null))).toBeNull();
  });

  it("finds none where the numbering part declares no bullet", () => {
    expect(pictureBulletOf(opened(paragraph("alpha"), ""))).toBeNull();
  });

  // A relationship naming a part the package does not hold is a bullet nothing can
  // be drawn for, which is a different thing from one drawn as a missing picture:
  // Word draws no gap in the line where it has no picture bullet to put there.
  it("finds none where the relationship names a part the package has not got", () => {
    const read = pictureBulletOf(
      openDocx(
        buildDocx({
          "word/document.xml":
            `<?xml version="1.0"?><w:document xmlns:w="${WORDPROCESSING_NS}">` +
            `<w:body>${paragraph("alpha")}${SECTION}</w:body></w:document>`,
          "word/numbering.xml": numbering(bullet("width:36pt;height:24pt")),
          "word/_rels/numbering.xml.rels": NUMBERING_RELS,
        }),
      ),
    );
    expect(read).toBeNull();
  });
});

describe("where Word draws it, and what that does to the line", () => {
  // Measured: a first paragraph indented 36pt with 2pt above it had the picture drawn
  // at (72, 38) at the 36 by 24pt its shape states.
  it("stands at the first paragraph's own left and top", () => {
    const laid = laidOut(
      paragraph("alpha", `<w:ind w:left="720"/><w:spacing w:before="40"/>`) + paragraph("beta"),
      bullet("width:36pt;height:24pt"),
    );
    expect(picturesOn(laid)).toStrictEqual([
      {
        part: "word/media/bullet.png",
        crop: { left: 0, top: 0, right: 0, bottom: 0 },
        leftPt: 72,
        topPt: 38,
        widthPt: 36,
        heightPt: 24,
      },
    ]);
  });

  // Its bottom is the baseline and the text follows at its right edge, which is what
  // an inline picture does and is why nothing else in the layout hears about it.
  it("seats the first line's baseline at its foot and the text after it", () => {
    const laid = laidOut(paragraph("alpha") + paragraph("beta"), bullet("width:36pt;height:24pt"));
    const first = laid.pages[0]?.body[0]?.lines[0];
    expect(first?.baselinePt).toBeCloseTo(60, 6);
    const segments = first?.line.segments ?? [];
    expect(segments[0]?.kind).toBe("drawing");
    expect(segments[1]?.offsetPt).toBeCloseTo(36, 6);
  });

  it("is drawn once, however many paragraphs the body holds", () => {
    const laid = laidOut(
      paragraph("alpha") + paragraph("beta") + paragraph("gamma"),
      bullet("width:36pt;height:24pt"),
    );
    expect(picturesOn(laid)).toHaveLength(1);
  });

  // Measured: where the body opens with a table, Word draws it inside the first cell
  // at that cell's own text, so the first paragraph of the body is the first in the
  // order the body writes them and not the first standing in the flow.
  it("stands in the first cell where the body opens with a table", () => {
    const laid = laidOut(
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblCellMar>` +
        `<w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>` +
        `<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>` +
        `</w:tblCellMar></w:tblPr><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>` +
        `</w:tcPr>${paragraph("inside")}</w:tc></w:tr></w:tbl>${paragraph("after")}`,
      bullet("width:36pt;height:24pt"),
    );
    expect(picturesOn(laid).map((each) => each.leftPt)).toStrictEqual([36]);
  });

  // Measured: a shape stating 36 by 24pt and hiding four fifths of its picture across
  // and half of it down was drawn 180 by 48pt with 36 by 24 of it showing. So the
  // stated size is the window the crop leaves, which is what the renderers already
  // draw a cropped picture as.
  it("takes the stated size as the window the crop leaves", () => {
    const laid = laidOut(
      paragraph("alpha"),
      bullet(
        "width:36pt;height:24pt",
        ` cropleft="26214f" cropright="26214f" croptop="16384f" cropbottom="16384f"`,
      ),
    );
    const [drawn] = picturesOn(laid);
    expect(drawn?.widthPt).toBeCloseTo(36, 6);
    expect(drawn?.crop.left).toBeCloseTo(0.4, 3);
    expect(drawn?.crop.top).toBeCloseTo(0.25, 3);
  });
});
