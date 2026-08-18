import { describe, expect, it } from "vitest";

import { openDocx } from "../docx/package.js";
import { buildDocx, wordDocument, WORDPROCESSING_NS } from "../testing/build-docx.js";
import { buildFace } from "../testing/build-font.js";
import { bestEffortMetrics, type FaceDefaults } from "./best-effort.js";
import { layOutDocument, type LaidOutDocument } from "./document.js";

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

// A page keeping whatever it is told between its edge and its text, and between
// its edge and the room it holds for a header.
const section = (topTwips: number, headerTwips: number, header: boolean): string =>
  `<w:sectPr>` +
  (header ? `<w:headerReference w:type="default" r:id="rId1"/>` : "") +
  `<w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="${String(topTwips)}" w:right="720" w:bottom="720" w:left="720"` +
  ` w:header="${String(headerTwips)}" w:footer="720" w:gutter="0"/>` +
  `</w:sectPr>`;

const HEADER_PARTS: Readonly<Record<string, string>> = {
  "word/header1.xml": `<?xml version="1.0"?>
    <w:hdr xmlns:w="${WORDPROCESSING_NS}"><w:p><w:r><w:t>above</w:t></w:r></w:p></w:hdr>`,
  "word/_rels/document.xml.rels": `<?xml version="1.0"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Target="header1.xml" Type="${R_NS}/header"/>
    </Relationships>`,
};

function laidOut(topTwips: number, headerTwips: number, header: boolean): LaidOutDocument {
  const body = `<w:p><w:r><w:t>first</w:t></w:r></w:p>${section(topTwips, headerTwips, header)}`;
  // Written out rather than taken from `wordDocument`, which declares no `r` prefix
  // for a header to be referenced through.
  const bytes = buildDocx({
    "word/document.xml":
      `<?xml version="1.0"?><w:document xmlns:w="${WORDPROCESSING_NS}" xmlns:r="${R_NS}">` +
      `<w:body>${body}</w:body></w:document>`,
    ...(header ? HEADER_PARTS : {}),
  });
  const laid = layOutDocument(openDocx(bytes), bestEffortMetrics([], DEFAULTS));
  if (laid.kind !== "laid-out") throw new Error(`blocked: ${laid.blocker.kind}`);
  return laid;
}

describe("where a page starts its body", () => {
  // The room a header would have taken is not taken by a header that is not there,
  // which is what the footer has always said on its own side.
  it("starts at the top margin where the page draws no header", () => {
    const laid = laidOut(20, 720, false);

    expect(laid.bodyTopPt).toBe(1);
    expect(laid.pages[0]?.body[0]?.topPt).toBe(1);
  });

  it("starts under a header that reaches past the top margin", () => {
    const laid = laidOut(20, 720, true);

    expect(laid.bodyTopPt).toBe(laid.headerTopPt + (laid.pages[0]?.headerHeightPt ?? 0));
    expect(laid.bodyTopPt).toBeGreaterThan(36);
  });

  it("leaves the body at the top margin where the header stops above it", () => {
    const laid = laidOut(5760, 720, true);

    expect(laid.bodyTopPt).toBe(288);
  });
});

// A document of two pages, each opened by a section keeping its own room above its
// header. The first section ends at the first paragraph, so the second's text opens
// a page of its own.
function twoSections(firstHeaderTwips: number, lastHeaderTwips: number): LaidOutDocument {
  const body =
    `<w:p><w:pPr>${section(1440, firstHeaderTwips, true)}</w:pPr>` +
    `<w:r><w:t>first</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t>second</w:t></w:r></w:p>` +
    section(1440, lastHeaderTwips, true);
  const bytes = buildDocx({
    "word/document.xml":
      `<?xml version="1.0"?><w:document xmlns:w="${WORDPROCESSING_NS}" xmlns:r="${R_NS}">` +
      `<w:body>${body}</w:body></w:document>`,
    ...HEADER_PARTS,
  });
  const laid = layOutDocument(openDocx(bytes), bestEffortMetrics([], DEFAULTS));
  if (laid.kind !== "laid-out") throw new Error(`blocked: ${laid.blocker.kind}`);
  return laid;
}

// **How far down a page its header starts is its own section's business.** Read
// off the last section alone until 2026-08-10, which put 404 of the 556 header
// pictures in the corpus exactly 31.5pt below where Word drew them: a section
// keeping 90 twips above its header under a document whose last section keeps 720.
describe("where a page hangs its header", () => {
  it("hangs it from the room its own section keeps for it", () => {
    const laid = twoSections(90, 720);

    expect(laid.pages).toHaveLength(2);
    expect(laid.pages[0]?.headerTopPt).toBe(4.5);
    expect(laid.pages[1]?.headerTopPt).toBe(36);
  });

  // One part drawn under two sections that keep different room for it is two
  // drawings and not one, which a story remembered by its part alone got wrong.
  it("draws the one part it names at both heights", () => {
    const laid = twoSections(90, 720);
    const topOf = (at: number): number | undefined => laid.pages[at]?.header[0]?.topPt;

    expect(topOf(1)).toBe((topOf(0) ?? 0) + 31.5);
  });

  it("hangs it in one place where both sections keep the same room", () => {
    const laid = twoSections(720, 720);

    expect(laid.pages[0]?.headerTopPt).toBe(36);
    expect(laid.pages[1]?.headerTopPt).toBe(36);
  });
});

// What a box keeps text off. Every case below was put to Word on 2026-08-14 by
// `text-box-band-probe`, three repeats each. The page is the same one every time: a
// line ruled exactly 24pt standing 108 to 132 beside a box whose frame opens at 60,
// and the box stands against the left of the frame so the answer is where the line
// opens rather than whether it broke.
//
// **Word answers eight of the nine as this project does, and the ninth it does not**:
// a box whose text has run out of it keeps text off that text, and this keeps text off
// the box alone. That case is pinned below as this project answers it, with the corpus
// result that put it back beside it in `bandFor`.
const TIGHT = `<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon>
  <wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/>
  <wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>`;

const SQUARE = `<wp:wrapSquare wrapText="bothSides"/>`;

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";

const emu = (points: number): string => String(Math.round(points * 12700));

const EXACTLY_A_LINE = `<w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/>`;

const held = (lines: number, word: string): string =>
  Array.from(
    { length: lines },
    () => `<w:p><w:pPr>${EXACTLY_A_LINE}</w:pPr><w:r><w:t>${word}</w:t></w:r></w:p>`,
  ).join("");

// A box standing where the page's second line does, keeping its own text off the
// frame's left. Word is told the box's own size and never asked to fit it.
const boxAnchor = (options: {
  readonly wrap: string;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly lines: number;
  readonly bottomInsetPt?: number;
  readonly word?: string;
  readonly effectRightPt?: number;
  readonly effectTopPt?: number;
  // Where the box stands, which is against the frame's left at the page's second line
  // unless a case wants it elsewhere.
  readonly leftPt?: number;
  readonly topPt?: number;
}): string =>
  `<w:r><w:drawing><wp:anchor xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:wps="${WPS_NS}"
     behindDoc="0" relativeHeight="5" distT="0" distB="0" distL="114300" distR="114300">
     <wp:positionH relativeFrom="column"><wp:posOffset>${emu((options.leftPt ?? 36) - 36)}</wp:posOffset></wp:positionH>
     <wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu((options.topPt ?? 60) - 36)}</wp:posOffset></wp:positionV>
     <wp:extent cx="${emu(options.widthPt)}" cy="${emu(options.heightPt)}"/>
     <wp:effectExtent l="0" t="${emu(options.effectTopPt ?? 0)}" r="${emu(options.effectRightPt ?? 0)}" b="0"/>
     ${options.wrap}
     <wp:docPr id="1" name="Box"/>
     <a:graphic><a:graphicData uri="${WPS_NS}"><wps:wsp>
       <wps:spPr><a:prstGeom prst="rect"/></wps:spPr>
       <wps:txbx><w:txbxContent>${held(options.lines, options.word ?? "boxwording")}</w:txbxContent></wps:txbx>
       <wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="${emu(options.bottomInsetPt ?? 0)}"><a:noAutofit/></wps:bodyPr>
     </wps:wsp></a:graphicData></a:graphic>
   </wp:anchor></w:drawing></w:r>`;

const pictureAnchor = (heightPt: number): string =>
  `<w:r><w:drawing><wp:anchor xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}"
     behindDoc="0" relativeHeight="5" distT="0" distB="0" distL="114300" distR="114300">
     <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
     <wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(24)}</wp:posOffset></wp:positionV>
     <wp:extent cx="${emu(120)}" cy="${emu(heightPt)}"/>
     ${TIGHT}
     <wp:docPr id="1" name="Picture"/>
     <a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic><pic:blipFill/>
       <pic:spPr><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic>
   </wp:anchor></w:drawing></w:r>`;

// Where the fourth line of the page opens, which is the whole answer: 36 is the
// frame's own left, and anything else is the band's right edge and the 9pt the anchor
// holds text off by.
function opensAtPt(anchored: string): number {
  const line = (word: string, at: number): string =>
    `<w:p><w:pPr>${EXACTLY_A_LINE}</w:pPr>${at === 0 ? anchored : ""}<w:r><w:t>${word}</w:t></w:r></w:p>`;
  const body =
    [line("one", 0), line("two", 1), line("three", 2), line("four", 3)].join("") +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720"/></w:sectPr>`;

  const laid = layOutDocument(
    openDocx(buildDocx({ "word/document.xml": wordDocument(body) })),
    bestEffortMetrics([], DEFAULTS),
  );
  if (laid.kind !== "laid-out") throw new Error(`blocked: ${laid.blocker.kind}`);
  const asked = laid.pages[0]?.body[3]?.lines[0];
  if (asked === undefined) throw new Error("expected a fourth line");
  return asked.leftPt;
}

const BESIDE_THE_BOX_PT = 165;
const BESIDE_A_WIDE_BOX_PT = 345;
const THE_FRAME_PT = 36;

describe("what a wrapping box keeps text off", () => {
  it("keeps it off the whole of a box whose text stops short of the foot", () => {
    expect(opensAtPt(boxAnchor({ wrap: TIGHT, widthPt: 120, heightPt: 200, lines: 1 }))).toBe(
      BESIDE_THE_BOX_PT,
    );
  });

  it("keeps it off a box whose text stops exactly where the line starts", () => {
    expect(opensAtPt(boxAnchor({ wrap: TIGHT, widthPt: 120, heightPt: 200, lines: 2 }))).toBe(
      BESIDE_THE_BOX_PT,
    );
  });

  it("keeps it off a box whose text covers the line outright", () => {
    expect(opensAtPt(boxAnchor({ wrap: TIGHT, widthPt: 120, heightPt: 200, lines: 3 }))).toBe(
      BESIDE_THE_BOX_PT,
    );
  });

  // **The one Word answers otherwise**: 48pt of box holding 96pt of text, with the
  // line standing below the box and inside the text. Word narrows it and this leaves
  // it, because keeping text off the text cost five corpus documents a quarter of
  // their cells against one document gained.
  it("leaves a line under a box its text has run out of", () => {
    expect(opensAtPt(boxAnchor({ wrap: TIGHT, widthPt: 120, heightPt: 48, lines: 4 }))).toBe(
      THE_FRAME_PT,
    );
  });

  it("keeps it off a box holding room under its own text", () => {
    expect(
      opensAtPt(
        boxAnchor({ wrap: TIGHT, widthPt: 120, heightPt: 200, lines: 1, bottomInsetPt: 30 }),
      ),
    ).toBe(BESIDE_THE_BOX_PT);
  });

  // A square wrap is the box and nothing else, in both directions: the same pair of
  // boxes answers the same as the tight one above and the opposite below.
  it("keeps it off the whole of a square box whose text stops short", () => {
    expect(opensAtPt(boxAnchor({ wrap: SQUARE, widthPt: 120, heightPt: 200, lines: 1 }))).toBe(
      BESIDE_THE_BOX_PT,
    );
  });

  it("leaves a line under a square box the text has run out of", () => {
    expect(opensAtPt(boxAnchor({ wrap: SQUARE, widthPt: 120, heightPt: 48, lines: 4 }))).toBe(
      THE_FRAME_PT,
    );
  });

  it("keeps it off a picture, which holds no text to run out of", () => {
    expect(opensAtPt(pictureAnchor(200))).toBe(BESIDE_THE_BOX_PT);
  });

  // And nothing follows the text across the page: a box far wider than what it holds
  // keeps its own width.
  it("keeps it off the width of a box whose text is narrower than it", () => {
    expect(
      opensAtPt(boxAnchor({ wrap: TIGHT, widthPt: 300, heightPt: 200, lines: 1, word: "tick" })),
    ).toBe(BESIDE_A_WIDE_BOX_PT);
  });
});

// **Text is kept off what a drawing is drawn as, not off what it measures.** Word
// writes down how far past its own extent a drawing reaches, and the two reference
// pages that state one put their lines exactly that much further out than this did
// without it: 1.00pt in one and 0.75 in another, both beside a square text box whose
// outline is stated at half a point, so it is the number Word writes and not the
// outline it draws. `effect-extent-probe` then put every edge to Word on 2026-08-15,
// three repeats each, and it grows the band on all of them: the two below across the
// page, and, down it, a box whose top stops a quarter of a point below a line breaks
// that line stating half a point over itself and leaves it alone stating none.
describe("what a drawing reaching past its own extent keeps text off", () => {
  it("keeps it off a whole point of effect", () => {
    expect(
      opensAtPt(
        boxAnchor({ wrap: SQUARE, widthPt: 120, heightPt: 200, lines: 1, effectRightPt: 1 }),
      ),
    ).toBe(BESIDE_THE_BOX_PT + 1);
  });

  it("keeps it off three quarters of one", () => {
    expect(
      opensAtPt(
        boxAnchor({ wrap: SQUARE, widthPt: 120, heightPt: 200, lines: 1, effectRightPt: 0.75 }),
      ),
    ).toBe(BESIDE_THE_BOX_PT + 0.75);
  });
});

// The pair that settles the edges a band reaches a line by rather than narrows it
// with. Word was asked it with the box against the right of the frame, where the
// answer is whether the line broke; here the same box stands against the left, where
// the answer is where the line opens, and the geometry it turns on is the same: a
// quarter of a point of daylight, and half a point of effect to cross it.
describe("what a drawing reaching past its top keeps text off", () => {
  const below = (effectTopPt: number): string =>
    boxAnchor({ wrap: SQUARE, widthPt: 120, heightPt: 200, lines: 1, topPt: 132.25, effectTopPt });

  it("leaves a line alone where the box stops a quarter of a point below it", () => {
    expect(opensAtPt(below(0))).toBe(THE_FRAME_PT);
  });

  it("keeps it off a box stating half a point over its own top", () => {
    expect(opensAtPt(below(0.5))).toBe(BESIDE_THE_BOX_PT);
  });
});

// **A column run is measured against the room its own page left it**, and finding that
// room takes a pass of the pages and one more to see that the pass changed nothing.
// Without the confirming pass a document of one run never settles and is measured as
// though its run stood at the top of a page, so its columns run past the foot and the
// break pass spills what is over into the next page in the order the text was written
// rather than into columns. Read off `c8ca0c3c8292` on 2026-08-18, whose second page came
// out 27.6pt low that way.
describe("a column run standing partway down a page", () => {
  const EXACT = `<w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/>`;
  const line = (text: string) => `<w:p><w:pPr>${EXACT}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const closes = (text: string, columns: string) =>
    `<w:p><w:pPr>${EXACT}<w:sectPr><w:type w:val="continuous"/>` +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720"` +
    ` w:footer="720" w:gutter="0"/>${columns}</w:sectPr></w:pPr>` +
    `<w:r><w:t>${text}</w:t></w:r></w:p>`;

  // The body runs 36 to 756. Twenty lines of 24 fill it to 516, which leaves the run
  // 240pt of its page: ten lines a column, twenty of its thirty blocks. The ten left
  // over open the next page, five in each of its columns.
  const laid = (): LaidOutDocument => {
    const fillers = Array.from({ length: 19 }, (_, at) => line(`f${String(at)}`)).join("");
    const run = Array.from({ length: 29 }, (_, at) => line(`r${String(at)}`)).join("");
    const bytes = buildDocx({
      "word/document.xml": wordDocument(
        // The fillers close a section of their own, since a `w:sectPr` describes the
        // section ending with the paragraph carrying it and one hung on the run's last
        // block would put the fillers in the run.
        fillers + closes("f19", "") + run + closes("r29", `<w:cols w:num="2" w:space="720"/>`),
      ),
    });
    const result = layOutDocument(openDocx(bytes), bestEffortMetrics([], DEFAULTS));
    if (result.kind !== "laid-out") throw new Error(`blocked: ${result.blocker.kind}`);
    return result;
  };

  it("fills its columns to the foot of the page it opened on and no further", () => {
    const page = laid().pages[0];
    const feet = (page?.body ?? []).map((box) => box.topPt + box.heightPt);
    expect(Math.max(0, ...feet)).toBeLessThanOrEqual(756);
  });

  it("carries what the room would not hold into the next page rather than past the foot", () => {
    const pages = laid().pages;
    // Ten of the run's thirty blocks are over: the room its page left it is 240, which is
    // ten lines a column, and the twenty that fit stand in two columns of 516 to 756.
    expect(pages).toHaveLength(2);
    expect(pages[0]?.body.filter((box) => box.topPt >= 516)).toHaveLength(20);
    expect(pages[1]?.body[0]?.topPt).toBe(36);
  });
});
