import { buildAuthoredDocx, FACE } from "./package.js";

// Documents written here rather than found in the wild, so that Word's own answers
// about them can be committed: nothing in them is anyone's collateral. Each one
// asks about a single feature the seven reference documents never exercise, and is
// small enough that a failure names its own cause.
//
// Every measurement is a position Word reports, so the content is built to make a
// rule readable off a number: text long enough to reach the edge being tested, and
// nothing else on the line to confuse it.

export type AuthoredDocument = {
  readonly id: string;
  readonly title: string;
  // What the document is asking Word, so a failure says which rule is out.
  readonly asks: string;
  // Whether the question needs to know where a character sits along its line
  // rather than only where the line starts. Each one costs a round trip to Word.
  readonly measuresCharacters?: boolean;
  // How many of the characters Word placed are expected to land where Word put
  // them. Short of all of them names a gap this suite has measured and the layout
  // does not yet answer, which is a number that cannot quietly grow.
  readonly charactersPlaced?: number;
  readonly bytes: Uint8Array;
};

const text = (value: string): string =>
  value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

const run = (value: string, properties = ""): string =>
  `<w:r>${properties === "" ? "" : `<w:rPr>${properties}</w:rPr>`}<w:t xml:space="preserve">${text(value)}</w:t></w:r>`;

const paragraph = (properties: string, content: string): string =>
  `<w:p>${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}${content}</w:p>`;

const EMPTY = paragraph("", "");

const TAB = "<w:r><w:tab/></w:r>";

// A text box that fits itself to its text, which is what the open question about
// outlines is really about. `wrap` decides whether the width fits too.
type FittedBox = {
  readonly name: string;
  readonly outline: string;
  readonly wrap: "none" | "square";
  readonly insetEmu?: number;
  readonly content: string;
};

const fittedBox = (box: FittedBox, id: number): string => {
  const inset = box.insetEmu ?? 0;
  const insets = `lIns="${String(inset)}" tIns="${String(inset)}" rIns="${String(inset)}" bIns="${String(inset)}"`;
  return `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
    <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="10" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
      <wp:extent cx="2286000" cy="1143000"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:wrapNone/>
      <wp:docPr id="${String(id)}" name="${box.name}"/>
      <wp:cNvGraphicFramePr/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <wps:wsp><wps:cNvSpPr txBox="1"/>
          <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="1143000"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>${box.outline}</wps:spPr>
          <wps:txbx><w:txbxContent>${box.content}</w:txbxContent></wps:txbx>
          <wps:bodyPr rot="0" vert="horz" wrap="${box.wrap}" ${insets} anchor="t" anchorCtr="0"><a:spAutoFit/></wps:bodyPr>
        </wps:wsp></a:graphicData></a:graphic>
    </wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>`;
};

const LINE = "Antidisestablishmentarian quadrature";

// Each box stands in a paragraph of its own so that the shapes come back in a
// known order, one per paragraph.
const boxDocument = (boxes: readonly FittedBox[]): string =>
  boxes.map((box, at) => paragraph("", fittedBox(box, at + 1))).join("") + EMPTY;

const OUTLINE_WIDTHS = [
  {
    name: "unstated",
    outline: `<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>`,
  },
  { name: "none", outline: "" },
  {
    name: "0.75pt",
    outline: `<a:ln w="9525"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>`,
  },
  {
    name: "1.5pt",
    outline: `<a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>`,
  },
  {
    name: "3pt",
    outline: `<a:ln w="38100"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>`,
  },
  {
    name: "6pt",
    outline: `<a:ln w="76200"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>`,
  },
] as const;

// How wide and how tall Word makes a box that fits itself to its text, over every
// outline and both wrap modes. This is what says whether an outline goes to a
// fitted box's width, its height, or neither.
function fittingDocument(): string {
  const boxes: FittedBox[] = [];
  for (const wrap of ["none", "square"] as const) {
    for (const width of OUTLINE_WIDTHS) {
      boxes.push({
        name: `fit-${wrap}-${width.name}`,
        outline: width.outline,
        wrap,
        content: paragraph("", run(LINE)),
      });
    }
    // The same box with nothing in it, which still fits itself to its own mark.
    boxes.push({ name: `empty-${wrap}`, outline: "", wrap, content: EMPTY });
    // And with room round the text, which the fit has to account for.
    boxes.push({
      name: `inset-${wrap}`,
      outline: "",
      wrap,
      insetEmu: 91440,
      content: paragraph("", run(LINE)),
    });
  }
  return boxDocument(boxes);
}

// Where a tab lands, over every alignment a stop can take and the default stops
// between them. Each paragraph starts with text so the tab has somewhere to leave.
function tabDocument(): string {
  const stops = [
    { title: "left", stop: `<w:tab w:val="left" w:pos="2880"/>` },
    { title: "center", stop: `<w:tab w:val="center" w:pos="2880"/>` },
    { title: "right", stop: `<w:tab w:val="right" w:pos="2880"/>` },
    { title: "decimal", stop: `<w:tab w:val="decimal" w:pos="2880"/>` },
    { title: "bar", stop: `<w:tab w:val="bar" w:pos="2880"/>` },
  ];

  const cases = stops.map(({ stop }) =>
    paragraph(`<w:tabs>${stop}</w:tabs>`, run("ab") + TAB + run("1234.56") + TAB + run("cd")),
  );

  return [
    // No stop of its own, so every tab runs to the default every 720 twips.
    paragraph("", run("ab") + TAB + run("cd") + TAB + run("ef")),
    // A tab past the last stop falls back to the defaults beyond it.
    paragraph(
      `<w:tabs><w:tab w:val="left" w:pos="1440"/></w:tabs>`,
      run("ab") + TAB + run("cd") + TAB + run("ef"),
    ),
    ...cases,
    // Two stops, so the second tab knows to skip the one it has passed.
    paragraph(
      `<w:tabs><w:tab w:val="left" w:pos="1440"/><w:tab w:val="right" w:pos="5760"/></w:tabs>`,
      run("ab") + TAB + run("cd") + TAB + run("ef"),
    ),
    EMPTY,
  ].join("");
}

export function authoredDocuments(): readonly AuthoredDocument[] {
  return [
    {
      id: "fitting",
      title: "Boxes that fit themselves to their text",
      asks: "what an outline, an inset and an empty paragraph do to a fitted box's width and height",
      bytes: buildAuthoredDocx({ body: fittingDocument() }),
    },
    {
      id: "tabs",
      title: "Tab stops",
      asks: "where a tab lands at each alignment, and where the default stops fall",
      measuresCharacters: true,
      // A stop that is not left-aligned is not honoured yet: Word centres, ends or
      // aligns the decimal point of the text after the tab on the stop, and this
      // starts it there. The four paragraphs that ask are the last four.
      charactersPlaced: 51,
      bytes: buildAuthoredDocx({ body: tabDocument() }),
    },
  ];
}

export { FACE };
