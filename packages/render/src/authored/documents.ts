import {
  bordersDocument,
  breakingDocument,
  characterSpacingDocument,
  compatibilityDocument,
  drawingDocument,
  keepingDocument,
  numberingDocument,
  sectionsDocument,
  pageDocument,
  spacingDocument,
  tableDocument,
  tableIndentDocument,
  unmappedCharacterDocument,
  unmappedInTextFaceDocument,
  wrappingDocument,
  wrapSidesDocument,
  NUMBERING,
  SPACING_STYLES,
} from "./features.js";
import { buildAuthoredDocx, settingsPart, FACE } from "./package.js";

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
  // How many of the places Word reported are expected to be the places this
  // project puts them. Short of all of them names a gap the suite has measured and
  // the layout does not yet answer, which is then a number that cannot quietly
  // grow. Left out, every one of them has to agree.
  readonly paragraphsPlaced?: number;
  readonly charactersPlaced?: number;
  // Faces the document names beyond the one they are all written in, which only a
  // document asking about faces states. A machine without one of them is left out
  // of the suite for that document rather than laying it out in something else.
  readonly statedFaces?: readonly string[];
  // How many of the images in Word's own rendering are glyphs rather than
  // drawings. Word paints a colour emoji as a bitmap, so a document that borrows a
  // character from the emoji face has images in its pdf that nothing in it asked
  // for and nothing here draws as a picture.
  readonly glyphsPaintedAsImages?: number;
  // Why this project refuses the document today, where it does. A document is
  // written to ask a question, and a question worth asking is often one nothing
  // here can answer yet: it is committed so that Word's answer can be had by
  // anyone, and left out of the suites that lay a document out until there is
  // something to compare. Naming the reason is what stops it being forgotten.
  readonly refuses?: string;
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

  // What a stop that is not left-aligned reaches over, once there is no tab after
  // it to end at, and what it makes of text that has no decimal point in it, has
  // two, or is wider than the room in front of the stop.
  const reach = (title: string, stop: string, content: string): string =>
    paragraph(`<w:tabs><w:tab w:val="${title}" w:pos="${stop}"/></w:tabs>`, content);

  const ends = [
    reach("right", "2880", run("ab") + TAB + run("cd ef")),
    reach("center", "2880", run("ab") + TAB + run("cd ef")),
    // A space at the end of what the stop reaches over, which is either part of
    // what is being lined up or is not.
    reach("right", "2880", run("ab") + TAB + run("cd ") + TAB + run("ef")),
    reach("decimal", "2880", run("ab") + TAB + run("123456") + TAB + run("cd")),
    reach("decimal", "2880", run("ab") + TAB + run("1.2.3") + TAB + run("cd")),
    reach("decimal", "2880", run("ab") + TAB + run("no point here") + TAB + run("cd")),
    // More text than there is room for in front of the stop, which cannot be lined
    // up on it without running back over what is already on the line.
    reach("right", "720", run("ab") + TAB + run("antidisestablishmentarian") + TAB + run("cd")),
    reach("center", "720", run("ab") + TAB + run("antidisestablishmentarian") + TAB + run("cd")),
  ];

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
    ...ends,
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
      bytes: buildAuthoredDocx({ body: tabDocument() }),
    },
    {
      id: "tables",
      title: "Tables",
      asks: "how tall a row is and how far a cell holds its text off its own walls",
      bytes: buildAuthoredDocx({ body: tableDocument() }),
    },
    {
      id: "borders",
      title: "Table borders, paragraph borders and shading",
      asks: "how thick each border is drawn, where it hangs off the edge it runs along, and what a fill covers",
      // Twenty one of these are answers this project reads off Word's own pdf
      // instead. Eleven are the wave borders and what stands under them, which
      // Word draws at no width the file states. Eight are rows whose true top
      // falls within a tenth of a point of the half way mark Word's answer is
      // rounded at, and land the other side of it. The last two are paragraphs
      // with a border of their own, which Word answers for a point or so below
      // where its own pdf draws their text.
      paragraphsPlaced: 122,
      bytes: buildAuthoredDocx({ body: bordersDocument() }),
    },
    {
      id: "spacing",
      title: "Spacing between and within paragraphs",
      asks: "where a line sits under each line rule, and what contextual spacing closes up",
      bytes: buildAuthoredDocx({ body: spacingDocument(), extraStyles: SPACING_STYLES }),
    },
    {
      id: "character-spacing",
      title: "Spacing between the characters of a run",
      asks: "where the extra width lands, whether a space and a tab take it, and what justification makes of it",
      measuresCharacters: true,
      bytes: buildAuthoredDocx({ body: characterSpacingDocument() }),
    },
    {
      id: "wrapping",
      title: "Text wrapping beside an object",
      asks: "where a line starts and ends beside a wrapping object, and where it breaks",
      measuresCharacters: true,
      bytes: buildAuthoredDocx({ body: wrappingDocument() }),
    },
    {
      id: "wrap-sides",
      title: "The side of an object text is allowed on",
      asks: "which side of a wrapping object Word puts a line on, and what a line refused its side does",
      bytes: buildAuthoredDocx({ body: wrapSidesDocument(), picture: true }),
    },
    {
      id: "legacy-wrap-sides",
      title: "The same sides in a document that declares no compatibility mode",
      asks: "which of the side rules the compatibility mode decides",
      bytes: buildAuthoredDocx({
        body: wrapSidesDocument(),
        settings: settingsPart({ compatibilityMode: null }),
        picture: true,
      }),
    },
    {
      id: "unmapped-characters",
      title: "A character the face it is written in does not map",
      asks: "what Word draws where a face has no glyph for a character, and in which face",
      measuresCharacters: true,
      statedFaces: ["Arial", "Symbol", "Wingdings", "Times New Roman"],
      bytes: buildAuthoredDocx({ body: unmappedCharacterDocument() }),
    },
    {
      id: "unmapped-in-a-text-face",
      title: "A character a text face has no glyph for",
      asks: "which face Word draws a character out of when a text face has none, and how tall that leaves the line",
      // Calibri is not among these: it is the face every authored document is
      // written in, and a machine without it has no suite at all. The last three
      // are named by nothing in the document and stated here all the same, since
      // they are the faces Word drew its characters out of: a machine without one
      // lays the document out a different way and answers a different question.
      statedFaces: [
        "Cambria",
        "Arial",
        "Times New Roman",
        "Verdana",
        "Georgia",
        "Apple Color Emoji",
        "Cambria Math",
        "Segoe UI Symbol",
      ],
      // The three repeats of the small black square, which is the one case Word
      // answered out of the emoji face: each is a 12pt bitmap at 12pt, a whole em.
      glyphsPaintedAsImages: 3,
      bytes: buildAuthoredDocx({ body: unmappedInTextFaceDocument() }),
    },
    {
      id: "drawings",
      title: "Inline drawings",
      asks: "how tall the line holding an inline drawing comes out under each line rule",
      // Five of the drawings this document seats below the top of their own line
      // are answered for by Word within a point of where Word's own rendering puts
      // them, and rounding to the point is not enough to close that: three of them
      // land a whole point off the nearest rounding of what the pdf draws, and one
      // run of three identical paragraphs rounds two ways. Every drawing in the
      // document is pinned against the rendering itself, to the twentieth of a
      // point, which is the oracle that answers here.
      paragraphsPlaced: 50,
      bytes: buildAuthoredDocx({ body: drawingDocument(), picture: true }),
    },
    {
      id: "breaking",
      title: "Page breaks through a paragraph's spacing",
      asks: "whether the room a paragraph asks for either side of itself holds it back at the foot of a page",
      bytes: buildAuthoredDocx({ body: breakingDocument() }),
    },
    {
      id: "sections",
      title: "A document made of more than one section",
      asks: "which section a paragraph's own properties govern, and what a break's type describes",
      refuses: "nothing here reads a section but the last, so the pages are the last one's",
      bytes: buildAuthoredDocx({ body: sectionsDocument() }),
    },
    {
      id: "keeping",
      title: "Paragraphs held to the one after them",
      asks: "whether w:keepNext moves a paragraph onto the page its next one landed on, and how far back a run of them pulls",
      bytes: buildAuthoredDocx({ body: keepingDocument() }),
    },
    {
      id: "pages",
      title: "Explicit page breaks",
      asks: "where the text goes when the document breaks its own pages",
      bytes: buildAuthoredDocx({ body: pageDocument() }),
    },
    {
      id: "legacy-wrapping",
      title: "Wrapping in a document that declares no compatibility mode",
      asks: "where Word puts an anchored object without one, and whose line falls past it",
      bytes: buildAuthoredDocx({
        body: compatibilityDocument(),
        settings: settingsPart({ compatibilityMode: null }),
        picture: true,
      }),
    },
    {
      id: "modern-wrapping",
      title: "The same wrapping in a document that declares one",
      asks: "whether declaring 15 is what leaves an object where the flow put it",
      bytes: buildAuthoredDocx({ body: compatibilityDocument(), picture: true }),
    },
    {
      id: "legacy-table-indent",
      title: "A table's indent in a document that declares no compatibility mode",
      asks: "whether an indent is measured to the table's edge or to the text inside its first cell",
      bytes: buildAuthoredDocx({
        body: tableIndentDocument(),
        settings: settingsPart({ compatibilityMode: null }),
      }),
    },
    {
      id: "table-indent",
      title: "The same indents in a document that declares one",
      asks: "whether declaring 15 is what puts the cell margin beyond the indent",
      bytes: buildAuthoredDocx({ body: tableIndentDocument() }),
    },
    {
      id: "numbering",
      title: "Numbered lists",
      asks: "where a number sits, and where the text after one starts when the number is wide",
      measuresCharacters: true,
      bytes: buildAuthoredDocx({
        body: numberingDocument(),
        extraStyles: SPACING_STYLES,
        numbering: NUMBERING,
      }),
    },
  ];
}

export { FACE };
