import {
  bordersDocument,
  breakingDocument,
  breaksInAParagraphDocument,
  characterSpacingDocument,
  columnsDocument,
  compatibilityDocument,
  conditionalTableDocument,
  drawingDocument,
  insignificantSpaceDocument,
  justifiedFittingDocument,
  keepingDocument,
  linedRowsDocument,
  mergedCellsDocument,
  numberedFirstLineDocument,
  numberingDocument,
  objectsAndTheFooterDocument,
  objectsPastTheFootDocument,
  overflowingSectionDocument,
  positionedTableDocument,
  footerRoomDocument,
  breakLineMarkDocument,
  headerNotNamedDocument,
  sectionsAndTheFirstPageDocument,
  lineMultipleDocument,
  twipGridDocument,
  resumingDocument,
  rotatedDrawingDocument,
  rotatedDrawingTiesDocument,
  sectionCloserDocument,
  sectionFlowDocument,
  sectionPagesDocument,
  sectionsDocument,
  spaceAboveABreakDocument,
  statedRowHeightsDocument,
  tearingDocument,
  pageDocument,
  raisedTextDocument,
  spacingDocument,
  spaceUnderAWrapDocument,
  tableDocument,
  tableIndentDocument,
  trailingSpaceDocument,
  unmappedCharacterDocument,
  unmappedInTextFaceDocument,
  wrappingDocument,
  wrapSidesDocument,
  CONDITIONAL_STYLES,
  FIRST_LINE_NUMBERING,
  NUMBERING,
  SPACING_STYLES,
  STATED_FOOTER,
  DRAWN_FOOTER,
  EMPTY_FOOTER,
  SPACED_FOOTER,
  NAMED_FIRST_HEADER,
  NAMED_DEFAULT_HEADER,
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
  // Whether Word's own pdf draws one of this document's pictures as more than one
  // image, which leaves it out of the pairing the pictures are held to.
  readonly picturesWordDrewInPieces?: boolean;
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
      id: "objects-past-the-foot",
      title: "An anchored object whose foot falls past the bottom of its page",
      asks: "whether such an object hangs where it was put, moves to the next page on its own, or takes the paragraph anchoring it with it",
      bytes: buildAuthoredDocx({ body: objectsPastTheFootDocument() }),
    },
    {
      id: "breaks-in-a-paragraph",
      title: "A break with nothing on the line it opens",
      asks: "whether a break at the end of a paragraph, or a second one after it, still opens a line",
      bytes: buildAuthoredDocx({ body: breaksInAParagraphDocument() }),
    },
    {
      id: "justified-fitting",
      title: "The word a justified line is squeezed to take",
      asks: "how far past its own room a justified line's text may reach before the last word is sent down",
      // One family is written in Times New Roman, whose space is a quarter of its em
      // where Calibri's is 0.2261: that is what says whether the ceiling on the
      // squeeze is a length of the em or of the space.
      statedFaces: ["Times New Roman"],
      bytes: buildAuthoredDocx({ body: justifiedFittingDocument() }),
    },
    {
      id: "legacy-justified-fitting",
      title: "The same question of a document that declares no compatibility mode",
      asks: "whether a justified line is squeezed at all in an old document",
      statedFaces: ["Times New Roman"],
      bytes: buildAuthoredDocx({
        body: justifiedFittingDocument(),
        settings: settingsPart({ compatibilityMode: null }),
      }),
    },
    {
      id: "objects-and-the-footer",
      title: "The same object on a page that draws a footer, and several at one anchor",
      asks: "which foot an object is judged against when a footer holds the text off the bottom margin, and whether an anchor answers for its objects one at a time",
      bytes: buildAuthoredDocx({ body: objectsAndTheFooterDocument(), footer: STATED_FOOTER }),
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
      id: "header-not-named",
      title: "A section naming a first-page header and no other",
      asks: "what a page draws where its section names no header for the kind of page it is",
      bytes: buildAuthoredDocx({
        body: headerNotNamedDocument(),
        headers: { first: NAMED_FIRST_HEADER },
        titlePage: true,
      }),
    },
    {
      id: "sections-and-the-first-page",
      title: "A second section that began part way down a page, and the header it draws next",
      asks: "whether a page after a section that opened mid-page draws that section's first-page header",
      bytes: buildAuthoredDocx({
        body: sectionsAndTheFirstPageDocument(),
        headers: { first: NAMED_FIRST_HEADER, default: NAMED_DEFAULT_HEADER },
        // The last section names neither, so it draws whatever the first named.
        namesHeaders: [],
        titlePage: true,
        sectionType: "continuous",
      }),
    },
    {
      id: "break-line-mark",
      title: "The line a break opens, and whose size it takes",
      asks: "whether the line a break opens is measured from the run holding the break or from the paragraph's mark",
      bytes: buildAuthoredDocx({ body: breakLineMarkDocument() }),
    },
    {
      id: "twip-grid",
      title: "Lines whose height falls between two twips, stacked until the page ends",
      asks: "whether a line's height is rounded to the twip before it is stacked down the page",
      bytes: buildAuthoredDocx({ body: twipGridDocument() }),
    },
    {
      id: "line-multiple",
      title: "Lines under a rule asking for a multiple of one",
      asks: "what a line multiple is a multiple of, and which side of the text the room it opens falls",
      bytes: buildAuthoredDocx({ body: lineMultipleDocument() }),
    },
    {
      id: "empty-footer",
      title: "A footer of a stated height that draws nothing",
      asks: "whether a footer holding one empty paragraph holds the body off the bottom margin",
      bytes: buildAuthoredDocx({ body: footerRoomDocument(), footer: EMPTY_FOOTER }),
    },
    {
      id: "drawn-footer",
      title: "The same footer with a line of text in it",
      asks: "how far a footer that does draw holds the body off the bottom margin, which the one above is read against",
      bytes: buildAuthoredDocx({ body: footerRoomDocument(), footer: DRAWN_FOOTER }),
    },
    {
      id: "spaced-footer",
      title: "A footer that draws nothing and asks for room under itself",
      asks: "whether the space a footer's last paragraph asks for below it holds the body any further off",
      bytes: buildAuthoredDocx({ body: footerRoomDocument(), footer: SPACED_FOOTER }),
    },
    {
      id: "rotated-drawings",
      title: "An inline drawing turned after it was drawn",
      asks: "what room a turned drawing's line keeps, and where the turned picture is drawn in it",
      bytes: buildAuthoredDocx({ body: rotatedDrawingDocument(), picture: true }),
    },
    {
      id: "rotated-drawing-ties",
      title: "A drawing turned exactly between one quarter and the next",
      asks: "which quarter the room a turned drawing's line keeps is rounded to when it stands between two",
      // Word draws the picture turned by an eighth in three overlapping pieces and
      // the one turned by three eighths nowhere at all, so its pdf cannot be paired
      // with what this project draws. Word's own answer for each paragraph says
      // what the document asks, which is the room its line kept.
      picturesWordDrewInPieces: true,
      bytes: buildAuthoredDocx({ body: rotatedDrawingTiesDocument(), picture: true }),
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
      bytes: buildAuthoredDocx({ body: sectionsDocument() }),
    },
    {
      id: "section-pages",
      title: "Which page each section of a document opens on",
      asks: "which of the break types a section states opens a page, and which carries on down the one it is already on",
      // Twelve of the eighteen: every section opens the page Word opened it on
      // until the one asking for an even page, which Word gave page 6 from page 4
      // and this project gives page 5. A break reaching for a page of a parity
      // leaves a blank one behind where the next page is the wrong one, and a page
      // holding nothing is a page this project cannot make yet. The three
      // paragraphs after it are one page early for the same reason and no other.
      paragraphsPlaced: 12,
      bytes: buildAuthoredDocx({ body: sectionPagesDocument() }),
    },
    {
      id: "section-closer",
      title: "The paragraph that carries a section break, at the foot of its page",
      asks: "whether the mark of a paragraph doing nothing but closing a section takes room where there is none left",
      // Eighteen of the twenty two, the four left over being the closers the
      // document is about. Word's report answers for one of those with the top of
      // the paragraph above it and from a left the text column never reaches, and
      // gives the same answer whether the closer fits on its page or not, so it is
      // not a place at all. Word agrees about the page of every one of the
      // twenty two, and its own pdf puts all 18 drawn lines where this project
      // puts them.
      paragraphsPlaced: 18,
      bytes: buildAuthoredDocx({ body: sectionCloserDocument() }),
    },
    {
      id: "overflowing-section",
      title: "A continuous section whose own text runs past the foot of its page",
      asks: "which section's page geometry a page a continuous section ran on to is made by",
      bytes: buildAuthoredDocx({ body: overflowingSectionDocument() }),
    },
    {
      id: "raised-text",
      title: "A run raised or lowered off its own baseline",
      asks: "how far off the baseline w:position draws a run, and whether the line grows to hold it",
      bytes: buildAuthoredDocx({ body: raisedTextDocument() }),
    },
    {
      id: "insignificant-space",
      title: "Whitespace at the edge of a run that did not ask to keep it",
      asks: "whether Word keeps the space at the edge of a w:t that states no xml:space",
      // Not compared line by line against Word's own drawing, which is the one
      // document here that cannot be: Word draws an item of its own wherever the
      // formatting changes, and the cases that put a bold run beside a plain one draw
      // a line in two pieces where this project draws it in one. Word's report
      // answers for every paragraph of it, and where a line broke is pinned in
      // `lines.test.ts` instead.
      bytes: buildAuthoredDocx({ body: insignificantSpaceDocument() }),
    },
    {
      id: "preserved-space",
      title: "The same document stating on its own root that whitespace is the text's own",
      asks: "whether Word reads xml:space from wherever it is stated rather than from the w:t alone",
      // The one thing this document says differently from the one above, and the
      // reason it exists: the worst-placed document in the corpus states `preserve`
      // on `w:document` and on no `w:t` anywhere.
      bytes: buildAuthoredDocx({ body: insignificantSpaceDocument(), preservesSpace: true }),
    },
    {
      id: "positioned-table",
      title: "A table positioned rather than flowed",
      asks: "where w:tblpPr stands a table across and down the page, and what the text it left does with the room",
      // Sixty nine of the seventy one. The two left over are the case whose table
      // leaves a usable run of the frame on both sides of itself: Word draws one
      // line there and fills both runs with it, and this project has no line that
      // can stand in two places, so it breaks the paragraph at the first run's end
      // instead and everything under it is a line low. Neither of the other six
      // cases leaves a run wider than the 18pt a line needs to the left of the
      // table, and neither does any of the seven documents in the wild.
      paragraphsPlaced: 69,
      bytes: buildAuthoredDocx({ body: positionedTableDocument() }),
    },
    {
      id: "space-above-a-break",
      title: "The room a paragraph asks for above itself, at the top of a page",
      asks: "which kinds of page break keep the space before the paragraph that opens the page, and which drop it",
      bytes: buildAuthoredDocx({ body: spaceAboveABreakDocument() }),
    },
    {
      id: "space-under-a-wrap",
      title: "The room a paragraph asks for above itself, under an object that wraps",
      asks: "whether the space before a paragraph is absorbed by the band that pushed its first line down, or kept below it",
      bytes: buildAuthoredDocx({ body: spaceUnderAWrapDocument() }),
    },
    {
      id: "columns",
      title: "Text running in more than one column",
      asks: "where a section's columns stand across the page, and when its text leaves one for the next",
      // Sixty two of the sixty three. The one left over is case d, whose column break
      // stands between two runs of one paragraph: that is a place inside a block and
      // the division into columns is made between them, so the whole paragraph goes
      // to the next column where Word leaves the first half of it behind. Every
      // column break in the corpus stands alone in its paragraph or opens one.
      paragraphsPlaced: 62,
      bytes: buildAuthoredDocx({ body: columnsDocument() }),
    },
    {
      id: "section-flow",
      title: "Where the text of each section sits down the page",
      asks: "what a continuous break does with the top margin of the section it opens",
      bytes: buildAuthoredDocx({ body: sectionFlowDocument() }),
    },
    {
      id: "keeping",
      title: "Paragraphs held to the one after them",
      asks: "whether w:keepNext moves a paragraph onto the page its next one landed on, and how far back a run of them pulls",
      bytes: buildAuthoredDocx({ body: keepingDocument() }),
    },
    {
      id: "lined-rows",
      title: "The line drawn between two rows",
      asks: "whether a border takes room of its own between two rows, on top of the margins either side of it",
      // Eighteen of these are rows of the cases lined at a point and at three,
      // which Word's report puts 0.55 to 0.8pt below where its own pdf draws them.
      // Every one of the document's 150 drawn lines lands where the pdf has it,
      // which is the oracle for a position and the one this document was written to
      // read.
      paragraphsPlaced: 193,
      bytes: buildAuthoredDocx({ body: linedRowsDocument() }),
    },
    {
      id: "legacy-lined-rows",
      title: "The same lines in a document that declares no compatibility mode",
      asks: "whether declaring 15 is what puts a row's border on top of its margin",
      paragraphsPlaced: 193,
      bytes: buildAuthoredDocx({
        body: linedRowsDocument(),
        settings: settingsPart({ compatibilityMode: null }),
      }),
    },
    {
      id: "stated-row-heights",
      title: "A row asking to stand taller than its own text",
      asks: "whether the line between two rows takes room on top of the height they each asked for",
      // A hundred and forty of the 155. The fifteen are Word's report alone, which
      // puts the last row of four of the eleven cases 0.55 to 0.7pt below its own
      // drawing of the same row and answers from a left the text column never
      // reaches. Word's own pdf puts all 110 drawn lines where this project puts
      // them, which is the oracle a height is read by here.
      paragraphsPlaced: 140,
      bytes: buildAuthoredDocx({ body: statedRowHeightsDocument() }),
    },
    {
      id: "conditional-table",
      title: "A table style that formats one place differently from another",
      asks: "which of a style's conditional formats reach a cell, in what order, and which of them w:tblLook turns off",
      bytes: buildAuthoredDocx({
        body: conditionalTableDocument(),
        extraStyles: CONDITIONAL_STYLES,
      }),
    },
    {
      id: "numbered-first-line",
      title: "The room the first line of a numbered paragraph has",
      asks: "whether a first line whose number tabbed short of the left indent runs the whole way to the right one",
      bytes: buildAuthoredDocx({
        body: numberedFirstLineDocument(),
        numbering: FIRST_LINE_NUMBERING,
      }),
    },
    {
      id: "legacy-numbered-first-line",
      title: "The same first lines in a document that declares no compatibility mode",
      asks: "whether declaring 15 changes the room a numbered first line has",
      bytes: buildAuthoredDocx({
        body: numberedFirstLineDocument(),
        numbering: FIRST_LINE_NUMBERING,
        settings: settingsPart({ compatibilityMode: null }),
      }),
    },
    {
      id: "merged-cells",
      title: "A cell merged down a run of rows",
      asks: "where a merged cell's text sits, what its rows are worth, and where the room it is short comes from",
      bytes: buildAuthoredDocx({ body: mergedCellsDocument() }),
    },
    {
      id: "legacy-merged-cells",
      title: "The same merges in a document that declares no compatibility mode",
      asks: "whether declaring 15 changes what a merge is worth",
      bytes: buildAuthoredDocx({
        body: mergedCellsDocument(),
        settings: settingsPart({ compatibilityMode: null }),
      }),
    },
    {
      id: "tearing",
      title: "A table row the foot of a page falls in",
      asks: "whether a row too tall for the room left is torn across the break or moved whole",
      // Four of these are the eighth and seventeenth line of the two rows written
      // in 32pt lines, which Word answers for 0.55pt below where its own pdf draws
      // them. The answers climb a twentieth of a point a line down those rows and
      // fall back every ninth, which is a grid the report is rounded to rather than
      // anything the layout did: every one of the document's 154 drawn lines lands
      // where the pdf has it.
      paragraphsPlaced: 160,
      bytes: buildAuthoredDocx({ body: tearingDocument() }),
    },
    {
      id: "resuming",
      title: "The page a torn row resumes on",
      asks: "what a row puts above its own text where it resumes, and what closes a cell holding a table",
      // Four of the 142 in each document, which Word's own report puts between 0.43
      // and 0.6pt below where its own pdf draws them. Every one of the document's
      // 121 drawn lines lands where the pdf has it, in both documents, and two of
      // the four stand in a case with no border in it at all: nothing here rounds.
      paragraphsPlaced: 138,
      bytes: buildAuthoredDocx({ body: resumingDocument() }),
    },
    {
      id: "legacy-resuming",
      title: "The same rows in a document that declares no compatibility mode",
      asks: "whether declaring 15 changes where a torn row resumes",
      paragraphsPlaced: 138,
      bytes: buildAuthoredDocx({
        body: resumingDocument(),
        settings: settingsPart({ compatibilityMode: null }),
      }),
    },
    {
      id: "trailing-space",
      title: "A space at the end of a line",
      asks: "whether a space raises the line it stands on, and how tall a paragraph holding only spaces is",
      bytes: buildAuthoredDocx({ body: trailingSpaceDocument() }),
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
