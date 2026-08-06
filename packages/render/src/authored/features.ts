import { LEFT_PT, PICTURE_ID, RIGHT_PT } from "./package.js";

// The bodies of the authored documents that ask about a feature of the flowing
// text rather than about a shape. Each paragraph is written so that the rule it
// asks about shows up in where Word puts that paragraph or a character of it.

const text = (value: string): string =>
  value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

export const run = (value: string, properties = ""): string =>
  `<w:r>${properties === "" ? "" : `<w:rPr>${properties}</w:rPr>`}<w:t xml:space="preserve">${text(value)}</w:t></w:r>`;

export const paragraph = (properties: string, content: string): string =>
  `<w:p>${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}${content}</w:p>`;

export const EMPTY = paragraph("", "");

// Long enough to run past any width these documents narrow a line to, so that
// where it breaks is a fact about the narrowing rather than about the text.
const FLOW =
  "chlorophyll quadrature windbreak granulation microphone repository " +
  "thunderclap pendulum wavelength cartography hemisphere kaleidoscope";

const SHORT = "quadrature windbreak";

// How tall a row is, how far a cell holds its text off its own walls, and where
// the text inside one starts. A cell's own margins stand instead of the table's.
//
// Every cell says which one it is, since Word's own answer about a paragraph
// inside a table is the row's origin rather than the paragraph's: only its pdf
// says where the second cell of a row put its text, and only text that names its
// cell can be told apart there. Every margin asked about is different at the top
// from the bottom, so that text seated in the middle of the room a margin opens
// cannot be mistaken for text held off the top wall by it.
export function tableDocument(): string {
  const grid = `<w:tblGrid><w:gridCol w:w="2700"/><w:gridCol w:w="2700"/></w:tblGrid>`;

  const cell = (properties: string, content: string): string =>
    `<w:tc><w:tcPr><w:tcW w:w="2700" w:type="dxa"/>${properties}</w:tcPr>${content}</w:tc>`;

  const row = (properties: string, cells: string): string =>
    `<w:tr>${properties === "" ? "" : `<w:trPr>${properties}</w:trPr>`}${cells}</w:tr>`;

  const table = (properties: string, rows: string): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="5400" w:type="dxa"/>${properties}</w:tblPr>${grid}${rows}</w:tbl>`;

  const margins = (top: number, bottom: number): string =>
    `<w:tblCellMar>
      <w:top w:w="${String(top)}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="${String(bottom)}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar>`;

  const own = (sides: Readonly<Record<string, number>>): string =>
    `<w:tcMar>${Object.entries(sides)
      .map(([side, twips]) => `<w:${side} w:w="${String(twips)}" w:type="dxa"/>`)
      .join("")}</w:tcMar>`;

  const pair = (left: string, right: string): string =>
    cell("", paragraph("", run(left))) + cell("", paragraph("", run(right)));

  return [
    paragraph("", run("above")),

    // A row is as tall as the most its cells hold, which the second cell here holds
    // two lines of.
    table(
      margins(0, 0),
      row(
        "",
        cell("", paragraph("", run("a one"))) +
          cell("", paragraph("", run("a two")) + paragraph("", run("a three"))),
      ),
    ),
    paragraph("", run("between")),

    // A table whose cells hold their text off the top and bottom walls, which this
    // project has been ignoring.
    table(margins(288, 576), row("", pair("b one", "b two"))),
    paragraph("", run("between")),

    // A row asking to be at least as tall as a number of twips, and one asking to
    // be exactly that tall whatever its text needs.
    table(
      margins(0, 0),
      row(`<w:trHeight w:val="1440"/>`, pair("c one", "c two")) +
        row(`<w:trHeight w:val="288" w:hRule="exact"/>`, pair("c three", "c four")) +
        row(`<w:trHeight w:val="1440" w:hRule="exact"/>`, pair("c five", "c six")),
    ),
    paragraph("", run("between")),

    // A cell with margins of its own, which stand instead of the table's. Whether
    // the cell beside it, which has none, is held off the top wall as well is what
    // says if a margin belongs to its cell or to the whole row.
    table(
      margins(0, 0),
      row(
        "",
        cell(own({ top: 432, left: 432, bottom: 0, right: 108 }), paragraph("", run("d one"))) +
          cell("", paragraph("", run("d two"))),
      ),
    ),
    paragraph("", run("between")),

    // The same question of the second cell, which is the one Word will not answer
    // for: a margin only the second cell asks for either lifts the row or does not.
    table(
      margins(0, 0),
      row(
        "",
        cell("", paragraph("", run("e one"))) +
          cell(own({ top: 432, bottom: 0 }), paragraph("", run("e two"))),
      ),
    ),
    paragraph("", run("between")),

    // A cell holding its text off the bottom wall alone, which grows the row
    // without moving the text down unless the room is shared out around it.
    table(
      margins(144, 144),
      row(
        "",
        cell(own({ bottom: 432 }), paragraph("", run("f one"))) +
          cell("", paragraph("", run("f two"))),
      ),
    ),
    paragraph("", run("between")),

    // A row told exactly how tall to be, with less room in it than its own margin
    // asks to hold the text off the wall by.
    table(
      margins(0, 0),
      row(
        `<w:trHeight w:val="288" w:hRule="exact"/>`,
        cell(own({ top: 432 }), paragraph("", run("g one"))) +
          cell("", paragraph("", run("g two"))),
      ),
    ),
    paragraph("", run("below")),
    EMPTY,
  ].join("");
}

// Where a line sits when its paragraph asks for room above and below it, and what
// happens to a run of paragraphs that ask not to keep any between their own kind.
export function spacingDocument(): string {
  const spacing = (properties: string): string => `<w:spacing ${properties}/>`;

  return [
    paragraph("", run("above")),
    paragraph(spacing(`w:before="240" w:after="240"`), run(SHORT)),
    paragraph(spacing(`w:before="0" w:after="0"`), run(SHORT)),

    // A line told exactly how tall to be, which is not what its text measures, and
    // one told a floor it already clears.
    paragraph(spacing(`w:line="480" w:lineRule="exact"`), run(SHORT)),
    paragraph(spacing(`w:line="120" w:lineRule="exact"`), run(SHORT)),
    paragraph(spacing(`w:line="480" w:lineRule="atLeast"`), run(SHORT)),
    paragraph(spacing(`w:line="120" w:lineRule="atLeast"`), run(SHORT)),
    paragraph(spacing(`w:line="360" w:lineRule="auto"`), run(SHORT)),

    // Which side of the text the room a rule leaves opens on. Each of these asks it
    // of a different amount of room, 3.36 and 21.36 points, so an answer that is a
    // share of the room rather than the whole of it cannot pass.
    paragraph(spacing(`w:line="360" w:lineRule="atLeast"`), run(SHORT)),
    paragraph(spacing(`w:line="720" w:lineRule="atLeast"`), run(SHORT)),
    paragraph(spacing(`w:line="720" w:lineRule="exact"`), run(SHORT)),
    // Text taller than the floor it is given, which leaves no room to seat it in.
    // This is the size the room is asked at as well: an answer measuring the room
    // against a height of its own rather than against the text's would seat this
    // one lower. A line of this size with room over it is not asked for, since
    // Word's answer about one lands two thirds of a point below where its own pdf
    // draws it, which is more than the whole answer is rounded to.
    paragraph(spacing(`w:line="480" w:lineRule="atLeast"`), run(SHORT, `<w:sz w:val="48"/>`)),
    // Every line of a paragraph takes the room, not the first alone: this one runs
    // to two lines and the paragraph after it says how tall both came out.
    paragraph(spacing(`w:line="480" w:lineRule="atLeast"`), run(FLOW)),
    // A paragraph with nothing in it but its mark answers to the rule as well.
    paragraph(spacing(`w:line="480" w:lineRule="atLeast"`), ""),
    paragraph(spacing(`w:line="480" w:lineRule="exact"`), ""),

    // What happens where one paragraph's room below it meets the next one's room
    // above. Both pairs ask for the same two amounts the other way round, so the
    // four answers worth telling apart all read differently: adding them gives 36
    // to both, taking the larger gives 24 to both, and letting either side win
    // outright gives 12 to one and 24 to the other.
    paragraph("", run("pairs")),
    paragraph(spacing(`w:after="480"`), run("wide after")),
    paragraph(spacing(`w:before="240"`), run("narrow before")),
    paragraph("", run("between pairs")),
    paragraph(spacing(`w:after="240"`), run("narrow after")),
    paragraph(spacing(`w:before="480"`), run("wide before")),
    paragraph("", run("after pairs")),

    // Three of a kind that keep no room between themselves but still stand off
    // whatever is either side of the run.
    paragraph("", run("before the run")),
    ...["one", "two", "three"].map((each) =>
      paragraph(
        `<w:pStyle w:val="Listed"/>${spacing(`w:before="240" w:after="240"`)}<w:contextualSpacing/>`,
        run(each),
      ),
    ),
    paragraph("", run("after the run")),
    EMPTY,
  ].join("");
}

// The style the contextual run above shares, so that Word has a kind to compare.
export const SPACING_STYLES = `<w:style w:type="paragraph" w:styleId="Listed"><w:name w:val="Listed"/></w:style>`;

// Where a line starts and ends when an object wraps beside it, and whether a line
// narrowed by one is broken again at the narrower width.
export function wrappingDocument(): string {
  const wrapping = (id: number, wrap: string, widthEmu: number, heightEmu: number): string =>
    `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
      <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${String(id)}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="${String(widthEmu)}" cy="${String(heightEmu)}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        ${wrap}
        <wp:docPr id="${String(id)}" name="wrap-${String(id)}"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:cNvSpPr/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${String(widthEmu)}" cy="${String(heightEmu)}"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              <a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></wps:spPr>
            <wps:bodyPr/>
          </wps:wsp></a:graphicData></a:graphic>
      </wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>`;

  return [
    paragraph("", run("above")),
    // A square wrap two inches wide beside four lines of text.
    paragraph("", wrapping(1, `<wp:wrapSquare wrapText="bothSides"/>`, 1828800, 914400)),
    paragraph("", run(FLOW)),
    paragraph("", run("between")),
    // One wrapped top and bottom, which takes the whole width with it.
    paragraph("", wrapping(2, `<wp:wrapTopAndBottom/>`, 1828800, 457200)),
    paragraph("", run(FLOW)),
    paragraph("", run("below")),
    EMPTY,
  ].join("");
}

const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";

const emu = (pt: number): string => String(Math.round(pt * 12700));

// A drawing in the flow of the text, which is the one thing in these documents
// that stands on a line without being measured from a face. Its run still states
// a face like any other, which is what says whether the line hears anything the
// drawing itself does not say.
const inlinePicture = (id: number, widthPt: number, heightPt: number, mark = ""): string => {
  const extent = `cx="${emu(widthPt)}" cy="${emu(heightPt)}"`;
  return `<w:r>${mark === "" ? "" : `<w:rPr>${mark}</w:rPr>`}<w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent ${extent}/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:docPr id="${String(id)}" name="picture-${String(id)}"/>
      <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
      <a:graphic><a:graphicData uri="${PIC_NS}">
        <pic:pic xmlns:pic="${PIC_NS}">
          <pic:nvPicPr><pic:cNvPr id="${String(id)}" name="picture-${String(id)}"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="${PICTURE_ID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext ${extent}/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
        </pic:pic>
      </a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r>`;
};

// Taller than any line the text of this document would make on its own, so that
// the drawing is what the line has to hold.
const TALL_PT = 30;
// Shorter than the 14.65pt line 12pt Calibri makes, so that a line holding one
// says which of the two decides its height.
const SQUAT_PT = 6;
const WIDE_PT = 30;

// How many times each case is written out. Word rounds the answer it gives for a
// paragraph to the whole point, and its pdf draws a picture where it drew it: the
// distance between one drawing and the next of a run of the same paragraph is the
// height of that paragraph exactly, whatever the line does with its text.
const REPEATS = 3;

// One paragraph holding one drawing: the rule its paragraph carries, how tall the
// drawing is, what its own run states, and what stands on the line beside it.
type DrawingCase = {
  readonly properties: string;
  readonly heightPt: number;
  readonly mark?: string;
  readonly beside?: string;
};

// How tall a line holding an inline drawing comes out under each line rule.
//
// A drawing stands on the baseline like a letter and reaches as far above it as
// it is tall, but it was never measured from a face: what a line multiple is
// taken of, and what a paragraph mark still has to say once a drawing is on its
// line, are questions the rules measured off text alone do not answer. The chart
// in the LibreOffice sample is 162pt tall under a rule asking for 1.2 lines, and
// Word gives that line about 165pt rather than 194.6.
export function drawingDocument(): string {
  const spacing = (properties: string): string => `<w:spacing ${properties}/>`;
  const auto = (twips: number): string => spacing(`w:line="${String(twips)}" w:lineRule="auto"`);
  const large = `<w:sz w:val="48"/>`;
  // The mark's own size, which a paragraph states after everything else it asks
  // for. A mark never raises a line it shares with a run, and a drawing is a run.
  const largeMark = `<w:rPr>${large}</w:rPr>`;

  const cases: readonly DrawingCase[] = [
    // The rule every other paragraph in these documents carries, and two multiples
    // of it: whatever the multiple is taken of, these three tell it from the
    // drawing's own height.
    { properties: "", heightPt: TALL_PT },
    { properties: auto(288), heightPt: TALL_PT },
    { properties: auto(480), heightPt: TALL_PT },

    // A slot shorter than the drawing and one taller than it, asked both ways.
    { properties: spacing(`w:line="288" w:lineRule="exact"`), heightPt: TALL_PT },
    { properties: spacing(`w:line="1200" w:lineRule="exact"`), heightPt: TALL_PT },
    { properties: spacing(`w:line="288" w:lineRule="atLeast"`), heightPt: TALL_PT },
    { properties: spacing(`w:line="1200" w:lineRule="atLeast"`), heightPt: TALL_PT },

    // Text on the line beside the drawing, at the size the rest of the document is
    // and at twice it. A multiple taken of the text shows up as a difference
    // between these two, and one taken of the drawing does not.
    { properties: auto(288), heightPt: TALL_PT, beside: run("beside") },
    { properties: auto(288), heightPt: TALL_PT, beside: run("beside", large) },
    // The same two questions of the mark instead of a run, at one multiple and at
    // none, which is what says whether the mark is on the line at all.
    { properties: `${auto(240)}${largeMark}`, heightPt: TALL_PT },
    { properties: `${auto(288)}${largeMark}`, heightPt: TALL_PT },
    // And of the drawing's own run, which is the other face on the line and the
    // only one the mark cannot speak for.
    { properties: auto(288), heightPt: TALL_PT, mark: large },

    // A drawing shorter than the line the paragraph's own text would make, which
    // either holds the line open or leaves it to the text. Asked of the mark and
    // of the drawing's run in turn, since whichever holds the line open is the
    // one the multiple above is taken of as well.
    { properties: "", heightPt: SQUAT_PT },
    { properties: auto(288), heightPt: SQUAT_PT },
    { properties: "", heightPt: SQUAT_PT, beside: run("beside") },
    { properties: largeMark, heightPt: SQUAT_PT },
    { properties: "", heightPt: SQUAT_PT, mark: large },
  ];

  let drawings = 0;
  const withDrawing = (each: DrawingCase): string => {
    drawings += 1;
    return paragraph(
      each.properties,
      inlinePicture(drawings, WIDE_PT, each.heightPt, each.mark) + (each.beside ?? ""),
    );
  };

  return [
    paragraph("", run("above")),
    ...cases.flatMap((each) => Array.from({ length: REPEATS }, () => withDrawing(each))),
    // A mark twice the size with nothing on its line to trump it, which says that
    // the size the cases above put on a mark is one Word read.
    paragraph(largeMark, ""),
    paragraph("", run("below")),
    EMPTY,
  ].join("");
}

// What a page break does with the room a paragraph asks for either side of
// itself.
//
// The body of an authored page is 720pt tall, and every paragraph here but the
// ones being asked about is told exactly how tall to be, so the room left at the
// foot of a page is arithmetic rather than a measurement. Each block fills a page:
// a marker line, nine fillers, a shim sized to leave the room the case wants, and
// the case itself. The marker opening the next block is where the flow resumed,
// which says whether room a break swallowed was kept anywhere.
export function breakingDocument(): string {
  const exactly = (pt: number): string =>
    `<w:spacing w:line="${String(pt * 20)}" w:lineRule="exact"/>`;
  const room = (properties: string): string => `<w:spacing ${properties}/>`;

  const FILLERS = 9;
  const FILLER_PT = 72;
  const MARKER_PT = 24;
  const BLOCK_PT = 720;

  // Enough room at the foot of the page for the case's own line and not for the
  // room it asks for around it, which is 14.65pt of line against 18pt of room.
  const LINE_PT = 24;
  // Less room than the line itself needs, which is the case that has always moved
  // the paragraph on.
  const NO_LINE_PT = 12;

  const marker = (text: string): string => paragraph(exactly(MARKER_PT), run(text));

  // A page's worth of paragraphs ending in the case being asked about, with the
  // room left in front of that case named in points.
  const block = (leftPt: number, ...cases: readonly string[]): readonly string[] => [
    ...Array.from({ length: FILLERS }, () => paragraph(exactly(FILLER_PT), run("filler"))),
    paragraph(exactly(BLOCK_PT - MARKER_PT - FILLERS * FILLER_PT - leftPt), run("shim")),
    ...cases,
  ];

  return [
    marker("above"),
    // Room below the paragraph that the page has no room for, which either holds
    // the paragraph back or is swallowed by the break.
    ...block(LINE_PT, paragraph(room(`w:after="360"`), run("after"))),
    marker("after the after"),
    // The same question of the room above a paragraph.
    ...block(LINE_PT, paragraph(room(`w:before="360"`), run("before"))),
    marker("after the before"),
    // A paragraph whose own line the page has no room for, which has always been
    // held back, so that the answers above read against something.
    ...block(NO_LINE_PT, paragraph("", run("no room"))),
    marker("after the squeeze"),
    // Room below the paragraph greater than a whole page, which cannot be kept
    // anywhere and either holds the paragraph back or is dropped.
    ...block(LINE_PT, paragraph(room(`w:after="14400"`), run("wide after"))),
    marker("after the wide after"),
    // The same as the first, of a paragraph with nothing in it but its mark, which
    // is what the sample from outside the family stacks at the foot of its page.
    ...block(LINE_PT, paragraph(room(`w:after="360"`), "")),
    marker("after the empty"),
    EMPTY,
  ].join("");
}

// Where the text goes when a document breaks its own pages.
//
// Every paragraph here is told exactly how tall to be, so where one landed is
// arithmetic: the body of an authored page starts 36pt down and is 720pt tall, and
// a paragraph at the top of a page reports 36. Each case is followed by a marker
// naming it, and it is the marker that answers: 36 says the page opened with the
// text after the break, 60 says the break left a line of its own above it, and the
// page it landed on says whether a break with nothing to carry over made a page
// with nothing on it.
export function pageDocument(): string {
  const BREAK = `<w:r><w:br w:type="page"/></w:r>`;
  // A paragraph asking to start a page of its own. It stands before the spacing in
  // w:pPr, which Word requires in that order.
  const OWN_PAGE = `<w:pageBreakBefore/>`;

  const LINE_PT = 24;
  const ROOM_TWIPS = 360;

  const room = (twips: number): string => `w:before="${String(twips)}" `;

  const line = (content: string, properties = "", spacing = ""): string =>
    paragraph(
      `${properties}<w:spacing ${spacing}w:line="${String(LINE_PT * 20)}" w:lineRule="exact"/>`,
      content,
    );

  const marker = (name: string): string => line(run(name));

  const cell = (content: string): string =>
    `<w:tc><w:tcPr><w:tcW w:w="2700" w:type="dxa"/></w:tcPr>${content}</w:tc>`;

  return [
    // A break asked for by the first paragraph of the document, which has no page
    // in front of it to leave: either it opens on page two or the ask is spent.
    line(run("first"), OWN_PAGE),
    marker("after first"),

    // A break in the middle of a paragraph's own text, which carries the rest of
    // the paragraph to the next page.
    line(run("split") + BREAK + run("resumed")),
    marker("after split"),

    // A break with nothing after it but the paragraph's own mark, which either
    // stands at the top of the next page or goes with the break.
    line(run("trailing") + BREAK),
    marker("after trailing"),

    // A paragraph holding the break and nothing else.
    line(BREAK),
    marker("after lone"),

    // Two breaks together, which either leave a page with nothing on it or not.
    line(BREAK + BREAK),
    marker("after double"),

    // A paragraph asking for a page of its own in the middle of one.
    line(run("asked"), OWN_PAGE),
    marker("after asked"),

    // The same ask of a paragraph already standing at the top of a page, since the
    // one in front of it ended with a break: either it is already where it asked to
    // be or it makes an empty page to get there.
    line(run("ends") + BREAK),
    line(run("at the top"), OWN_PAGE),
    marker("after at the top"),

    // The room a paragraph keeps above itself where a break has just put it at the
    // top of a page: either the page opens with the room or the break swallowed it.
    line(run("ends again") + BREAK),
    line(run("room above"), "", room(ROOM_TWIPS)),
    marker("after room above"),

    // The room the paragraph in front of a break keeps below itself, which has
    // nowhere on that page to go.
    line(run("room below") + BREAK, "", `w:after="${String(ROOM_TWIPS)}" `),
    marker("after room below"),

    // And the room a paragraph asking for a page of its own keeps above itself.
    line(run("asked with room"), OWN_PAGE, room(ROOM_TWIPS)),
    marker("after asked with room"),

    // A break inside a cell, which has a wall round it that the flow of the page
    // does not. The margins are stated because nothing else here states them: an
    // authored document declares no table style, and a table that asks for none of
    // its own is held off its walls by nothing at all.
    `<w:tbl><w:tblPr><w:tblW w:w="5400" w:type="dxa"/><w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2700"/><w:gridCol w:w="2700"/></w:tblGrid>
      <w:tr>${cell(line(run("in a cell") + BREAK + run("still in it")))}${cell(line(run("beside it")))}</w:tr>
    </w:tbl>`,
    marker("after the table"),
    EMPTY,
  ].join("");
}

// Where Word draws the lines round a table and the colour behind its text.
//
// Nothing here is read off a paragraph's position: a border is painted rather
// than laid out, and Word's pdf reports each one as a filled rectangle. Every
// border and every fill in this document is given a colour of its own, since the
// colour is the only thing in that report that says which question it answers.
// The cases that could move the text as well carry text that names them, so the
// same rendering says whether the layout moved.
export function bordersDocument(): string {
  const CELL_TWIPS = 1440;
  const MARGIN_TWIPS = 108;

  const edge = (side: string, style: string, eighths: number, color: string, space = 0): string =>
    `<w:${side} w:val="${style}" w:sz="${String(eighths)}" w:space="${String(space)}" w:color="${color}"/>`;

  const around = (style: string, eighths: number, color: string, space = 0): string =>
    ["top", "left", "bottom", "right"]
      .map((side) => edge(side, style, eighths, color, space))
      .join("");

  const cellBorders = (edges: string): string => `<w:tcBorders>${edges}</w:tcBorders>`;
  const shading = (color: string, pattern = "clear", foreground = "auto"): string =>
    `<w:shd w:val="${pattern}" w:color="${foreground}" w:fill="${color}"/>`;

  const cell = (properties: string, content: string, widthTwips = CELL_TWIPS): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${String(widthTwips)}" w:type="dxa"/>${properties}</w:tcPr>${content}</w:tc>`;

  // The margins are stated because an authored document declares no table style,
  // and a table asking for none of its own would hold its text off its walls by
  // nothing at all.
  const margins = `<w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="${String(MARGIN_TWIPS)}" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="${String(MARGIN_TWIPS)}" w:type="dxa"/>
    </w:tblCellMar>`;

  const table = (properties: string, columns: number, cells: string): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="${String(columns * CELL_TWIPS)}" w:type="dxa"/>
      <w:tblInd w:w="0" w:type="dxa"/>${margins}${properties}</w:tblPr>
      <w:tblGrid>${Array.from({ length: columns }, () => `<w:gridCol w:w="${String(CELL_TWIPS)}"/>`).join("")}</w:tblGrid>
      <w:tr>${cells}</w:tr></w:tbl>`;

  // One cell with a border of its own on every side, which says how thick Word
  // draws each width and where it hangs the line off the cell's own edge. The
  // text names the case so the same rendering says whether a thick border moved
  // it: a cell edge is at 36pt and its text starts 5.4pt inside that.
  const alone = (name: string, properties: string): string =>
    table("", 1, cell(properties, paragraph("", run(name))));

  const WIDTHS = [
    { eighths: 2, color: "FF0000" },
    { eighths: 4, color: "00B050" },
    { eighths: 6, color: "0070C0" },
    { eighths: 8, color: "FF00FF" },
    { eighths: 12, color: "00B0F0" },
    { eighths: 18, color: "FFC000" },
    { eighths: 24, color: "7030A0" },
    { eighths: 36, color: "C00000" },
    { eighths: 48, color: "008080" },
  ] as const;

  // The styles a document in the wild asks for. Each is drawn at the one width, so
  // what comes back says how many rectangles the style is made of and how they
  // stand against the single line of the same width.
  const STYLES = [
    { style: "single", color: "E97132" },
    { style: "double", color: "196B24" },
    { style: "dashed", color: "0F9ED5" },
    { style: "dotted", color: "A02B93" },
    { style: "thick", color: "4EA72E" },
    { style: "dotDash", color: "B10202" },
    { style: "none", color: "3B7D23" },
    { style: "nil", color: "D86DCB" },
  ] as const;

  const STYLE_EIGHTHS = 12;

  return [
    paragraph("", run("above")),

    // How thick each width is drawn and where it sits.
    ...WIDTHS.flatMap((each) => [
      alone(`w${String(each.eighths)}`, cellBorders(around("single", each.eighths, each.color))),
      paragraph("", run("between")),
    ]),

    // What each style makes of the one width.
    ...STYLES.flatMap((each) => [
      alone(`s-${each.style}`, cellBorders(around(each.style, STYLE_EIGHTHS, each.color))),
      paragraph("", run("between")),
    ]),

    // Two cells that both state the line between them, which only one of them can
    // draw. The wider is asked for from each side in turn, and then the two are
    // made the same width so that something other than the width has to settle it.
    table(
      "",
      2,
      cell(cellBorders(edge("right", "single", 8, "FF0000")), paragraph("", run("c1 thin"))) +
        cell(cellBorders(edge("left", "single", 24, "0070C0")), paragraph("", run("c1 thick"))),
    ),
    paragraph("", run("between")),
    table(
      "",
      2,
      cell(cellBorders(edge("right", "single", 24, "FF0000")), paragraph("", run("c2 thick"))) +
        cell(cellBorders(edge("left", "single", 8, "0070C0")), paragraph("", run("c2 thin"))),
    ),
    paragraph("", run("between")),
    table(
      "",
      2,
      cell(cellBorders(edge("right", "single", 12, "FF0000")), paragraph("", run("c3 left"))) +
        cell(cellBorders(edge("left", "single", 12, "0070C0")), paragraph("", run("c3 right"))),
    ),
    paragraph("", run("between")),
    // One side asking for no line at all against a side that asks for one, which
    // says whether nil is a width of nothing or a refusal that carries.
    table(
      "",
      2,
      cell(cellBorders(edge("right", "nil", 0, "auto")), paragraph("", run("c4 nil"))) +
        cell(cellBorders(edge("left", "single", 12, "0070C0")), paragraph("", run("c4 line"))),
    ),
    paragraph("", run("between")),

    // The table states the lines and a cell overrides one of them: whether a cell
    // can rub out a line the table drew, and whether the table's own inside line
    // stands where neither cell asks for anything.
    table(
      `<w:tblBorders>${around("single", 8, "FFC000")}${edge("insideV", "single", 8, "FFC000")}${edge("insideH", "single", 8, "FFC000")}</w:tblBorders>`,
      2,
      cell("", paragraph("", run("d1 kept"))) +
        cell(cellBorders(edge("left", "nil", 0, "auto")), paragraph("", run("d1 rubbed"))),
    ),
    paragraph("", run("between")),
    // The table's inside line against a wider one a cell asks for.
    table(
      `<w:tblBorders>${edge("insideV", "single", 4, "FFC000")}</w:tblBorders>`,
      2,
      cell(cellBorders(edge("right", "single", 24, "7030A0")), paragraph("", run("d2 wide"))) +
        cell("", paragraph("", run("d2 none"))),
    ),
    paragraph("", run("between")),

    // What a fill covers: a cell on its own, a cell whose border is thick enough
    // to say which side of the line the fill stops at, and a fill under a pattern.
    table("", 1, cell(shading("FFF2CC"), paragraph("", run("e1 fill")))),
    paragraph("", run("between")),
    table(
      "",
      1,
      cell(
        `${cellBorders(around("single", 48, "C00000"))}${shading("DEEBF7")}`,
        paragraph("", run("e2 fill")),
      ),
    ),
    paragraph("", run("between")),
    table("", 1, cell(shading("FFFF00", "pct25", "FF0000"), paragraph("", run("e3 pattern")))),
    paragraph("", run("between")),
    // A fill the paragraph asks for rather than the cell, inside a cell that asks
    // for none: the two rectangles are not the same shape.
    table("", 1, cell("", paragraph(shading("E2EFDA"), run("e4 paragraph")))),
    paragraph("", run("between")),

    // A paragraph's own fill and its own border, out in the flow of the page where
    // the text area's edges are known: the text runs from 36pt to 576pt.
    paragraph(shading("FBE5D6"), run("f1 fill")),
    paragraph("", run("between")),
    paragraph(`<w:ind w:left="720" w:right="1440"/>${shading("E2F0D9")}`, run("f2 indented fill")),
    paragraph("", run("between")),
    // A border under the paragraph, at no distance from it and then at twelve
    // points, which says both where the line goes and whether the room it asks for
    // is room the paragraph takes.
    paragraph(`<w:pBdr>${edge("bottom", "single", 12, "FF0000")}</w:pBdr>`, run("f3 under")),
    paragraph("", run("between")),
    paragraph(
      `<w:pBdr>${edge("bottom", "single", 12, "0070C0", 12)}</w:pBdr>`,
      run("f4 under at twelve"),
    ),
    paragraph("", run("between")),
    // A box round the paragraph, which says how far past the text it reaches at
    // each side and what the room it asks for does there.
    paragraph(`<w:pBdr>${around("single", 12, "7030A0", 6)}</w:pBdr>`, run("f5 boxed")),
    paragraph("", run("between")),
    paragraph(
      `<w:ind w:left="720" w:right="1440"/><w:pBdr>${around("single", 12, "00B050", 0)}</w:pBdr>`,
      run("f6 boxed and indented"),
    ),
    paragraph("", run("between")),
    // Three paragraphs asking for the same box, which Word either draws three
    // times or joins into one.
    ...["f7 first", "f7 second", "f7 third"].map((each) =>
      paragraph(`<w:pBdr>${around("single", 12, "C00000", 0)}</w:pBdr>`, run(each)),
    ),
    paragraph("", run("between")),
    // A fill under a paragraph that keeps room above and below itself, which says
    // whether the room is part of what is filled.
    paragraph(`<w:spacing w:before="240" w:after="240"/>${shading("FFE699")}`, run("f8 spaced")),
    paragraph("", run("between")),

    // The lines of a table stand where the widths of the borders themselves put
    // them, and half of an outer one falls outside the table. These ask which side
    // of the line the table starts on, and how much room a line between two rows
    // takes from each of them: one border at a time, wide enough that half of it
    // cannot be mistaken for the whole.
    table("", 1, cell(cellBorders(edge("left", "single", 48, "FF6699")), paragraph("", run("g1")))),
    paragraph("", run("between")),
    table("", 1, cell(cellBorders(edge("top", "single", 48, "66FF99")), paragraph("", run("g2")))),
    paragraph("", run("between")),
    `<w:tbl><w:tblPr><w:tblW w:w="${String(CELL_TWIPS)}" w:type="dxa"/>
      <w:tblInd w:w="0" w:type="dxa"/>${margins}
      <w:tblBorders>${edge("insideH", "single", 48, "9966FF")}</w:tblBorders></w:tblPr>
      <w:tblGrid><w:gridCol w:w="${String(CELL_TWIPS)}"/></w:tblGrid>
      <w:tr>${cell("", paragraph("", run("g3 over")))}</w:tr>
      <w:tr>${cell("", paragraph("", run("g3 under")))}</w:tr></w:tbl>`,
    paragraph("", run("between")),
    // A line of text pushed against the right wall of a cell, which is the only way
    // to read where that wall is: with a wide border either side of it, the answer
    // says whether the room a cell gives its text is the whole of the cell or the
    // cell inside its own borders.
    table(
      "",
      1,
      cell(
        cellBorders(around("single", 48, "996633")),
        paragraph(`<w:jc w:val="right"/>`, run("g4")),
      ),
    ),
    paragraph("", run("between")),
    table("", 1, cell("", paragraph(`<w:jc w:val="right"/>`, run("g5")))),
    paragraph("", run("between")),

    // Where a table's own indent puts it. The lines say where the table's edge
    // went and the text says where the cell's own margin then held it off that
    // edge, which is the whole question: an indent smaller than the margin either
    // adds to it or is swallowed by it.
    ...[0, 54, 720].flatMap((twips, at) => [
      `<w:tbl><w:tblPr><w:tblW w:w="${String(CELL_TWIPS)}" w:type="dxa"/>
        <w:tblInd w:w="${String(twips)}" w:type="dxa"/>${margins}</w:tblPr>
        <w:tblGrid><w:gridCol w:w="${String(CELL_TWIPS)}"/></w:tblGrid>
        <w:tr>${cell(
          cellBorders(around("single", 8, ["FF3399", "33CC33", "3366FF"][at] ?? "000000")),
          paragraph("", run(`i${String(twips)} indent`)),
        )}</w:tr></w:tbl>`,
      paragraph("", run("between")),
    ]),

    // A wave, which is not drawn between the walls of its own width at all: at
    // four eighths and at twenty four it reaches the same 2.64pt across, and at
    // twelve 3.36. Nothing here reads that off a width, so the wave is drawn as a
    // plain line of the width it states and the rows it lines come out shorter
    // than Word's. These cases stand last for that reason: what a case cannot
    // answer for should not move everything under it.
    ...[4, 12, 24].flatMap((eighths) => [
      alone(`h${String(eighths)} wave`, cellBorders(around("wave", eighths, "156082"))),
      paragraph("", run("between")),
    ]),
    paragraph("", run("below")),
    EMPTY,
  ].join("");
}

// A list numbered in the body and the same list inside a text box, which is what
// says whether a box starts the counting again.
export function numberingDocument(): string {
  const numbered = (numId: number, level: number, value: string): string =>
    paragraph(
      `<w:pStyle w:val="Listed"/><w:numPr><w:ilvl w:val="${String(level)}"/><w:numId w:val="${String(numId)}"/></w:numPr>`,
      run(value),
    );

  return [
    paragraph("", run("above")),
    numbered(1, 0, "one"),
    numbered(1, 0, "two"),
    numbered(1, 1, "under two"),
    numbered(1, 0, "three"),
    paragraph("", run("between")),
    // The same list again, which carries on rather than starting over.
    numbered(1, 0, "four"),
    // A level whose number is wider than the room its indent leaves in front of
    // the text, which has to push the first line along.
    numbered(2, 0, "wide marker"),
    numbered(2, 0, "wide marker again"),
    paragraph("", run("below")),
    EMPTY,
  ].join("");
}

export const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="Section %1.%1.%1:"/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="144"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

// Where a wrapping object stands, and what falls past it, in a document that
// declares no compatibility mode.
//
// Word puts such a document's objects on the twip grid and leaves a modern one's
// where the flow put them. An object anchored at the top of its paragraph stands
// at the foot of the paragraph before it, so rounding down puts its wrap band over
// that paragraph's last line, which then falls to the object's foot. Whether the
// rounding goes up or down is all that decides it, and the size of the face above
// is what moves the fraction: at 11pt each case here rounds down and fires, at
// 10.5pt it rounds up and nothing moves. The same body is written twice, once
// declaring 15 and once declaring nothing, so the difference is the setting.
//
// Every case stands on a page of its own, and the object is exactly as wide as the
// column so that nothing can sit beside it.
export function compatibilityDocument(): string {
  const anchored = (id: number, widthPt: number): string =>
    `<w:r><w:drawing><wp:anchor behindDoc="0" distT="0" distB="0" distL="0" distR="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1" relativeHeight="${String(id)}">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV>
      <wp:extent cx="${emu(widthPt)}" cy="${emu(OBJECT_PT)}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:wrapSquare wrapText="largest"/>
      <wp:docPr id="${String(id)}" name="anchored-${String(id)}"/>
      <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
      <a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic xmlns:pic="${PIC_NS}">
        <pic:nvPicPr><pic:cNvPr id="${String(id)}" name="anchored-${String(id)}"/><pic:cNvPicPr/></pic:nvPicPr>
        <pic:blipFill><a:blip r:embed="${PICTURE_ID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
        <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(widthPt)}" cy="${emu(OBJECT_PT)}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
      </pic:pic></a:graphicData></a:graphic>
    </wp:anchor></w:drawing></w:r>`;

  const cases = [
    { name: "one line rounding down", sizeHalfPt: 22, lines: 1, widthPt: COLUMN_PT },
    { name: "one line rounding up", sizeHalfPt: 21, lines: 1, widthPt: COLUMN_PT },
    // Three lines, to say which of them the object moves.
    { name: "three lines rounding down", sizeHalfPt: 28, lines: 3, widthPt: COLUMN_PT },
    { name: "three lines rounding up", sizeHalfPt: 26, lines: 3, widthPt: COLUMN_PT },
    // An object narrow enough to leave the line somewhere to sit, which is the
    // whole difference between falling past one and being narrowed by it.
    { name: "beside a narrow object", sizeHalfPt: 22, lines: 1, widthPt: 240 },
    // Nothing but a paragraph mark above the anchor, which is the shape a document
    // out of another word processor met this rule with.
    { name: "an empty paragraph rounding down", sizeHalfPt: 22, lines: 0, widthPt: COLUMN_PT },
  ];

  return cases
    .map((each, at) => {
      const mark = `<w:sz w:val="${String(each.sizeHalfPt)}"/><w:szCs w:val="${String(each.sizeHalfPt)}"/>`;
      const spacing = `<w:spacing w:before="0" w:after="225" w:line="288" w:lineRule="auto"/>`;
      const above = Array.from({ length: each.lines }, (_, line) =>
        run(`above ${String(at + 1)} line ${String(line + 1)}`, mark),
      ).join(`<w:r><w:rPr>${mark}</w:rPr><w:br/></w:r>`);

      return (
        paragraph(at === 0 ? "" : "<w:pageBreakBefore/>", run(`case ${String(at + 1)}`)) +
        paragraph(`${spacing}<w:rPr>${mark}</w:rPr>`, above) +
        paragraph("", anchored(at + 1, each.widthPt)) +
        paragraph("", run(`below ${String(at + 1)}`))
      );
    })
    .join("");
}

// As wide as the text column, so the object leaves no room beside it, and short
// enough that a case and the object it holds fit on one page.
const COLUMN_PT = RIGHT_PT - LEFT_PT;
const OBJECT_PT = 72;

// A third of the column, so that sliding one across changes which side has the
// most room by a wide margin while both sides stay able to hold a line.
const WRAP_OBJECT_PT = 180;

// Which side of a wrapping object text may sit on, which `wrapText` names.
//
// Every case holds one object of the same size at a different place across the
// column, so what changes between them is how much room is left either side. Each
// asks its question by where the lines beside the object start: a line on the left
// opens at the column's own edge, a line on the right opens past the object, and
// a line allowed neither side falls below it. The two `largest` cases with room to
// spare on one side put that side against the plain `left` and `right` cases at the
// same offset, so a reading that always takes the wider side and one that takes the
// side it was told cannot both pass.
export function wrapSidesDocument(): string {
  const anchored = (id: number, wrapText: string, offsetPt: number): string =>
    `<w:r><w:drawing><wp:anchor behindDoc="0" distT="0" distB="0" distL="0" distR="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1" relativeHeight="${String(id)}">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="column"><wp:posOffset>${emu(offsetPt)}</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV>
      <wp:extent cx="${emu(WRAP_OBJECT_PT)}" cy="${emu(OBJECT_PT)}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:wrapSquare wrapText="${wrapText}"/>
      <wp:docPr id="${String(id)}" name="sided-${String(id)}"/>
      <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
      <a:graphic><a:graphicData uri="${PIC_NS}"><pic:pic xmlns:pic="${PIC_NS}">
        <pic:nvPicPr><pic:cNvPr id="${String(id)}" name="sided-${String(id)}"/><pic:cNvPicPr/></pic:nvPicPr>
        <pic:blipFill><a:blip r:embed="${PICTURE_ID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
        <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(WRAP_OBJECT_PT)}" cy="${emu(OBJECT_PT)}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
      </pic:pic></a:graphicData></a:graphic>
    </wp:anchor></w:drawing></w:r>`;

  const cases = [
    { wrapText: "largest", offsetPt: 120 },
    { wrapText: "largest", offsetPt: 240 },
    // Both sides the same, which is the one case where the widest side cannot
    // decide it.
    { wrapText: "largest", offsetPt: 180 },
    // Told a side against the run of the room: the narrow one here, and below it
    // the narrow one there.
    { wrapText: "left", offsetPt: 120 },
    { wrapText: "right", offsetPt: 240 },
    // Told a side with nothing on it at all, which is what says whether a line
    // refused its side gives up on the object or takes the other one.
    { wrapText: "left", offsetPt: 0 },
    { wrapText: "right", offsetPt: COLUMN_PT - WRAP_OBJECT_PT },
  ];

  return cases
    .map((each, at) => {
      const name = `${each.wrapText} at ${String(each.offsetPt)}`;
      return (
        paragraph(`${at === 0 ? "" : "<w:pageBreakBefore/>"}${EXACT_LINE}`, run(name)) +
        paragraph(EXACT_LINE, anchored(at + 1, each.wrapText, each.offsetPt)) +
        // Three lines, since the object is tall enough for all of them and a rule
        // that moved only the first would read as no rule at all.
        [1, 2, 3]
          .map((line) => paragraph(EXACT_LINE, run(`${String(at + 1)} beside ${String(line)}`)))
          .join("")
      );
    })
    .join("");
}

// A whole number of twips a line, so that every paragraph in the document stands on
// the twip grid. A legacy document's objects are rounded to that grid, and an
// object rounded off the place the flow gave it drops the paragraph above it past
// itself, which is a rule of its own and not the one being asked about here.
const EXACT_LINE = `<w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/>`;

// What Word draws for a character the face it is written in does not map.
//
// Three of these once stopped this project outright, and each turned out to be a
// different rule. `U+202F` is a narrow no-break space Arial has no glyph for;
// `U+00A0` is the ordinary no-break space, which every face here maps except
// Symbol; and `U+2022` is a bullet written into a Wingdings run, where Wingdings
// addresses its glyphs through a symbol page and has nothing at that code point.
// The last of the three is the only one Word answers out of another face
// altogether, which is why the face it reaches for is asked here in its own right.
//
// Each case is a pair: two letters alone, and the same two letters with the
// character between them. The difference between the two widths Word's own pdf
// draws is the width of the character, and a third line puts an ordinary space in
// the same place so a character drawn as a space says so. The pdf also names the
// face each run was drawn in, which is what says whether Word gave up on the
// stated face and reached for another.
export function unmappedCharacterDocument(): string {
  const cases = [
    { face: "Arial", character: " ", name: "narrow no-break space" },
    { face: "Symbol", character: " ", name: "no-break space" },
    { face: "Wingdings", character: "•", name: "bullet" },
    // The same character in a face that does map it, so the question is the face
    // rather than the character.
    { face: "Calibri", character: " ", name: "no-break space" },
    { face: "Calibri", character: " ", name: "narrow no-break space" },
    // The face Word reached for above, asked directly. Two files on this machine
    // answer to the name and they differ: the one Word ships states no line gap
    // and maps the hyphen, and the system's states a gap of 87 units and does not.
    // How tall these lines come out says which of them Word lays a paragraph out
    // in, which is the same question as which one the bullet above was drawn from.
    { face: "Times New Roman", character: "•", name: "bullet" },
  ];

  const inFace = (face: string, text: string): string =>
    run(text, `<w:rFonts w:ascii="${face}" w:hAnsi="${face}" w:cs="${face}"/>`);

  return cases
    .flatMap((each, at) => [
      paragraph(at === 0 ? "" : "<w:pageBreakBefore/>", run(`${each.face} ${each.name}`)),
      // Two letters alone, the two with the character between them, and the two
      // with a plain space between them.
      paragraph("", inFace(each.face, "HH")),
      paragraph("", inFace(each.face, `H${each.character}H`)),
      paragraph("", inFace(each.face, "H H")),
    ])
    .join("");
}

// The same question of a text face, which the document above never asks: every one
// of its faces either maps the character or addresses its glyphs through a symbol
// page, and those are the two rules already measured. A text face with no glyph and
// no page to alias to is the third, and it is the whole of what this project still
// turns a document down for.
//
// A case is a character in a text face that has no glyph for it. Each stands on a
// page of its own, so the faces Word's pdf names on that page beyond the one stated
// are the faces Word reached for itself, with no other case to confuse them.
//
// Every case is written out three times over, since the height of a line is the
// distance between one repeat and the next: a character borrowed from another face
// raises the line it stands on where that face is the taller, and the two letters
// either side of it are drawn in the stated face whatever happens to the character
// between them. Above them the same two letters alone, three times over as well, so
// the stated face's own line is the same measurement rather than a number from
// somewhere else.
//
// Word answered on 2026-08-06: it reaches for another face for a text face as it
// does for a symbol one, and it reached for five different faces over the eleven
// cases. Which face is not one name: a sans face borrows from Arial and a serif one
// from Times New Roman, and a character neither of those carries goes on to
// whichever face on the machine has it. See `docs/gaps.md`.
export function unmappedInTextFaceDocument(): string {
  const cases = [
    // The four geometric bullets a real document wants. Calibri carries three of
    // them and Cambria none, and Cambria is what a document naming a face nothing
    // supplies is laid out in, which is how a document meets them at all.
    { face: "Calibri", character: "■", name: "black square" },
    { face: "Cambria", character: "■", name: "black square" },
    // The same character in two more faces that have no glyph for it, one of each
    // kind, since two cases cannot tell a face chosen for the character from a face
    // chosen to go with the one that asked.
    { face: "Verdana", character: "■", name: "black square" },
    { face: "Georgia", character: "■", name: "black square" },
    { face: "Cambria", character: "▪", name: "small black square" },
    { face: "Cambria", character: "●", name: "black circle" },
    { face: "Cambria", character: "◦", name: "white bullet" },
    // The hyphen, which the two files answering to the name Times New Roman differ
    // over: the system's maps nothing at it and Word's own maps it at 682 units.
    // Asking it of the face itself is what says which copy answers for a glyph,
    // since the line height already says the system's answers for the metrics.
    { face: "Arial", character: "‐", name: "hyphen" },
    { face: "Times New Roman", character: "‐", name: "hyphen" },
    // A character with no glyph in any face on this machine, and one that is not
    // meant to be drawn at all: whether Word gives up on the character or on the
    // page is a rule of its own.
    { face: "Cambria", character: "\u{1D44E}", name: "italic a" },
    { face: "Calibri", character: "\u2060", name: "word joiner" },
  ];

  const inFace = (face: string, text: string): string =>
    run(text, `<w:rFonts w:ascii="${face}" w:hAnsi="${face}" w:cs="${face}"/>`);

  const thrice = (content: string): string => [content, content, content].join("");

  return cases
    .flatMap((each, at) => [
      // The line naming the case is in the face the case is about, so that a face
      // beyond it on the page is one Word chose rather than one the document wrote.
      paragraph(
        at === 0 ? "" : "<w:pageBreakBefore/>",
        inFace(each.face, `${each.face} ${each.name}`),
      ),
      thrice(paragraph("", inFace(each.face, "HH"))),
      thrice(paragraph("", inFace(each.face, `H${each.character}H`))),
      paragraph("", inFace(each.face, "H H")),
    ])
    .join("");
}
