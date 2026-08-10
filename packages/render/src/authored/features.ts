import { LEFT_PT, PICTURE_ID, RIGHT_PT, TOP_PT } from "./package.js";

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

// What `w:spacing` in a run's properties does to where each character sits.
//
// Five points, so that every answer stands clear of the whole point Word rounds
// its own to: over six characters the two readings of the first question are
// thirty points apart.
export function characterSpacingDocument(): string {
  const spaced = (twentieths: number): string => `<w:spacing w:val="${String(twentieths)}"/>`;
  const WIDE = spaced(100);
  const TIGHT = spaced(-20);
  const STOP = `<w:tabs><w:tab w:val="left" w:pos="2880"/></w:tabs>`;
  const TABBED = "<w:r><w:tab/></w:r>";

  // After every character, or only in the gaps between them? The paragraph mark
  // sits at the end of the text, so a trailing gap is the whole difference; the
  // right aligned pair asks the same from the other end, where a trailing gap
  // pushes the visible text one gap further left.
  const lands = [
    paragraph("", run("abcdef")),
    paragraph("", run("abcdef", WIDE)),
    paragraph(`<w:jc w:val="right"/>`, run("abcdef")),
    paragraph(`<w:jc w:val="right"/>`, run("abcdef", WIDE)),
  ];

  // Whether a space takes it, and whether a tab does. The stop is what makes the
  // tab readable: text after it starts at the stop if the tab took no gap.
  const between = [
    paragraph("", run("ab cd")),
    paragraph("", run("ab cd", WIDE)),
    paragraph(STOP, run("ab") + TABBED + run("cd")),
    paragraph(STOP, run("ab", WIDE) + TABBED + run("cd", WIDE)),
    paragraph(STOP, run("ab", WIDE) + TABBED + run("cd")),
  ];

  // The unspaced line says what the stretch is worth on its own, so the spaced
  // one says whether it was worked out over the widened text or the bare text.
  const justified = [
    paragraph(`<w:jc w:val="both"/>`, run(FLOW)),
    paragraph(`<w:jc w:val="both"/>`, run(FLOW, WIDE)),
  ];

  // Where one reading stops and the other starts: whether the last character of
  // a spaced run carries a gap into the plain text after it.
  const boundary = [
    paragraph("", run("abc") + run("def")),
    paragraph("", run("abc", WIDE) + run("def")),
    paragraph("", run("abc", TIGHT) + run("def")),
  ];

  const condensed = [paragraph("", run("abcdef", TIGHT))];

  return [...lands, ...between, ...justified, ...boundary, ...condensed, EMPTY].join("");
}

// What becomes of an anchored object whose foot falls past the bottom of the page
// its paragraph stands on.
//
// Eight documents in the corpus are a page of anchored text boxes over a page of
// flowing text, and they are the whole of what the clean corpus still misses outside
// sections. In every one of them the boxes near the top of the page land within a
// tenth of a point and the ones at the foot are drawn by Word on the page after. The
// page they hold no text of their own to be read by, so the corpus cannot say which
// of three things Word did, and all three put the box on the next page:
//
//   1. the object moved on and the flow stayed where it was,
//   2. the paragraph anchoring it moved on and took the object with it, or
//   3. the object hung past the foot and Word drew none of it.
//
// Every box here carries a line of its own text, so Word's own drawing says which
// page the object landed on and where, and the paragraph anchoring it and the one
// after say what the flow did. Each case opens a page of its own and is held down it
// by a shim told exactly how tall to be, so the room left under it is arithmetic:
// the page is 720pt of body, the marker takes 24 and the shim the rest.
export function objectsPastTheFootDocument(): string {
  const LINE_PT = 24;
  const BODY_PT = 720;
  const BOX_PT = 300;

  const exactly = (pt: number): string =>
    `<w:spacing w:line="${String(pt * 20)}" w:lineRule="exact"/>`;

  const OWN_PAGE = `<w:pageBreakBefore/>`;

  // A box of a stated size holding one line, which is what makes the object
  // readable: a picture says only where it was drawn and a pdf says nothing about
  // which anchor drew it.
  const boxed = (
    id: number,
    name: string,
    heightPt: number,
    wrap: string,
    offsetPt: number,
  ): string =>
    `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
      <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${String(id)}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(offsetPt)}</wp:posOffset></wp:positionV>
        <wp:extent cx="${emu(BOX_PT)}" cy="${emu(heightPt)}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        ${wrap}
        <wp:docPr id="${String(id)}" name="${name}"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:cNvSpPr txBox="1"/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(BOX_PT)}" cy="${emu(heightPt)}"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr>
            <wps:txbx><w:txbxContent>${paragraph(exactly(LINE_PT), run(`${name} boxed`))}</w:txbxContent></wps:txbx>
            <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"/>
          </wps:wsp></a:graphicData></a:graphic>
      </wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>`;

  const NONE = `<wp:wrapNone/>`;
  const SQUARE = `<wp:wrapSquare wrapText="bothSides"/>`;

  type Case = {
    readonly name: string;
    // The room standing under the shim, which is what the object is offered.
    readonly leftPt: number;
    readonly heightPt: number;
    readonly wrap: string;
    // How far below its paragraph the object hangs. Nought in every case but the
    // one written to the geometry of a document in the wild.
    readonly offsetPt?: number;
    // Whether the block is the last in the document, so the object has no page to
    // be moved onto.
    readonly last?: boolean;
  };

  const block = (of: Case, id: number): string =>
    paragraph(`${OWN_PAGE}${exactly(LINE_PT)}`, run(`${of.name} marks`)) +
    paragraph(exactly(BODY_PT - LINE_PT - of.leftPt), run(`${of.name} shims`)) +
    paragraph(
      exactly(LINE_PT),
      boxed(id, of.name, of.heightPt, of.wrap, of.offsetPt ?? 0) + run(`${of.name} anchors`),
    ) +
    (of.last === true ? "" : paragraph(exactly(LINE_PT), run(`${of.name} follows`)));

  const CASES: readonly Case[] = [
    // A hundred points of room and a box three times that, wrapping nothing, which
    // is the plain case and the one the corpus documents are.
    { name: "a", leftPt: 100, heightPt: BOX_PT, wrap: NONE },
    // The same room and a box that fits in it, which the rest read against.
    { name: "b", leftPt: 100, heightPt: 60, wrap: NONE },
    // A box whose foot falls past the sheet rather than past the text, since a page
    // has 36pt of margin under its text and a box may reach into it and no further.
    { name: "c", leftPt: 100, heightPt: 700, wrap: NONE },
    // The first case again with the wrap the corpus documents use, in case what a
    // box does with the text beside it decides this too.
    { name: "d", leftPt: 100, heightPt: BOX_PT, wrap: SQUARE },
    // A box taller than the whole body, which no page can hold however far it is
    // moved.
    { name: "e", leftPt: 600, heightPt: 900, wrap: NONE },
    // A wrap that takes the whole width with it rather than leaving a side, which
    // is the other wrap a real document writes.
    { name: "f", leftPt: 100, heightPt: BOX_PT, wrap: `<wp:wrapTopAndBottom/>` },
    // A wrapping box that fits in the room left, which tells a rule about the room
    // from a rule about the wrap.
    { name: "g", leftPt: 100, heightPt: 60, wrap: SQUARE },
    // A wrapping box reaching past the foot of the text and not past the sheet,
    // which says which of the two the room is measured to.
    { name: "h", leftPt: 100, heightPt: 120, wrap: SQUARE },
    // The same box hung well below its own paragraph rather than starting at it,
    // which is how a real document writes one: its objects sit 74pt under the
    // paragraph that anchors them and Word neither moved them on nor left them
    // where they fell.
    { name: "i", leftPt: 100, heightPt: 120, wrap: SQUARE, offsetPt: 74 },
    // And a box that will not fit anchored to the last paragraph there is, so there
    // is no page under it to be moved to.
    { name: "j", leftPt: 100, heightPt: 120, wrap: SQUARE, last: true },
    // A box hung below its paragraph with room enough that drawing its foot on the
    // bottom of the text still leaves its top below the anchor. In the real document
    // Word did exactly that rather than move anything, and `i` says it will not do
    // it where the box would have to rise above its own anchor to fit.
    { name: "k", leftPt: 156, heightPt: 120, wrap: SQUARE, offsetPt: 74 },
  ];

  return [...CASES.map((each, at) => block(each, at + 1)), EMPTY].join("");
}

// The same question asked of a page that draws a footer, and of a paragraph
// anchoring more than one object.
//
// `objects-past-the-foot` measured both of the rules this asks about on a page
// with no footer and one box at a time, and five corpus documents of one template
// contradict them together: a page of anchored boxes whose foot holds a square
// wrapped box a twentieth of a point past the top of the footer, a second that fits
// under it, and a third wrapping nothing that hangs past the sheet, all three at one
// anchor. Word moves that anchor to the next page. Neither measured rule moves
// anything there: the square one has room to rise and be drawn up to the foot, and
// the one wrapping nothing hangs where it was put however far past.
//
// So two things are unmeasured, and the cases below split them. **Which foot an
// object is judged against on a page that draws a footer**: the page here keeps
// 36pt of bottom margin and hangs a footer 24pt tall 36pt above the bottom edge, so
// the top of the footer at 732, the bottom margin at 756 and the edge of the sheet
// at 792 are three different answers and a box can be dropped between any two of
// them. And **whether several objects at one anchor answer as one does**: a box
// that cannot fit standing beside one that can, and one wrapping nothing beside a
// square one.
//
// Every case is written out three times so that an answer is a rule rather than an
// accident, and each repeat marks itself: two boxes holding the same words could
// not be told apart in a rendering.
export function objectsAndTheFooterDocument(): string {
  const LINE_PT = 24;
  const SHIM_PT = 500;

  const exactly = (pt: number): string =>
    `<w:spacing w:line="${String(pt * 20)}" w:lineRule="exact"/>`;

  const OWN_PAGE = `<w:pageBreakBefore/>`;

  const NONE = `<wp:wrapNone/>`;
  const SQUARE = `<wp:wrapSquare wrapText="bothSides"/>`;

  type Box = {
    readonly name: string;
    readonly wrap: string;
    readonly widthPt: number;
    readonly heightPt: number;
    // Where the box stands across the column, so that two at one anchor leave the
    // paragraph's own line a run of the frame to stand in rather than driving it
    // under both of them and off the page for a reason that is not being asked
    // about.
    readonly leftPt: number;
    readonly offsetPt: number;
  };

  const boxed = (of: Box, id: number): string =>
    `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
      <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${String(id)}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>${emu(of.leftPt)}</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(of.offsetPt)}</wp:posOffset></wp:positionV>
        <wp:extent cx="${emu(of.widthPt)}" cy="${emu(of.heightPt)}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        ${of.wrap}
        <wp:docPr id="${String(id)}" name="${of.name}"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:cNvSpPr txBox="1"/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(of.widthPt)}" cy="${emu(of.heightPt)}"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr>
            <wps:txbx><w:txbxContent>${paragraph(exactly(LINE_PT), run(`${of.name} boxed`))}</w:txbxContent></wps:txbx>
            <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"/>
          </wps:wsp></a:graphicData></a:graphic>
      </wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>`;

  const square = (heightPt: number, offsetPt: number, leftPt = 0): Omit<Box, "name"> => ({
    wrap: SQUARE,
    widthPt: 160,
    heightPt,
    leftPt,
    offsetPt,
  });

  const none = (heightPt: number, offsetPt: number, leftPt = 0): Omit<Box, "name"> => ({
    wrap: NONE,
    widthPt: 240,
    heightPt,
    leftPt,
    offsetPt,
  });

  // The anchor's own top, which every foot below is stated against: the marker
  // takes the first 24pt of the body and the shim the next 500.
  const ANCHOR_PT = TOP_PT + LINE_PT + SHIM_PT;

  type Case = {
    readonly name: string;
    readonly boxes: readonly (readonly [string, Omit<Box, "name">])[];
  };

  // A box standing at the anchor's own top and reaching down to a stated foot,
  // since what every case here is written around is where its foot falls: the top
  // of the footer stands at 732, the bottom margin at 756 and the edge of the sheet
  // at 792.
  const reachingTo = (footPt: number, leftPt = 0): Omit<Box, "name"> =>
    square(footPt - ANCHOR_PT, 0, leftPt);

  const CASES: readonly Case[] = [
    // Short of every foot there could be, which the rest are read against.
    { name: "a", boxes: [["", reachingTo(720)]] },
    // Past the top of the footer and short of the bottom margin.
    { name: "b", boxes: [["", reachingTo(740)]] },
    // Past the bottom margin and short of the sheet.
    { name: "c", boxes: [["", reachingTo(760)]] },
    // Past the sheet itself.
    { name: "d", boxes: [["", reachingTo(800)]] },
    // A box hung 100pt below its anchor with room to rise: Word draws such a one up
    // until its foot rests on the foot of the text, so where it comes to rest says
    // which foot that is without a page break having to be read at all.
    { name: "e", boxes: [["", square(150, 100)]] },
    // A box wrapping nothing, which on a page with no footer hangs past the foot
    // and moves nothing.
    { name: "f", boxes: [["", none(300, 0)]] },
    // The same box beside one that fits, which is the shape the corpus documents
    // hold and the one the measured rules cannot explain.
    {
      name: "g",
      boxes: [
        ["square", reachingTo(720)],
        ["none", none(300, 0, 200)],
      ],
    },
    // A box that cannot fit under any foot beside one that fits under all of them,
    // which says whether an anchor answers for its objects one at a time.
    {
      name: "h",
      boxes: [
        ["fits", reachingTo(720)],
        ["past", reachingTo(800, 380)],
      ],
    },
    // The corpus template itself: a square box hung 3pt past the top of the footer
    // and standing 7pt short of it, a second whose foot is at 722, and one wrapping
    // nothing whose foot is 76pt past the sheet.
    {
      name: "i",
      boxes: [
        ["rises", square(165, 10)],
        ["fits", square(150, 12, 380)],
        ["hangs", none(300, 8, 60)],
      ],
    },
  ];

  const REPEATS = [1, 2, 3];

  const block = (of: Case, repeat: number, from: number): string => {
    const mark = `${of.name}${String(repeat)}`;
    const boxes = of.boxes.map(([what, box], at) =>
      boxed({ ...box, name: what === "" ? mark : `${mark} ${what}` }, from + at),
    );
    return (
      paragraph(`${OWN_PAGE}${exactly(LINE_PT)}`, run(`${mark} marks`)) +
      paragraph(exactly(SHIM_PT), run(`${mark} shims`)) +
      paragraph(exactly(LINE_PT), boxes.join("") + run(`${mark} anchors`)) +
      paragraph(exactly(LINE_PT), run(`${mark} follows`))
    );
  };

  let id = 1;
  const blocks: string[] = [];
  for (const each of CASES) {
    for (const repeat of REPEATS) {
      blocks.push(block(each, repeat, id));
      id += each.boxes.length;
    }
  }

  return [...blocks, EMPTY].join("");
}

// A footer of a stated height, so that the top of it is a number the document
// itself states: a page 792pt tall keeping 36pt for a footer puts the top of a
// 24pt one at 732, which is 24pt above the bottom margin.
export const STATED_FOOTER = paragraph(
  `<w:spacing w:line="480" w:lineRule="exact"/>`,
  run("the footer"),
);

// What a break does to the line under it when there is nothing on that line.
//
// Two corpus documents of one converted template are 7 of 45 and 6 of 43 lines
// placed, and both are wrong from the first break down: one writes two breaks in a
// row in the middle of a paragraph and Word draws an empty line between them, the
// other ends a paragraph with a break and Word gives that break a line of its own.
// This project drew neither, since a line with nothing on it was no line at all.
//
// Every case is a paragraph told exactly how tall its lines are and written out three
// times, so the height it turned out to be is the distance from one repeat to the
// next rather than a difference of two rounded answers. Each repeat marks itself,
// since two lines of the same words cannot be told apart in a rendering.
export function breaksInAParagraphDocument(): string {
  const LINE_PT = 24;

  const exactly = (pt: number): string =>
    `<w:spacing w:before="0" w:after="0" w:line="${String(pt * 20)}" w:lineRule="exact"/>`;

  const OWN_PAGE = `<w:pageBreakBefore/>`;
  const BREAK = `<w:r><w:br/></w:r>`;

  const CASES: readonly (readonly [name: string, content: (mark: string) => string])[] = [
    // One line and no break at all, which every height below is read against.
    ["plain", (mark) => run(`${mark} one`)],
    // A break between two runs of text, which is the case that has always drawn two
    // lines.
    ["between", (mark) => run(`${mark} one`) + BREAK + run(`${mark} two`)],
    // A break with nothing after it but the paragraph's own mark.
    ["trailing", (mark) => run(`${mark} one`) + BREAK],
    // Two of them, which leaves no line for the second to end.
    ["two trailing", (mark) => run(`${mark} one`) + BREAK + BREAK],
    // A break with nothing in front of it.
    ["leading", (mark) => BREAK + run(`${mark} one`)],
    // Two breaks between two runs of text, which is what the first of the two corpus
    // documents holds.
    ["two between", (mark) => run(`${mark} one`) + BREAK + BREAK + run(`${mark} two`)],
    // And three, in case what a second break opens a third does not.
    ["three between", (mark) => run(`${mark} one`) + BREAK + BREAK + BREAK + run(`${mark} two`)],
    // A paragraph holding a break and nothing else.
    ["only a break", () => BREAK],
  ];

  const block = ([name, content]: readonly [string, (mark: string) => string]): string => {
    const mark = name.replace(/[^a-z]/g, "");
    return (
      paragraph(`${OWN_PAGE}${exactly(LINE_PT)}`, run(`${mark} marks`)) +
      [1, 2, 3].map((at) => paragraph(exactly(LINE_PT), content(`${mark}${String(at)}`))).join("") +
      paragraph(exactly(LINE_PT), run(`${mark} follows`))
    );
  };

  return [...CASES.map(block), EMPTY].join("");
}

// How much a justified line will be squeezed to take one more word.
//
// 118 of the 718 corpus documents justify a paragraph, and between them they hold
// 15407 lines and miss 5512 of them, against 78.9% placed over every document that
// justifies nothing. Reading one of them says why: Word draws a justified line
// holding text whose own width is **wider than the line**, by 3.8pt over 14 spaces
// in one case and 4.8pt over 13 in another. This project fits a line at its natural
// width and never squeezes it, so it breaks a word earlier and the whole paragraph
// is a line long.
//
// Sixty seven cases, each written out three times, each a justified paragraph
// whose last word overflows the room by a stated amount. Where Word draws that word
// says what it will accept, and the three digits opening every case say which case
// and which repeat a line belongs to.
//
// The numbers are this project's own measure of Calibri at 24pt, which the authored
// suite pins to the hundredth elsewhere: `mmmm` is 76.6875pt, `mmmmmmmm` 153.375,
// `aa` 22.9922, a space 5.4258, and the three digits in front of every case 36.4921.
// So a case of four `mmmm` is 364.9453pt of text over four spaces, one of twelve `aa`
// is 377.5078 over twelve and one of three `mmmmmmmm` is 512.8945 over three, and the
// room each is given is that width less the overflow it is being asked about.
export function justifiedFittingDocument(): string {
  const OWN_PAGE = `<w:pageBreakBefore/>`;

  // Enough words after the case's own to fill a second line, so the line being asked
  // about is never the last one: Word justifies every line of a paragraph but that.
  const FILLER = "zzzz zzzz zzzz zzzz zzzz zzzz zzzz zzzz";

  type Family = {
    readonly name: string;
    readonly word: string;
    readonly words: number;
    readonly sizePt: number;
    // A face of its own where the question is about the face: Times New Roman makes
    // its space a quarter of the em where Calibri makes it 0.2261, so the two say
    // differently whether the ceiling below is a length of the em or of the space.
    readonly face?: string;
    // The room that leaves the line exactly the width of its own text, in twips off
    // the right of a 540pt frame.
    readonly indentTwips: number;
    // How many characters the line holds, which is what the answer turned out to
    // turn on.
    readonly characters: number;
    // What the room is narrowed by, in points, against the width of the whole first
    // line. A twip is a twentieth of a point, so the closest two of these are two
    // twips apart.
    readonly overflowsPt: readonly number[];
  };

  // Three shapes of line at 24pt and one of them again at 12pt. The first two ask
  // the same overflow of four spaces and of twelve, the third asks three spaces to
  // hold it on a line half again as wide, and the fourth asks whether the answer is
  // a length or a fraction of the size the text is set in. Each sweep is close
  // around where the line stopped taking the word, and holds one wide case either
  // side of it for the record.
  // The first four are the shapes that killed every rule of one term: four spaces,
  // twelve, three on a line half again as wide, and the first again at half the
  // size, each holding the widest overflow Word took and the narrowest it refused.
  // The six after them vary nothing but how many spaces the line holds, since that
  // is the axis a rule of one term and a rule of two disagree on: a rule per space
  // rises straight from three spaces to twelve, and one that also holds a floor of
  // its own is flat until the spaces are worth more than the floor.
  const FAMILIES: readonly Family[] = [
    {
      name: "few",
      word: "mmmm",
      words: 4,
      sizePt: 24,
      indentTwips: 3502,
      characters: 23,
      overflowsPt: [4.6, 6],
    },
    {
      name: "many",
      word: "aa",
      words: 12,
      sizePt: 24,
      indentTwips: 3251,
      characters: 39,
      overflowsPt: [7.6, 12],
    },
    {
      name: "wide",
      word: "mmmmmmmm",
      words: 3,
      sizePt: 24,
      indentTwips: 543,
      characters: 30,
      overflowsPt: [4, 5],
    },
    {
      name: "small",
      word: "mmmm",
      words: 4,
      sizePt: 12,
      indentTwips: 7151,
      characters: 23,
      overflowsPt: [2.2, 2.3],
    },
    {
      name: "three",
      word: "aa",
      words: 3,
      sizePt: 24,
      indentTwips: 8365,
      characters: 12,
      overflowsPt: [4, 4.4],
    },
    {
      name: "four",
      word: "aa",
      words: 4,
      sizePt: 24,
      indentTwips: 7797,
      characters: 15,
      overflowsPt: [5.2, 5.6, 6, 6.4],
    },
    {
      name: "six",
      word: "aa",
      words: 6,
      sizePt: 24,
      indentTwips: 6660,
      characters: 21,
      overflowsPt: [6, 7, 8, 8.4, 8.8, 9.6],
    },
    {
      name: "eight",
      word: "aa",
      words: 8,
      sizePt: 24,
      indentTwips: 5523,
      characters: 27,
      overflowsPt: [10.2, 10.4, 10.6, 10.8, 11],
    },
    {
      name: "ten",
      word: "aa",
      words: 10,
      sizePt: 24,
      indentTwips: 4387,
      characters: 33,
      overflowsPt: [10.1, 10.2, 10.4, 10.6],
    },
    {
      name: "half",
      word: "aa",
      words: 12,
      sizePt: 12,
      indentTwips: 7025,
      characters: 39,
      overflowsPt: [4.6, 5, 5.2, 5.4, 6, 7, 8],
    },
    {
      name: "sixteen",
      word: "aa",
      words: 16,
      sizePt: 24,
      indentTwips: 976,
      characters: 51,
      overflowsPt: [10, 10.2, 10.4, 10.6, 11],
    },
    {
      name: "sixteen wide",
      word: "aa",
      words: 16,
      sizePt: 24,
      indentTwips: 976,
      characters: 51,
      overflowsPt: [9.5, 9.6, 9.7, 9.8, 9.9],
    },
    {
      name: "sixteen small",
      word: "aa",
      words: 16,
      sizePt: 12,
      indentTwips: 5888,
      characters: 51,
      overflowsPt: [4, 4.6, 4.8, 5, 5.4],
    },
    {
      name: "twenty four small",
      word: "aa",
      words: 24,
      sizePt: 12,
      indentTwips: 3615,
      characters: 75,
      overflowsPt: [4, 4.6, 4.8, 5, 5.4],
    },
    {
      name: "roman",
      word: "aa",
      words: 12,
      sizePt: 24,
      face: "Times New Roman",
      indentTwips: 3527,
      characters: 39,
      overflowsPt: [9.6, 10, 10.4, 10.8, 11.2, 11.6],
    },
    {
      name: "twelve",
      word: "aa",
      words: 12,
      sizePt: 24,
      indentTwips: 3250,
      characters: 39,
      overflowsPt: [10.1, 10.2, 10.4, 10.6],
    },
  ];

  const cases = FAMILIES.flatMap((family) =>
    family.overflowsPt.map((overflowPt) => ({ family, overflowPt })),
  );

  return [
    ...cases.flatMap(({ family, overflowPt }, at) => {
      const number = String(at + 1).padStart(2, "0");
      const face =
        family.face === undefined
          ? ""
          : `<w:rFonts w:ascii="${family.face}" w:hAnsi="${family.face}" w:cs="${family.face}"/>`;
      const size = `${face}<w:sz w:val="${String(family.sizePt * 2)}"/><w:szCs w:val="${String(family.sizePt * 2)}"/>`;
      const indent = `<w:ind w:right="${String(family.indentTwips + Math.round(overflowPt * 20))}"/>`;
      const words = Array.from({ length: family.words }, () => family.word).join(" ");
      return [
        paragraph(OWN_PAGE, run(`${family.name} short of ${String(overflowPt)} above`)),
        ...[1, 2, 3].map((repeat) =>
          paragraph(
            `<w:jc w:val="both"/>${indent}`,
            run(`${number}${String(repeat)} ${words} ${FILLER}`, size),
          ),
        ),
      ];
    }),
    EMPTY,
  ].join("");
}

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

// What a section break does, and which section's page the text either side of one
// is drawn on.
//
// Two things are unknown and neither is worth guessing. **Which section a
// `w:sectPr` on a paragraph governs**: the paragraph's own, or the one after it.
// And **what `w:type` describes**: how the section holding it begins, or how the
// one after it does. The spec is read both ways.
//
// Every section here keeps a left margin of its own and every paragraph says which
// section wrote it, so where a line is drawn across the page answers the first
// question and which page it lands on answers the second. **Read off Word's own
// pdf**, not its report: the report gives a paragraph's position from its own text
// boundary, which is the same number whatever the margin is.
export function sectionsDocument(): string {
  // Far enough apart that no rounding puts a line under the wrong one, and none of
  // them the 36pt the body's own section keeps.
  const ONE_INCH = 1440;
  const TWO_INCHES = 2880;
  const THREE_INCHES = 4320;

  const sectionProperties = (leftTwips: number, type: string): string =>
    `<w:sectPr>` +
    (type === "" ? "" : `<w:type w:val="${type}"/>`) +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="${String(leftTwips)}" w:header="720" w:footer="720" w:gutter="0"/>` +
    `</w:sectPr>`;

  // The paragraph carrying section properties is the last of whichever section they
  // turn out to govern, which is what this document is for.
  const closes = (name: string, leftTwips: number, type: string): string =>
    paragraph(sectionProperties(leftTwips, type), run(name));

  const line = (name: string): string => paragraph("", run(name));

  return [
    line("one opens"),
    closes("one closes", ONE_INCH, ""),
    line("two opens"),
    closes("two closes", TWO_INCHES, "continuous"),
    line("three opens"),
    closes("three closes", THREE_INCHES, "nextPage"),
    // The last section is the body's own, at the 36pt margin the rest of the suite
    // uses, so a line drawn there is one no paragraph's properties reached.
    line("four opens"),
    line("four closes"),
  ].join("");
}

// Which page a section opens on, asked of every `w:type` there is.
//
// The `sections` document says a continuous break opens no page and a `nextPage`
// one does, and it cannot be held to that: every section in it keeps a margin of
// its own, so nothing here can lay it out and the answer sits in prose. Every
// section here is the page the body's own section is, to the twip, so the only
// thing a page number can be about is the break, and the document lays out the
// moment the break is read.
//
// A section's `w:type` says how that section begins against the one before it, so
// the type asked about goes on the section whose opening is being watched. The
// first section's own type is therefore about nothing, and says `continuous` to
// say so, which is what the documents in the wild that raised this question write
// there.
//
// The three types nothing has measured are at the end, so that a page they turn
// out to open and this project does not costs the answers after it and no others.
export function sectionPagesDocument(): string {
  // The body's own page, written out again: a section stating anything else would
  // be asking a second question and there is no way to lay it out yet.
  const sectionProperties = (type: string): string =>
    `<w:sectPr>` +
    (type === "" ? "" : `<w:type w:val="${type}"/>`) +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>` +
    `<w:cols w:space="720"/>` +
    `</w:sectPr>`;

  const closes = (name: string, type: string): string =>
    paragraph(sectionProperties(type), run(`${name} closes`));

  const opens = (name: string): string => paragraph("", run(`${name} opens`));

  const section = (name: string, type: string): readonly string[] => [
    opens(name),
    closes(name, type),
  ];

  return [
    ...section("one", "continuous"),
    // Stating no type at all, which is what a document writes far more often than
    // it writes `nextPage`, and the case a three page document in the wild turns on.
    ...section("two", ""),
    ...section("three", "continuous"),
    ...section("four", "nextPage"),
    ...section("five", "continuous"),
    ...section("six", "nextColumn"),
    ...section("seven", "evenPage"),
    ...section("eight", "oddPage"),
    // The last section is the body's own, which states no type either.
    opens("nine"),
    paragraph("", run("nine closes")),
  ].join("");
}

// What room the paragraph carrying a section break takes at the foot of the page it
// ends, which `section-pages` never asks: every closer in that one stands near the
// top of a page with the whole of it below.
//
// A document in the wild closes a section with an empty paragraph six points past
// the foot of its page. Read as an ordinary paragraph its mark will not fit, so it
// moves on and the break under it then opens a second page, and the document came
// out a blank page longer than Word drew it. So the question is whether the mark of
// a paragraph that does nothing but carry a break is laid out at all.
//
// Each case opens a page of its own and is held down it by a shim told exactly how
// tall to be, so the room left under it is arithmetic: the page is 720pt of body,
// the marker takes 24 and the shim the rest. The page the paragraph after a closer
// lands on is the reading, since Word's answer for a closer is not its own: an
// empty one comes back at the top of the paragraph above it, which it reports
// whether the closer fits or not and is therefore about neither.
//
// Two questions, and the first needs the closer past the foot of the page while the
// second needs it nowhere near one. **Whether a closer moves on where its line will
// not fit**, read off `a` against `c`. And **whether it takes room where there is
// room**, read off `e`, whose follower stands one line lower for every line the
// closer took.
//
// A break is read at the section it opens, so what decides whether a page follows a
// closer is the type the **next** closer states. `a`, `b` and `c` each want one,
// `d` and `e` want none, and the closer under each is what says so.
export function sectionCloserDocument(): string {
  const exactly = (pt: number): string =>
    `<w:spacing w:line="${String(pt * 20)}" w:lineRule="exact"/>`;

  const OWN_PAGE = `<w:pageBreakBefore/>`;
  const LINE_PT = 24;
  const BODY_PT = 720;

  const sectionProperties = (type: string): string =>
    `<w:sectPr>` +
    (type === "" ? "" : `<w:type w:val="${type}"/>`) +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>` +
    `<w:cols w:space="720"/>` +
    `</w:sectPr>`;

  // `leftPt` is the room standing under the shim, which is what the closer is
  // offered. Less than a line of it is the case that tells the two readings apart.
  const block = (
    name: string,
    leftPt: number,
    type: string,
    closerText: string,
  ): readonly string[] => [
    paragraph(`${OWN_PAGE}${exactly(LINE_PT)}`, run(`${name} marks`)),
    paragraph(exactly(BODY_PT - LINE_PT - leftPt), run(`${name} shims`)),
    paragraph(`${exactly(LINE_PT)}${sectionProperties(type)}`, closerText),
    paragraph(exactly(LINE_PT), run(`${name} follows`)),
  ];

  const CLOSER_FITS_PT = 36;
  const CLOSER_DOES_NOT_PT = 12;
  // Ten lines of room under the shim, so nothing in `e` is near a page break and
  // the only thing that can move its follower is the closer above it.
  const ROOM_TO_SPARE_PT = 240;

  return [
    // Half a line of room and an empty closer, which is the case in the wild.
    ...block("a", CLOSER_DOES_NOT_PT, "", ""),
    // A line and a half of room, so the closer fits however it is read.
    ...block("b", CLOSER_FITS_PT, "", ""),
    // Half a line of room and a closer with text in it, which has to be drawn
    // somewhere whatever the answer about a bare mark turns out to be.
    ...block("c", CLOSER_DOES_NOT_PT, "", run("c closes")),
    // Half a line of room again, and no page opening after it.
    ...block("d", CLOSER_DOES_NOT_PT, "", ""),
    // Room to spare and no page opening after it, so the follower stands on the
    // closer's own page and says how much of it the closer took.
    ...block("e", ROOM_TO_SPARE_PT, "continuous", ""),
    // Closing nothing anyone is reading, and stating the type that keeps `d` and
    // `e` on the pages they were already on.
    paragraph(sectionProperties("continuous"), run("f closes")),
    // The body's own section governs whatever follows the last closer, and Word
    // writes a paragraph of its own there when the document offers none.
    paragraph("", run("g closes")),
  ].join("");
}

// Where a section's text sits down the page, which the `sections` document cannot
// say: every margin in that one is the same at the top, so the question of whether
// a section's own top margin reaches its text never arises there.
//
// Two things are asked. **What a continuous break does with the top margin of the
// section it opens**, since that section opens no page and there is no top of one
// for the margin to be measured from: either the text carries straight on under the
// line above it, or the margin puts it somewhere of its own. And **whether a
// section opening a page keeps its own top margin there**, which is the plain case
// and is worth pinning beside the other.
//
// Read off Word's own pdf, as `sections` is, and for the same reason.
export function sectionFlowDocument(): string {
  const HALF_INCH = 720;
  const TWO_INCHES = 2880;
  const THREE_INCHES = 4320;

  const sectionProperties = (topTwips: number, type: string): string =>
    `<w:sectPr>` +
    (type === "" ? "" : `<w:type w:val="${type}"/>`) +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="${String(topTwips)}" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>` +
    `</w:sectPr>`;

  const closes = (name: string, topTwips: number, type: string): string =>
    paragraph(sectionProperties(topTwips, type), run(name));

  const line = (name: string): string => paragraph("", run(name));

  return [
    // Half an inch down, which is where every authored page starts.
    line("one a"),
    closes("one b", HALF_INCH, ""),
    // Three inches down, and continuous. Text carrying straight on under "one b"
    // says the margin is not reached at all; text three inches down says it is.
    line("two a"),
    closes("two b", THREE_INCHES, "continuous"),
    // Two inches down, opening a page of its own, where there is a top of a page
    // for a margin to be measured from.
    line("three a"),
    closes("three b", TWO_INCHES, "nextPage"),
    // And the body's own, which every authored document keeps half an inch down,
    // so its text reads against the first section's rather than against a margin
    // of its own.
    line("four a"),
    line("four b"),
  ].join("");
}

// Which page a continuous section's own text is drawn on when it runs past the foot
// of the page it opened on: the page the run opened with, or the section's own.
//
// A continuous section opens no page, so its text starts under the line above it on
// whatever page that was, and the page it started on was made by an earlier section.
// When its text runs on there is a page nothing has stated the geometry of, and the
// two readings differ by the whole of the margins. This is the one thing about a
// section nothing had put to Word, and the page geometry has to move off the document
// and onto the page either way, so it decides what a page is made of before anything
// is built on it.
//
// The two sections are three inches apart in their left margins and an inch and a
// half apart at the top, so where a line is drawn says which of them governed the
// page it stands on. Read off Word's own pdf: the report gives a paragraph's position
// from its own text boundary, which is the same number whatever the margin is.
export function overflowingSectionDocument(): string {
  const LINE_PT = 24;

  const exactly = `<w:spacing w:line="${String(LINE_PT * 20)}" w:lineRule="exact"/>`;

  const sectionProperties = (topTwips: number, leftTwips: number, type: string): string =>
    `<w:sectPr>` +
    (type === "" ? "" : `<w:type w:val="${type}"/>`) +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="${String(topTwips)}" w:right="720" w:bottom="720" w:left="${String(leftTwips)}" w:header="720" w:footer="720" w:gutter="0"/>` +
    `<w:cols w:space="720"/>` +
    `</w:sectPr>`;

  const line = (name: string): string => paragraph(exactly, run(name));

  const closes = (name: string, topTwips: number, leftTwips: number, type: string): string =>
    paragraph(`${exactly}${sectionProperties(topTwips, leftTwips, type)}`, run(name));

  // Enough to fill the page it opens on twice over, so there are two pages of it
  // beyond the one the run began with and a page that ran on from a page that ran on
  // is read as well as the first.
  const RUNS_ON = 60;

  return [
    // An inch across and half an inch down, which is where the run opens.
    line("one a"),
    closes("one b", 720, 1440, ""),
    // Four inches across and two inches down, and continuous, so this section opens
    // no page of its own and every page its text reaches is one it did not make.
    line("two a"),
    ...Array.from({ length: RUNS_ON }, (_, at) => line(`two ${String(at + 1)}`)),
    closes("two z", 2880, 5760, "continuous"),
    // And the body's own, which every authored document keeps half an inch down and
    // half an inch across, so a line drawn there is one neither section reached.
    line("three a"),
    line("three b"),
  ].join("");
}

// What a run raised or lowered off its own baseline does to the line it stands on.
//
// `w:position` is read by the fidelity report and by nothing else, so a run stating
// one is drawn flat on the baseline today. Two things are in question and only one
// of them is about the run itself: **how far off the baseline Word draws it**, and
// **whether the line grows to hold it**, which is what puts every line below it in
// the wrong place rather than only the run.
//
// The values asked are the ones the wild states. Of the 2783 `w:position` elements
// in the corpus, 2189 are -1 and 174 are 5: a run half a point down and a run two
// and a half points up. A raise of twelve half-points is asked beside them, since a
// rule too small to read at a half-point is unmistakable at six.
//
// Eleven cases, each opening a page of its own, and the paragraph of each written
// out three times so that the height of a line is the distance between one repeat
// and the next rather than a difference of two rounded answers. A plain line closes
// every case, which says where the last of the three left the page.
//
// **A raised run standing beside a plain one is what says where the line's own
// baseline is**: Word draws the two as items of its own, so the drawing gives the
// raise and the line at once. The cases where the raised run stands alone answer a
// different question, which is what a line does when nothing on it is on the
// baseline at all.
export function raisedTextDocument(): string {
  const OWN_PAGE = `<w:pageBreakBefore/>`;
  const SUPERSCRIPT = `<w:vertAlign w:val="superscript"/>`;
  const SUBSCRIPT = `<w:vertAlign w:val="subscript"/>`;

  const raisedBy = (halfPoints: number): string => `<w:position w:val="${String(halfPoints)}"/>`;

  type Case = {
    readonly name: string;
    // Word's own unit, a half-point, and its own sign: positive lifts the run.
    readonly halfPoints: number;
    // Whether a plain run opens the line in front of the raised one.
    readonly besideAPlainRun: boolean;
    // What the paragraphs of the case say about their own lines, where they say
    // anything: left out, they take the single line the document's defaults set.
    readonly spacing?: string;
    // Whether the raised run is a script as well, which Word already moves off the
    // baseline and shrinks on its own.
    readonly superscript?: boolean;
    readonly subscript?: boolean;
    // What the raised run is set in, where that is not the document's own 12pt.
    readonly sizeHalfPoints?: number;
  };

  const CASES: readonly Case[] = [
    // Nothing raised at all, which is the height every other case is read against.
    { name: "a", halfPoints: 0, besideAPlainRun: true },
    // The commonest value in the corpus, and the smallest: half a point down.
    { name: "b", halfPoints: -1, besideAPlainRun: true },
    // The second commonest: two and a half points up.
    { name: "c", halfPoints: 5, besideAPlainRun: true },
    // Six points either way, which no rounding can hide.
    { name: "d", halfPoints: 12, besideAPlainRun: true },
    { name: "e", halfPoints: -12, besideAPlainRun: true },
    // The same six points with nothing on the baseline beside them. If a line takes
    // its height off a raised run's own ascent and descent, one holding nothing else
    // is as tall as it ever was and simply sits higher.
    { name: "f", halfPoints: 12, besideAPlainRun: false },
    { name: "g", halfPoints: -12, besideAPlainRun: false },
    // Twice the raise, alone: whatever f does, this says whether it goes on doing it.
    { name: "h", halfPoints: 24, besideAPlainRun: false },
    // The same raise under a line told exactly how tall to be, which is where a
    // growing line has nowhere to grow.
    {
      name: "i",
      halfPoints: 12,
      besideAPlainRun: true,
      spacing: `<w:spacing w:line="480" w:lineRule="exact"/>`,
    },
    // And under a floor low enough that the raise reaches past it: 16pt against a
    // plain line of about 14.6.
    {
      name: "j",
      halfPoints: 12,
      besideAPlainRun: true,
      spacing: `<w:spacing w:line="320" w:lineRule="atLeast"/>`,
    },
    // A superscript raised on top of what Word already raises it by, which says
    // whether the two add and whether the shrunk size is what the raise is measured
    // against.
    { name: "k", halfPoints: 12, besideAPlainRun: true, superscript: true },
    // Lowered further than a 12pt line reaches above its own baseline, alone: the
    // other side of whatever f does when a raise runs past the descent.
    { name: "l", halfPoints: -30, besideAPlainRun: false },
    // A superscript and a subscript asking for no raise of their own, which is what
    // says how far Word moves each and what that does to the line. This project
    // moves both by a third of the size, which was never measured, and a raise it
    // gets wrong is now a line height as well as a drawing.
    { name: "n", halfPoints: 0, besideAPlainRun: true, superscript: true },
    { name: "o", halfPoints: 0, besideAPlainRun: true, subscript: true },
    // The same two at twice the size, which says whether either is a share of it.
    { name: "p", halfPoints: 0, besideAPlainRun: true, superscript: true, sizeHalfPoints: 48 },
    { name: "q", halfPoints: 0, besideAPlainRun: true, subscript: true, sizeHalfPoints: 48 },
    // And a subscript with nothing on the baseline beside it, where whatever it does
    // to the line has nothing to hide behind.
    { name: "r", halfPoints: 0, besideAPlainRun: false, subscript: true },
    // A third size of each, since two points fix a line and three say it is one.
    { name: "s", halfPoints: 0, besideAPlainRun: true, superscript: true, sizeHalfPoints: 72 },
    { name: "t", halfPoints: 0, besideAPlainRun: true, subscript: true, sizeHalfPoints: 72 },
    // Six points up under a line and a half, which is where a multiple decides
    // whether the raise is taken before it or after: a line of 14.64 grown to 20.64
    // comes out 30.96 taken before and 27.96 taken after.
    {
      name: "m",
      halfPoints: 12,
      besideAPlainRun: true,
      spacing: `<w:spacing w:line="360" w:lineRule="auto"/>`,
    },
  ];

  const ORDINALS = ["one", "two", "three"] as const;

  const block = (of: Case): readonly string[] => {
    const properties = of.spacing ?? "";
    const size =
      of.sizeHalfPoints === undefined ? "" : `<w:sz w:val="${String(of.sizeHalfPoints)}"/>`;
    const script =
      (of.superscript === true ? SUPERSCRIPT : "") + (of.subscript === true ? SUBSCRIPT : "");
    const marks = `${size}${script}${raisedBy(of.halfPoints)}`;

    const repeat = (ordinal: string): string =>
      paragraph(
        properties,
        (of.besideAPlainRun ? run(`${of.name} ${ordinal} `) : "") + run("shifted", marks),
      );

    return [
      paragraph(`${OWN_PAGE}${properties}`, run(`${of.name} marks`)),
      ...ORDINALS.map(repeat),
      paragraph(properties, run(`${of.name} after`)),
    ];
  };

  return [...CASES.flatMap(block), EMPTY].join("");
}

// What becomes of the room a paragraph asks for above itself when a page begins
// under it, which depends on what began the page.
//
// `breaking` says a paragraph the foot of a page carried on to draws its first line
// at the top of the next page with none of that room above it, and this project has
// been dropping the room at every break there is. Two corpus documents of five
// sections each say that is too wide a rule: every page a section break opens in
// them stands exactly the room the paragraph opening it asks for below where this
// project puts it, 3.4pt at one of the breaks and 3.95pt at two more.
//
// So what is in question is the break rather than the room. Four kinds of them open
// a page here: the foot of the page forcing one, a break inside the text, a
// paragraph asking for a page of its own, and a section break. Each is written once
// asking for 18pt above the paragraph and once asking for none, so the answer is the
// distance between two lines Word drew rather than a line read against a font's
// ascent, and each case is written out three times: the second and the third stand
// mid page where the room is certainly kept, so the distance between one and the
// next gives the line's own height beside it.
//
// **Read off Word's own pdf**, which is where a line was drawn. Word's report
// answers for the paragraph, and a paragraph begins above whatever room it asked
// for, which is the same number whether the page kept that room or not.
export function spaceAboveABreakDocument(): string {
  const exactly = (pt: number): string =>
    `<w:spacing w:line="${String(pt * 20)}" w:lineRule="exact"/>`;
  const roomAbove = (twips: number): string => `<w:spacing w:before="${String(twips)}"/>`;

  const OWN_PAGE = `<w:pageBreakBefore/>`;
  const PAGE_BREAK = `<w:r><w:br w:type="page"/></w:r>`;

  // The body's own page, written out again: every section here is the page the
  // document's own is, to the twip, so the only thing a break can be about is the
  // break.
  const sectionProperties =
    `<w:sectPr>` +
    `<w:type w:val="nextPage"/>` +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>` +
    `<w:cols w:space="720"/>` +
    `</w:sectPr>`;

  // Room the paragraph could not be standing in by accident, against the 3.4pt and
  // 3.95pt the documents in the wild ask for.
  const ROOM_TWIPS = 360;

  const MARKER_PT = 24;
  const FILLERS = 9;
  const FILLER_PT = 72;
  const BODY_PT = 720;
  // Twelve points of room under the shim, which is less than the 14.65pt line the
  // case needs whether or not it asks for room above itself.
  const SHIM_PT = BODY_PT - MARKER_PT - FILLERS * FILLER_PT - 12;

  // How the page the case opens on was opened.
  type Opening = "the page filling" | "a break in the text" | "a page of its own" | "a section";

  const block = (name: string, opening: Opening, roomTwips: number): readonly string[] => {
    const marker = paragraph(
      `${OWN_PAGE}${exactly(MARKER_PT)}${opening === "a section" ? sectionProperties : ""}`,
      run(`${name} above`) + (opening === "a break in the text" ? PAGE_BREAK : ""),
    );
    const fills =
      opening !== "the page filling"
        ? []
        : [
            ...Array.from({ length: FILLERS }, () =>
              paragraph(exactly(FILLER_PT), run(`${name} fills`)),
            ),
            paragraph(exactly(SHIM_PT), run(`${name} shims`)),
          ];
    const cases = ["one", "two", "three"].map((which, at) =>
      paragraph(
        `${at === 0 && opening === "a page of its own" ? OWN_PAGE : ""}${roomAbove(roomTwips)}`,
        run(`${name} ${which}`),
      ),
    );
    return [marker, ...fills, ...cases];
  };

  return [
    ...block("soft", "the page filling", ROOM_TWIPS),
    ...block("soft bare", "the page filling", 0),
    ...block("hard", "a break in the text", ROOM_TWIPS),
    ...block("hard bare", "a break in the text", 0),
    ...block("mark", "a page of its own", ROOM_TWIPS),
    ...block("mark bare", "a page of its own", 0),
    ...block("section", "a section", ROOM_TWIPS),
    ...block("section bare", "a section", 0),
  ].join("");
}

// A table positioned rather than flowed, which is what `w:tblpPr` asks for.
//
// Seven corpus documents state one and they miss 459 lines between them, 300 of
// them in the documents needing no face stood in, which is the second largest thing
// left after the columns. Three of those seven are among the eight this project
// could not explain at all, and the ranking calls the whole cluster `merged-cells`,
// which is stated by four documents and explains nothing about any of them.
//
// **The corpus says two things are out and neither has been put to Word.** One
// document's table takes 165pt of the flow here and none of it there: twelve empty
// paragraphs standing under it are drawn beside it instead, and the paragraph after
// them lands where the flow would have put it had the table never been in it.
// Another's is drawn 299.4pt to the right of where this project puts it and 3.25pt
// above, which is what `w:tblpXSpec="right"` and `w:tblpY="-65"` would come to.
//
// Every case is the same two rows of two cells, each row told exactly how tall to
// be, so the table is 28.8pt tall and 144pt wide whatever a case does with it, and
// every flow line is told exactly how tall to be as well: the marker takes 24pt from
// the top of the body at 36, so a flow that closed over the table draws its next
// line's baseline at a place arithmetic gives. Each case's line after the table is
// written out three times, so the distance between one repeat and the next says the
// line height the case was drawn at rather than leaving it to be inferred.
//
// **Read off Word's own pdf.** Word's report answers for a paragraph in a table with
// the origin of the row rather than the cell, and for the horizontal with nought,
// which is the whole of what is being asked here.
export function positionedTableDocument(): string {
  const LINE_PT = 24;
  const ROW_PT = 14.4;
  const CELL_TWIPS = 1440;

  const exactly = (pt: number): string =>
    `<w:spacing w:line="${String(pt * 20)}" w:lineRule="exact"/>`;
  const OWN_PAGE = `<w:pageBreakBefore/>`;

  // Stated because an authored document declares no table style, and a table whose
  // cells state no margin of their own is then held off its walls by nothing.
  const CELL_MARGINS =
    `<w:tblCellMar>` +
    `<w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>` +
    `<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>` +
    `</w:tblCellMar>`;

  const cell = (text: string): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${String(CELL_TWIPS)}" w:type="dxa"/></w:tcPr>` +
    paragraph(exactly(ROW_PT), run(text)) +
    `</w:tc>`;

  const row = (left: string, right: string): string =>
    `<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="${String(ROW_PT * 20)}"/></w:trPr>` +
    cell(left) +
    cell(right) +
    `</w:tr>`;

  // `positioning` is the whole of `w:tblpPr`, or nothing at all for the case that
  // asks what the same table does when it is left in the flow.
  const table = (name: string, positioning: string): string =>
    `<w:tbl><w:tblPr>${positioning}` +
    `<w:tblW w:w="${String(CELL_TWIPS * 2)}" w:type="dxa"/>` +
    CELL_MARGINS +
    `</w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="${String(CELL_TWIPS)}"/><w:gridCol w:w="${String(CELL_TWIPS)}"/></w:tblGrid>` +
    row(`${name} a`, `${name} b`) +
    row(`${name} c`, `${name} d`) +
    `</w:tbl>`;

  const positioned = (properties: string): string => `<w:tblpPr ${properties}/>`;

  const block = (name: string, positioning: string, after = ""): readonly string[] => [
    paragraph(`${OWN_PAGE}${exactly(LINE_PT)}`, run(`${name} above`)),
    table(name, positioning),
    ...["one", "two", "three"].map((which) =>
      paragraph(
        exactly(LINE_PT),
        run(after === "" ? `${name} ${which}` : `${name} ${which} ${after}`),
      ),
    ),
  ];

  // An inch across, which is neither the margin the page keeps nor the edge of the
  // sheet, so the three anchors cannot be confused with one another.
  const ACROSS = 1440;

  return [
    // The same table left in the flow, which says what the flow does with its room
    // and what the cases below are read against.
    ...block("flowed", ""),
    // Anchored to the text and placed an inch from the column, which is where a
    // document stating no `w:horzAnchor` puts it.
    ...block("column", positioned(`w:vertAnchor="text" w:tblpX="${String(ACROSS)}"`)),
    // The same inch from the edge of the sheet, which is half an inch further left
    // than the column's if the anchor is read.
    ...block(
      "page",
      positioned(`w:vertAnchor="text" w:horzAnchor="page" w:tblpX="${String(ACROSS)}"`),
    ),
    // And the same inch from the margin, which one column makes the same place as
    // the column's own: what this separates is the page from the other two.
    ...block(
      "margin",
      positioned(`w:vertAnchor="text" w:horzAnchor="margin" w:tblpX="${String(ACROSS)}"`),
    ),
    // Asked for the right of the margin rather than a distance, which is what three
    // of the seven documents state.
    ...block("right", positioned(`w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="right"`)),
    // Lifted 18pt off wherever the text anchor puts it, which is the other half of
    // what those three state.
    ...block(
      "lifted",
      positioned(`w:vertAnchor="text" w:tblpX="${String(ACROSS)}" w:tblpY="-360"`),
    ),
    // Held at the column's own left with 9pt kept clear either side, and followed by
    // text long enough to reach it: either the text stands beside the table or it
    // runs under it.
    ...block(
      "beside",
      positioned(`w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:tblpX="0"`),
      FLOW,
    ),
    EMPTY,
  ].join("");
}

// What becomes of the room a paragraph asks for above itself when a wrap has
// already pushed its first line down the page.
//
// Two readings, and a real document turns on them. **The room is absorbed**: the
// line goes to the foot of the band and the space before it counted for nothing,
// which is what this project does. Or **the room is kept**: the line goes that far
// below the foot of the band, and a paragraph asking for ten points above itself
// starts ten points lower than the object under which it stands.
//
// Two corpus documents of one template are 9.8pt out from the first object on their
// first page all the way down, and the paragraph under that object states 211 twips
// of space before it, which is 10.55pt. Word draws the object itself exactly where
// this project puts it, so nothing about the object is in question: only what the
// paragraph under it does with the room it asked for.
//
// Every case opens a page of its own, and every line is told exactly how tall to be,
// so where a line lands is arithmetic: the body starts 36pt down, the marker takes
// 24 and the anchoring paragraph is empty and takes 24 more.
export function spaceUnderAWrapDocument(): string {
  const LINE_PT = 24;
  const BOX_PT = 200;
  const OWN_PAGE = `<w:pageBreakBefore/>`;

  const exactly = (pt: number): string =>
    `<w:spacing w:line="${String(pt * 20)}" w:lineRule="exact"/>`;
  const above = (pt: number, linePt: number): string =>
    `<w:spacing w:before="${String(pt * 20)}" w:line="${String(linePt * 20)}" w:lineRule="exact"/>`;

  const boxed = (id: number, name: string, heightPt: number, wrap: string, offsetPt = 0): string =>
    `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
      <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${String(id)}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(offsetPt)}</wp:posOffset></wp:positionV>
        <wp:extent cx="${emu(BOX_PT)}" cy="${emu(heightPt)}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        ${wrap}
        <wp:docPr id="${String(id)}" name="${name}"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:cNvSpPr txBox="1"/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(BOX_PT)}" cy="${emu(heightPt)}"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr>
            <wps:txbx><w:txbxContent>${paragraph(exactly(LINE_PT), run(`${name} boxed`))}</w:txbxContent></wps:txbx>
            <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"/>
          </wps:wsp></a:graphicData></a:graphic>
      </wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>`;

  const TOP_AND_BOTTOM = `<wp:wrapTopAndBottom/>`;
  const SQUARE = `<wp:wrapSquare wrapText="bothSides"/>`;

  type Case = {
    readonly name: string;
    // How tall the object is, which is where the foot of its band falls: the
    // anchoring paragraph starts 60pt down the page and the object with it.
    readonly heightPt: number;
    // What the paragraph under it asks for above itself.
    readonly abovePt: number;
    readonly wrap: string;
    // How far below its own paragraph the object hangs. Nought puts it over the
    // anchoring paragraph's own line; anything past that line's foot leaves it
    // where it is, which is how the documents in the wild write one.
    readonly offsetPt?: number;
  };

  const block = (of: Case, id: number): string =>
    paragraph(`${OWN_PAGE}${exactly(LINE_PT)}`, run(`${of.name} marks`)) +
    paragraph(
      exactly(LINE_PT),
      boxed(id, of.name, of.heightPt, of.wrap, of.offsetPt ?? 0) + run(`${of.name} anchors`),
    ) +
    paragraph(above(of.abovePt, LINE_PT), run(`${of.name} follows`)) +
    paragraph(exactly(LINE_PT), run(`${of.name} after`));

  const CASES: readonly Case[] = [
    // The case in the wild: a band ending far below where the paragraph would have
    // started, and 36pt of room asked for above it. Absorbed puts the line at 160,
    // kept puts it at 196.
    { name: "a", heightPt: 100, abovePt: 36, wrap: TOP_AND_BOTTOM },
    // The same asking for nothing, which both readings put at 160.
    { name: "b", heightPt: 100, abovePt: 0, wrap: TOP_AND_BOTTOM },
    // A band whose foot falls between where the paragraph would have started and
    // where its own room would have put it. Absorbed and kept differ here too, and
    // so does taking the lower of the two: 120 for that, 136 for kept.
    { name: "c", heightPt: 40, abovePt: 36, wrap: SQUARE },
    // The same band and the same room, wrapped the other way, in case what the
    // text does beside an object decides this as well.
    { name: "d", heightPt: 100, abovePt: 36, wrap: SQUARE },
    // A band ending above the paragraph altogether, which pushes nothing: the line
    // stands where its own room puts it whatever the answer is, and says so.
    { name: "e", heightPt: 12, abovePt: 36, wrap: TOP_AND_BOTTOM },
    // The geometry of the documents in the wild, and the case the whole document is
    // for: the object hangs below its anchoring paragraph's own line, so nothing
    // about that line is in question, and the only thing that can move the follower
    // is what becomes of the room it asks for. Absorbed puts its line at 190, kept
    // puts it at 226.
    { name: "f", heightPt: 100, abovePt: 36, wrap: TOP_AND_BOTTOM, offsetPt: 30 },
    // A square wrap the follower's own line clears and the room above it does not,
    // asked with room too wide for the band's foot and that room to be confused:
    // the line stands at 144 and the band ends at 100, so a line narrowed here is
    // one narrowed by the room it asked for.
    { name: "g", heightPt: 40, abovePt: 60, wrap: SQUARE },
  ];

  return [...CASES.map((each, at) => block(each, at + 1)), EMPTY].join("");
}

// Where a section's columns stand across the page, and when its text leaves one of
// them for the next. Nothing here has ever been put to Word.
//
// Six questions, each a section of its own opening a page: **where the columns of an
// equal-width section fall**, **the same of three of them**, **where a section
// stating its own widths puts them**, **what `w:br w:type="column"` does**, **what
// becomes of a multi-column section too short to fill its columns**, and **what
// happens when the last column of a page is full**.
//
// Every section keeps a body 144pt tall: the page is letter and the bottom margin is
// nine inches, so a column holds exactly six 24pt lines and a filled column is six
// paragraphs rather than thirty. Every line names itself, so Word's own drawing says
// which column each of them landed in.
//
// **The first and last line of every case is right-aligned**, which is the only
// thing a drawing can say about how wide a column is: a right-aligned line starts
// where its column ends. The first of them stands in the first column and the last
// in whichever column the case ran on to, so the two together give the width and the
// gap without any line having to wrap.
export function columnsDocument(): string {
  const LINE_PT = 24;
  const exactly = `<w:spacing w:line="${String(LINE_PT * 20)}" w:lineRule="exact"/>`;
  const RIGHT = `<w:jc w:val="right"/>`;
  const COLUMN_BREAK = `<w:r><w:br w:type="column"/></w:r>`;

  const sectionProperties = (columns: string, type: string): string =>
    `<w:sectPr>` +
    (type === "" ? "" : `<w:type w:val="${type}"/>`) +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="720" w:right="720" w:bottom="12240" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>` +
    columns +
    `</w:sectPr>`;

  // Half an inch of gap and the whole 540pt to share, so two columns are 252pt and
  // three are 168pt if Word takes the gaps off the text and divides what is left.
  const TWO = `<w:cols w:num="2" w:space="720"/>`;
  const THREE = `<w:cols w:num="3" w:space="360"/>`;
  // Two columns Word is told the width of rather than asked to divide: a quarter of
  // the text and the rest of it, with the same half inch between them.
  const STATED =
    `<w:cols w:num="2" w:equalWidth="0" w:space="720">` +
    `<w:col w:w="3600" w:space="720"/><w:col w:w="6480"/></w:cols>`;
  const ONE = `<w:cols w:space="720"/>`;

  const line = (name: string): string => paragraph(exactly, run(name));
  const opens = (name: string): string => paragraph(`${exactly}${RIGHT}`, run(name));
  const closes = (name: string, columns: string, type: string): string =>
    paragraph(`${exactly}${RIGHT}${sectionProperties(columns, type)}`, run(name));

  const filling = (name: string, count: number): readonly string[] =>
    Array.from({ length: count }, (_, at) => line(`${name}${String(at + 1)}`));

  return [
    // Nine lines in two columns: six fill the first and three stand in the second.
    opens("a top"),
    ...filling("a", 7),
    closes("a foot", TWO, ""),
    // Eighteen in three, which is the whole of the first two and most of the third.
    opens("b top"),
    ...filling("b", 16),
    closes("b foot", THREE, ""),
    // Nine again, in the two columns the section states the widths of.
    opens("c top"),
    ...filling("c", 7),
    closes("c foot", STATED, ""),
    // A column break three lines into a column with room for six, written inside a
    // paragraph rather than between two, so that what it does to the rest of that
    // paragraph is part of the answer.
    opens("d top"),
    line("d1"),
    paragraph(exactly, run("d2") + COLUMN_BREAK + run("d3")),
    line("d4"),
    closes("d foot", TWO, ""),
    // Four lines in two columns that hold twelve, closed by a continuous break into
    // a single column. Two lines in each column say Word evened them out; four in
    // the first say it did not, and where the single column's own text stands says
    // what the section left behind it either way.
    opens("e top"),
    line("e1"),
    line("e2"),
    closes("e foot", TWO, ""),
    line("e under"),
    closes("e last", ONE, "continuous"),
    // Fourteen lines in two columns that hold twelve, so the last two have nowhere
    // on the page to go.
    opens("f top"),
    ...filling("f", 12),
    closes("f foot", TWO, ""),
    // And the body's own section, which is one column and the page every authored
    // document keeps, so a line drawn there is one no case reached.
    line("z a"),
    line("z b"),
  ].join("");
}

// Whether `w:keepNext` moves a paragraph onto the page its next one landed on,
// and how far back a run of them pulls.
//
// Every paragraph here is told exactly how tall to be, so the room left at the foot
// of a page is arithmetic rather than a measurement. Each block is a marker naming
// the case, eight fillers, a shim sized to leave the room the case wants, and the
// case itself; the marker opening the next block is where the flow resumed. Unlike
// the blocks in `breakingDocument`, each marker asks for a page of its own, since
// half the cases here are asking whether a paragraph moved on, and a block whose
// case moved would start partway down a page and leave the next one less room than
// it was told to leave.
export function keepingDocument(): string {
  const exactly = (pt: number): string =>
    `<w:spacing w:line="${String(pt * 20)}" w:lineRule="exact"/>`;

  const FILLERS = 8;
  const FILLER_PT = 72;
  const MARKER_PT = 24;
  const BLOCK_PT = 720;
  const LINE_PT = 24;

  // Room for two of the case's lines and not for three.
  const TWO_PT = 54;
  // Room for one of them and not for two.
  const ONE_PT = 30;

  const KEEP = `<w:keepNext/>`;
  const OWN_PAGE = `<w:pageBreakBefore/>`;
  const BREAK = `<w:r><w:br/></w:r>`;
  const PAGE_BREAK = `<w:r><w:br w:type="page"/></w:r>`;

  const line = (name: string, properties = ""): string =>
    paragraph(`${properties}${exactly(LINE_PT)}`, run(name));

  const held = (name: string): string => line(name, KEEP);

  const block = (name: string, leftPt: number, ...cases: readonly string[]): readonly string[] => [
    paragraph(`${OWN_PAGE}${exactly(MARKER_PT)}`, run(name)),
    ...Array.from({ length: FILLERS }, () => paragraph(exactly(FILLER_PT), run("filler"))),
    paragraph(exactly(BLOCK_PT - MARKER_PT - FILLERS * FILLER_PT - leftPt), run("shim")),
    ...cases,
  ];

  return [
    // Room for the held paragraph and for the one it holds. Nothing should move,
    // and this is what the other cases read against.
    ...block("fits", TWO_PT, held("held"), line("next")),
    // Room for the held paragraph alone, which is the ordinary case.
    ...block("moved", ONE_PT, held("held"), line("next")),
    // Three in a chain with room for the first two, which says how far back the
    // last one pulls: the paragraph before it, or all of them.
    ...block("chain", TWO_PT, held("chain one"), held("chain two"), line("chain three")),
    // A pair that can never stand together, the next paragraph being a single line
    // taller than the whole body. Word has to give up somewhere, and whether it
    // gives up where the pair started or after moving it once is the whole of what
    // says the rule is a loop with a stop in it rather than one look ahead.
    ...block(
      "never",
      ONE_PT,
      held("held"),
      paragraph(exactly(BLOCK_PT + ONE_PT), run("taller than a page")),
    ),
    // The same pair with a chain in front of it, which is the one case where what
    // stops the rule shows: the paragraph at the head of a chain has nowhere left
    // to move to and the ones behind it still have.
    ...block(
      "never in a chain",
      TWO_PT,
      held("chained one"),
      held("chained two"),
      paragraph(exactly(BLOCK_PT + ONE_PT), run("chained taller than a page")),
    ),
    // A paragraph asking for a page of its own, held by the one above it. There is
    // room here for both, so anything that moves is the two rules meeting rather
    // than the foot of the page.
    ...block("against a break", TWO_PT, held("held"), line("own page", OWN_PAGE)),
    // The same meeting from the other side, the break being one the held paragraph
    // asked for itself rather than one its next asked for.
    ...block(
      "against its own break",
      TWO_PT,
      paragraph(`${KEEP}${exactly(LINE_PT)}`, run("holds and breaks") + PAGE_BREAK),
      line("after the break"),
    ),
    // A held paragraph the break runs through, whose last line therefore already
    // stands on the page its next one landed on. Four lines with room for two,
    // which is one more than widow control moves back, so whatever moves here is
    // this rule. Says whether the rule is about the paragraph or about its end.
    ...block(
      "split",
      TWO_PT,
      paragraph(
        `${KEEP}${exactly(LINE_PT)}`,
        run("split one") +
          BREAK +
          run("split two") +
          BREAK +
          run("split three") +
          BREAK +
          run("split four"),
      ),
      line("next"),
    ),
    // The last paragraph of the document, holding onto nothing. Nothing follows it,
    // which is why it is last.
    ...block("holds nothing", ONE_PT, held("last")),
  ].join("");
}

// Whether the line drawn between two rows takes room of its own, on top of the
// margins the cells either side of it hold their text off the wall by.
//
// What is built says it does not: the text clears the half of the line that falls
// inside the cell, and a margin already holding it further off than that leaves the
// border asking for nothing. That was read off tables of one row, where the
// question never comes up. A real document of twelve pages drifts about a point a
// row down every table in it, and its cells are held off their walls by 5pt and
// lined with one, which is exactly where the two readings differ.
//
// Every line here is told to be 20pt exactly and every row holds one, so the
// distance from one row to the next is the whole of the answer, and each case is
// four rows so that three of those distances are read rather than one. Each case
// opens a page of its own, since a break between two rows makes a nonsense of the
// distance across it.
export function linedRowsDocument(): string {
  const CELL_TWIPS = 2880;
  const LINE_PT = 20;

  const exactly = `<w:spacing w:line="${String(LINE_PT * 20)}" w:lineRule="exact"/>`;

  const around = (eighths: number): string =>
    eighths === 0
      ? ""
      : `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"]
          .map(
            (side) =>
              `<w:${side} w:val="single" w:sz="${String(eighths)}" w:space="0" w:color="FF0000"/>`,
          )
          .join("")}</w:tblBorders>`;

  const margins = (twips: number): string =>
    `<w:tblCellMar>
      <w:top w:w="${String(twips)}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="${String(twips)}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar>`;

  const cell = (content: string, properties = ""): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${String(CELL_TWIPS)}" w:type="dxa"/>${properties}</w:tcPr>${content}</w:tc>`;

  const line = (name: string): string => paragraph(exactly, run(name));

  const table = (name: string, eighths: number, marginTwips: number, own = ""): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="${String(CELL_TWIPS * 2)}" w:type="dxa"/>
      <w:tblInd w:w="0" w:type="dxa"/>${margins(marginTwips)}${around(eighths)}</w:tblPr>
      <w:tblGrid><w:gridCol w:w="${String(CELL_TWIPS)}"/><w:gridCol w:w="${String(CELL_TWIPS)}"/></w:tblGrid>
      ${Array.from(
        { length: 4 },
        (_, at) =>
          `<w:tr>${cell(line(`${name}${String(at + 1)} left`), own)}${cell(line(`${name}${String(at + 1)} right`))}</w:tr>`,
      ).join("")}</w:tbl>`;

  // A line of its own either side of each table, so that where the table starts and
  // where it ends are read as well as the distance from row to row.
  const block = (name: string, eighths: number, marginTwips: number, own = ""): string =>
    paragraph(`<w:pageBreakBefore/>${exactly}`, run(`case ${name}`)) +
    table(name, eighths, marginTwips, own) +
    paragraph(exactly, run(`${name} after`));

  const NIL_TOP =
    `<w:tcBorders><w:top w:val="nil"/>` +
    `<w:bottom w:val="single" w:sz="2" w:space="0" w:color="FF0000"/></w:tcBorders>`;

  const SIX_POINTS = `<w:tcBorders>${["top", "bottom"]
    .map((side) => `<w:${side} w:val="single" w:sz="48" w:space="0" w:color="0070C0"/>`)
    .join("")}</w:tcBorders>`;

  return [
    // No line at all, at each margin, which is what the rest read against.
    block("a", 0, 0),
    block("b", 0, 100),
    // A line and no margin, where every reading agrees: what falls inside the cell
    // is all there is to clear, and it is the same room either way.
    block("c", 8, 0),
    block("d", 48, 0),
    // A line and a margin wider than half of it, which is where the readings part.
    // 5pt of margin against a 1pt line is 30pt a row if the line asks for nothing,
    // 31 if the whole of it stands between the rows.
    block("e", 8, 100),
    block("f", 24, 100),
    // And the same against a 6pt line: 30pt a row, 32 if each row clears the whole
    // line rather than half of it, 36 if the line stands between them.
    block("g", 48, 100),
    // A line only the first cell of each row asks for, which says whether the room
    // one cell's line takes is the row's or its own.
    block("h", 0, 100, SIX_POINTS),
    // Widths that are not whole points, which is what tells a floor under the room
    // a line takes from a rounding up of it. A quarter point takes a whole one, and
    // these three say whether one and a half does too.
    block("i", 4, 100),
    block("j", 12, 100),
    block("k", 18, 100),
    // The thinnest line a document can ask for, and the one between it and half a
    // point. A real one-pager is lined with a quarter of a point throughout, and it
    // is the case every other reading here has to be checked against.
    block("l", 2, 100),
    block("m", 6, 100),
    // A quarter point line against the 2.75pt margin a real one-pager holds its
    // cells off their walls by, which is the one case in the suites that reads the
    // other way. Its rows come out a quarter point taller here than Word draws
    // them, and this says whether the line is what that quarter point is.
    block("n", 2, 55),
    // The same again with the line stated the way that one-pager states it: every
    // cell refusing a line at its top and asking for one at its foot, so the line
    // between two rows is asked for from one side only. Read against n, this says
    // whether a line only one of the two rows owns is still room for both.
    block("o", 0, 55, NIL_TOP),
    // A table is the last thing in the body, which Word will not have.
    EMPTY,
  ].join("");
}

// How tall a row asking for a height taller than its text comes out, and whether
// the line between two such rows takes room on top of what they asked for.
//
// `lined-rows` settled what a border costs a row measured by its own text: the half
// falling inside the cell is cleared, and the margin is cleared after it. It says
// nothing about a row that states a height, because the height it settled is the
// one the text asks for.
//
// A three page document in the wild is 1.2pt out by its last row and the whole of
// that is at the two row boundaries above it. Every row of it states a
// `w:trHeight` taller than its text, and every row of it comes out here at exactly
// the number stated. Word's rows are a fraction taller, which is the size of the
// line between them.
//
// Each case is four rows of two cells with one 20pt line each, so a row's height is
// the distance from one repeat to the next rather than a difference of two rounded
// answers, and a line of its own stands either side of the table so where it starts
// and where it ends are read too. Every stated height is well above what the text
// needs, except in the one case written to say that a stated height is a floor.
export function statedRowHeightsDocument(): string {
  const CELL_TWIPS = 2880;
  const LINE_PT = 20;

  const exactly = `<w:spacing w:line="${String(LINE_PT * 20)}" w:lineRule="exact"/>`;

  const around = (eighths: number): string =>
    eighths === 0
      ? ""
      : `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"]
          .map(
            (side) =>
              `<w:${side} w:val="single" w:sz="${String(eighths)}" w:space="0" w:color="FF0000"/>`,
          )
          .join("")}</w:tblBorders>`;

  // An authored document declares no table style, so a table stating no margins is
  // held off its walls by nothing at all.
  const margins = (twips: number): string =>
    `<w:tblCellMar>
      <w:top w:w="${String(twips)}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="${String(twips)}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar>`;

  const cell = (content: string): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${String(CELL_TWIPS)}" w:type="dxa"/></w:tcPr>${content}</w:tc>`;

  const line = (name: string): string => paragraph(exactly, run(name));

  type Case = {
    readonly name: string;
    readonly heightPt: number;
    readonly rule: "atLeast" | "exact";
    readonly eighths: number;
    readonly marginTwips: number;
  };

  const table = (of: Case): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="${String(CELL_TWIPS * 2)}" w:type="dxa"/>
      <w:tblInd w:w="0" w:type="dxa"/>${margins(of.marginTwips)}${around(of.eighths)}</w:tblPr>
      <w:tblGrid><w:gridCol w:w="${String(CELL_TWIPS)}"/><w:gridCol w:w="${String(CELL_TWIPS)}"/></w:tblGrid>
      ${Array.from(
        { length: 4 },
        (_, at) =>
          `<w:tr><w:trPr><w:trHeight w:val="${String(of.heightPt * 20)}"${of.rule === "exact" ? ` w:hRule="exact"` : ""}/></w:trPr>` +
          `${cell(line(`${of.name}${String(at + 1)} left`))}${cell(line(`${of.name}${String(at + 1)} right`))}</w:tr>`,
      ).join("")}</w:tbl>`;

  const block = (of: Case): string =>
    paragraph(`<w:pageBreakBefore/>${exactly}`, run(`case ${of.name}`)) +
    table(of) +
    paragraph(exactly, run(`${of.name} after`));

  const ASKED_PT = 60;

  return [
    // A stated height and no line at all, which the rest read against.
    block({ name: "a", heightPt: ASKED_PT, rule: "atLeast", eighths: 0, marginTwips: 0 }),
    // The same with a margin, which says whether a margin is inside the height a row
    // asked for or on top of it.
    block({ name: "b", heightPt: ASKED_PT, rule: "atLeast", eighths: 0, marginTwips: 100 }),
    // A line at a point and at six, against no margin. Rows 60pt apart say the
    // stated height is the whole of the row; 61 and 66 say the line stands on top of
    // it, and 60.5 and 63 say half of it does.
    block({ name: "c", heightPt: ASKED_PT, rule: "atLeast", eighths: 8, marginTwips: 0 }),
    block({ name: "d", heightPt: ASKED_PT, rule: "atLeast", eighths: 48, marginTwips: 0 }),
    // The same two with a margin the line has to be cleared before, which is what
    // `lined-rows` found a row measured by its text does.
    block({ name: "e", heightPt: ASKED_PT, rule: "atLeast", eighths: 8, marginTwips: 100 }),
    block({ name: "f", heightPt: ASKED_PT, rule: "atLeast", eighths: 48, marginTwips: 100 }),
    // A height under what the text needs, so the text wins and the case says that a
    // stated height is a floor rather than the answer.
    block({ name: "g", heightPt: 10, rule: "atLeast", eighths: 48, marginTwips: 100 }),
    // And a height stated exactly, which is a row that cannot grow for anything.
    // Word draws this one a margin taller than the number stated, so the three under
    // it take the margin and the line away one at a time to say which of them it is.
    block({ name: "h", heightPt: ASKED_PT, rule: "exact", eighths: 48, marginTwips: 100 }),
    // Exact with a margin and no line at all.
    block({ name: "i", heightPt: ASKED_PT, rule: "exact", eighths: 0, marginTwips: 100 }),
    // Exact with a line and no margin.
    block({ name: "j", heightPt: ASKED_PT, rule: "exact", eighths: 48, marginTwips: 0 }),
    // And exact with neither, which is the number stated and nothing else.
    block({ name: "k", heightPt: ASKED_PT, rule: "exact", eighths: 0, marginTwips: 0 }),
    // A table is the last thing in the body, which Word will not have.
    EMPTY,
  ].join("");
}

// Whether a row the foot of a page falls in is torn across the break or moved
// whole, and what a stated height, `w:cantSplit` and a row above it change.
//
// Every paragraph here is told exactly how tall to be, so the room left under the
// last filler is arithmetic: each case opens a page of its own, seven fillers and a
// shim fill it down to 102pt from the foot, and the case's table starts there. That
// leaves room for four of its 24pt lines with six points to spare, so nothing here
// is decided at a boundary.
//
// Every line names itself rather than repeating a word, so that Word's own drawing
// of one can be paired with ours by its text.
export function tearingDocument(): string {
  const exactly = (pt: number): string =>
    `<w:spacing w:line="${String(pt * 20)}" w:lineRule="exact"/>`;

  const FILLERS = 7;
  const FILLER_PT = 72;
  const MARKER_PT = 24;
  const BLOCK_PT = 720;
  const LINE_PT = 24;
  // Room for four of the case's lines and not for five.
  const ROOM_PT = 102;

  // The whole width of the text column, so a line in a cell breaks where a line
  // outside one would.
  const WIDTH = 10800;

  const cell = (content: string): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${String(WIDTH)}" w:type="dxa"/></w:tcPr>${content}</w:tc>`;

  const row = (properties: string, content: string): string =>
    `<w:tr>${properties === "" ? "" : `<w:trPr>${properties}</w:trPr>`}${cell(content)}</w:tr>`;

  // An authored document declares no table style, so a table that states no cell
  // margins is held off its walls by nothing at all. Above and below they are
  // nought here, which is what makes a row exactly as tall as its lines.
  const table = (rows: string): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="${String(WIDTH)}" w:type="dxa"/><w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="${String(WIDTH)}"/></w:tblGrid>${rows}</w:tbl>`;

  const named = (name: string, at: number): string => `${name} ${String(at).padStart(2, "0")}`;

  const lines = (name: string, from: number, to: number, linePt = LINE_PT): string =>
    Array.from({ length: to - from + 1 }, (_, at) =>
      paragraph(exactly(linePt), run(named(name, from + at))),
    ).join("");

  const block = (name: string, content: string): readonly string[] => [
    paragraph(`<w:pageBreakBefore/>${exactly(MARKER_PT)}`, run(`case ${name}`)),
    ...Array.from({ length: FILLERS }, (_, at) =>
      paragraph(exactly(FILLER_PT), run(`${name} fill ${String(at + 1)}`)),
    ),
    paragraph(exactly(BLOCK_PT - MARKER_PT - FILLERS * FILLER_PT - ROOM_PT), run(`${name} shim`)),
    content,
  ];

  // A line taller than the ones above, chosen so that neither the room left at the
  // foot of the page nor the whole body is a whole number of them.
  const TALL_PT = 32;

  return [
    // A row of six lines with room for four, saying whether the four stay.
    ...block("a", table(row("", lines("a", 1, 6)))),
    // The same row told it may not be torn, which is the answer to read the first
    // against: a case that moves here and stays there says the oracle can tell the
    // two apart at all.
    ...block("b", table(row(`<w:cantSplit/>`, lines("b", 1, 6)))),
    // A row holding two lines and asking to be 150pt tall, so what does not fit is
    // the height it stated rather than anything drawn. Says which of the two the
    // break is measured against.
    ...block("c", table(row(`<w:trHeight w:val="3000"/>`, lines("c", 1, 2)))),
    // A row that does not fit under one that does, which says whether what moves is
    // the row or the table it is in.
    ...block("d", table(row("", lines("d", 1, 2)) + row("", lines("d", 3, 8)))),
    // A row of 768pt, which no page has room for whole. Word has to tear this one
    // somewhere, and where says whether a row that will not fit is moved first: 22
    // of its lines fill a page of their own, and three of them fit in the room left
    // here.
    ...block("e", table(row("", lines("e", 1, 24, TALL_PT)))),
    // A row asking to be 150pt tall and holding 144pt, so that the height it stated
    // and the text in it disagree about nothing but which is the taller. Read
    // against c, this says whether a stated height rules the break even where the
    // text alone would have been torn.
    ...block("f", table(row(`<w:trHeight w:val="3000"/>`, lines("f", 1, 6)))),
    // The same six lines under a stated height of 48pt, which the room left has
    // room for and the text has not. Says whether it is the larger of the two the
    // break is measured against or the stated one whatever it says.
    ...block("g", table(row(`<w:trHeight w:val="960"/>`, lines("g", 1, 6)))),
    // A row that may not be torn and that no page has room for whole, which is the
    // one case where the two answers cannot both be had.
    ...block("h", table(row(`<w:cantSplit/>`, lines("h", 1, 24, TALL_PT)))),
    // A table is the last thing in the body, which Word will not have.
    EMPTY,
  ].join("");
}

// Whether a space raises the line it stands on, and how tall a paragraph holding
// nothing but spaces comes out.
//
// Nothing here is told how tall to be: the height being asked about is the
// distance from one repeat of a case to the next, which is why each is written
// three times. Every run in a case is 24pt against the 12pt every paragraph mark
// keeps, so a line that took the run's height stands twice as tall as one that did
// not and no rounding is in the way. Each case opens a page of its own so that a
// break can never fall between two repeats and turn a height into nonsense.
export function trailingSpaceDocument(): string {
  const BIG = `<w:sz w:val="48"/><w:szCs w:val="48"/>`;
  const TAB = `<w:r><w:rPr>${BIG}</w:rPr><w:tab/></w:r>`;

  const big = (value: string): string => run(value, BIG);
  const plain = (value: string): string => run(value);

  const block = (name: string, repeat: (at: number) => string): readonly string[] => [
    paragraph(`<w:pageBreakBefore/>`, run(`case ${name}`)),
    ...Array.from({ length: 3 }, (_, at) => repeat(at + 1)),
  ];

  const named = (name: string, at: number): string => `${name}0${String(at)}`;

  return [
    // A space and nothing else, which is the case a real document met: either the
    // run it is written in raises the line or the paragraph mark is the whole of
    // what is left to measure.
    ...block("a", () => paragraph("", big(" "))),
    // The same space at the end of a line that has text in front of it, which says
    // whether what matters is the space being all there is or the space being last.
    ...block("b", (at) => paragraph("", plain(named("b", at)) + big(" "))),
    // A space between two words, which no line ends at.
    ...block("c", (at) => paragraph("", plain(named("c", at)) + big(" ") + plain("zz"))),
    // A space in front of the text rather than behind it.
    ...block("d", (at) => paragraph("", big(" ") + plain(named("d", at)))),
    // Two of them together, since a run of spaces may not answer as one does.
    ...block("e", () => paragraph("", big("  "))),
    // A tab alone, which is already measured to hold the line open at the tallest
    // mark the paragraph has. It is here so that a space answering otherwise is
    // read against something rather than guessed at.
    ...block("f", () => paragraph("", TAB)),
    // Text in the same run, which has to raise the line however the spaces answer.
    ...block("g", (at) => paragraph("", big(named("g", at)))),
    // Nothing at all, under a mark of its own that is the big size, which is the
    // one case where the mark is the only thing there is to measure.
    ...block("h", () => paragraph(`<w:rPr>${BIG}</w:rPr>`, "")),
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

// Where a table's own indent puts it, and what the cell's margin then does to
// that. The same body is written twice, once declaring 15 and once declaring
// nothing, so the difference between the two is the setting and nothing else.
//
// The two answers this can come out as are a whole cell margin apart: an indent
// measured to the table's edge holds the text off the margin by the indent and
// the cell margin both, and one measured to the cell's text holds it off by the
// indent alone, which leaves the table's own edge hanging outside the margin
// where a case states no indent at all.
//
// So every case says both. Its text says where the text went, and its border says
// where the edge went, and each case states an indent against a cell margin the
// indent is under, over or exactly equal to. The last two state no margin at all:
// with nothing to swallow the indent, the two answers are the same, which is what
// says the margin is what the difference is made of.
export function tableIndentDocument(): string {
  const CELL_TWIPS = 2880;

  const cases = [
    { indent: 0, margin: 108 },
    { indent: 54, margin: 108 },
    { indent: 108, margin: 108 },
    { indent: 720, margin: 108 },
    { indent: 0, margin: 0 },
    { indent: 720, margin: 0 },
  ];

  // A cell margin of its own on the cell rather than on the table, since a real
  // document states both and only this says which of the two an indent is
  // measured against.
  const OWN_MARGIN_TWIPS = 288;

  const named = (indent: number, margin: number): string => `i${String(indent)} m${String(margin)}`;

  const own = (twips: number): string =>
    `<w:tcMar><w:top w:w="0" w:type="dxa"/><w:left w:w="${String(twips)}" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="${String(twips)}" w:type="dxa"/></w:tcMar>`;

  const indented = (indent: number, margin: number, own: string, name: string): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="${String(CELL_TWIPS)}" w:type="dxa"/>
      <w:tblInd w:w="${String(indent)}" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/><w:left w:w="${String(margin)}" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/><w:right w:w="${String(margin)}" w:type="dxa"/>
      </w:tblCellMar>
      <w:tblBorders><w:left w:val="single" w:sz="8" w:space="0" w:color="C00000"/></w:tblBorders></w:tblPr>
      <w:tblGrid><w:gridCol w:w="${String(CELL_TWIPS)}"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:w="${String(CELL_TWIPS)}" w:type="dxa"/>${own}
        <w:tcBorders><w:left w:val="single" w:sz="8" w:space="0" w:color="C00000"/></w:tcBorders></w:tcPr>
        ${paragraph("", run(name))}</w:tc></w:tr></w:tbl>`;

  return [
    paragraph("", run("above")),
    ...cases.flatMap(({ indent, margin }) => [
      indented(indent, margin, "", named(indent, margin)),
      paragraph("", run("between")),
    ]),
    // The same indent again with the cell stating a margin of its own, which is
    // wider than the table's and is what the text should stand off if a cell's
    // own margin is what the indent is measured against.
    indented(108, 108, own(OWN_MARGIN_TWIPS), `i108 own${String(OWN_MARGIN_TWIPS)}`),
    paragraph("", run("between")),
    // Two rows whose cells state different margins of their own. A table has one
    // edge and the rows disagree about how far their text stands off it, so which
    // of them the indent is measured against is what this says.
    `<w:tbl><w:tblPr><w:tblW w:w="${String(CELL_TWIPS)}" w:type="dxa"/>
      <w:tblInd w:w="108" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar>
      <w:tblBorders><w:left w:val="single" w:sz="8" w:space="0" w:color="C00000"/></w:tblBorders></w:tblPr>
      <w:tblGrid><w:gridCol w:w="${String(CELL_TWIPS)}"/></w:tblGrid>
      ${[0, OWN_MARGIN_TWIPS]
        .map(
          (twips) =>
            `<w:tr><w:tc><w:tcPr><w:tcW w:w="${String(CELL_TWIPS)}" w:type="dxa"/>${own(twips)}
              <w:tcBorders><w:left w:val="single" w:sz="8" w:space="0" w:color="C00000"/></w:tcBorders></w:tcPr>
              ${paragraph("", run(`i108 row${String(twips)}`))}</w:tc></w:tr>`,
        )
        .join("")}</w:tbl>`,
    paragraph("", run("below")),
    EMPTY,
  ].join("");
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
// whichever face on the machine has it. `WORD_CHARACTER_FALLBACK_FACES` in core
// names them in the order Word reaches for them.
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

// Whether the whitespace at the edge of a `w:t` survives when the run does not ask
// for it with `xml:space="preserve"`.
//
// This project drops it, on the reading that whitespace at the edge of an element is
// what xml leaves insignificant and `xml:space` is what asks for it back. **The worst
// placed document in the corpus says otherwise**: it writes a heading as a run
// holding `Protocolo de demostración`, a run holding one bare space, and a run
// holding the rest, and Word breaks the line at that space while this project has
// nothing to break at and carries an unbreakable 287pt word off the end of the line.
// Every line of the document below it is then in the wrong place.
//
// **953 of the corpus `w:t` elements over 113 documents hold whitespace at an edge
// without asking for it**, and 3469 more over four documents hold nothing else at
// all, so what Word does here is worth asking properly.
//
// Every case is right-aligned, so where the line starts says how wide it is and a
// space either side of the question is worth 12pt of it: the runs are at 48pt and a
// space there is nearly four points, which no rounding reaches. Each case is written
// out three times, and each stands on a page of its own. **Read off Word's own pdf**,
// which is where the line was drawn; Word's report answers nought for the left of a
// right-aligned line.
export function insignificantSpaceDocument(): string {
  const BIG = `<w:sz w:val="96"/><w:szCs w:val="96"/>`;
  const OWN_PAGE = `<w:pageBreakBefore/>`;

  // A run whose text is written exactly as given, with no request to keep the
  // whitespace at its edges. `run` above always asks, which is the whole of what
  // this document is about.
  const bare = (value: string): string => `<w:r><w:rPr>${BIG}</w:rPr><w:t>${value}</w:t></w:r>`;
  const kept = (value: string): string =>
    `<w:r><w:rPr>${BIG}</w:rPr><w:t xml:space="preserve">${value}</w:t></w:r>`;

  const RIGHT = `<w:jc w:val="right"/>`;

  // Each repeat carries a mark of its own, since two lines drawn with the same text
  // cannot be told apart in a pdf and the comparison would be guessing which was
  // which. The mark is written at the end of the last word, so it moves every case's
  // line by the same amount and the widths still differ by the space alone.
  const marked = (at: number): string => "x".repeat(at + 1);

  const block = (name: string, content: string): readonly string[] => [
    paragraph(OWN_PAGE, run(`${name} above`)),
    ...Array.from({ length: 3 }, (_, at) => paragraph(RIGHT, content + bare(marked(at)))),
  ];

  // The same runs at a size two words of which overflow the room the paragraph is
  // narrowed to, so where the line broke is the answer rather than how wide it was.
  const SMALL = `<w:sz w:val="48"/><w:szCs w:val="48"/>`;
  const NARROW = `<w:ind w:right="7200"/>`;

  const bare24 = (value: string): string => `<w:r><w:rPr>${SMALL}</w:rPr><w:t>${value}</w:t></w:r>`;
  const kept24 = (value: string): string =>
    `<w:r><w:rPr>${SMALL}</w:rPr><w:t xml:space="preserve">${value}</w:t></w:r>`;
  const bold24 = (value: string): string =>
    `<w:r><w:rPr><w:b/>${SMALL}</w:rPr><w:t>${value}</w:t></w:r>`;
  const EMPTY_RUN = `<w:r></w:r>`;

  const narrow = (name: string, content: string): readonly string[] => [
    paragraph(OWN_PAGE, run(`${name} above`)),
    ...Array.from({ length: 3 }, (_, at) => paragraph(NARROW, content + bare24(marked(at)))),
  ];

  // Room enough for the word two halves make but not for it beside what stands in
  // front of it, which is the case in the wild and the one the narrow cases above
  // cannot ask: their two halves make a word too long for any line at all, so Word
  // cut it wherever it overflowed and where the runs divided never came into it.
  const ROOMIER = `<w:ind w:right="4000"/>`;

  const roomier = (name: string, content: string): readonly string[] => [
    paragraph(OWN_PAGE, run(`${name} above`)),
    ...Array.from({ length: 3 }, (_, at) => paragraph(ROOMIER, content + bare24(marked(at)))),
  ];

  // The nine letters and the two in front of the long word are what leaves the room
  // too small for it, and the word itself is 21 letters over the two runs, which one
  // line of its own holds.
  const AHEAD = "aaaaaaaaa bb ";
  const HALF = "cccccccccccc";
  const REST = "eeeeeeeee fff";

  return [
    // The two ends of the answer, so every case below is read against a width rather
    // than against a font's own numbers: no space at all, and one that asked to stay.
    ...block("none", bare("aaaa") + bare("bbbb")),
    ...block("asked", kept("aaaa ") + bare("bbbb")),
    // A space at the end of a run that did not ask to keep it, which is the commonest
    // of the three in the wild.
    ...block("trailing", bare("aaaa ") + bare("bbbb")),
    // And at the start of the run after it.
    ...block("leading", bare("aaaa") + bare(" bbbb")),
    // A run holding nothing but the space, which is the case the worst-placed
    // document in the corpus turns on.
    ...block("alone", bare("aaaa") + bare(" ") + bare("bbbb")),
    // The same run asking to keep it, which is what the case above is read against.
    ...block("alone asked", bare("aaaa") + kept(" ") + bare("bbbb")),

    // And where a line has to break, since dropping the space is only half the
    // question: the document in the wild breaks exactly where its dropped space was,
    // and either the space survived after all or a run boundary is a place a line may
    // break. These four are narrowed to 180pt, which one twelve-letter word at 24pt
    // fills and two overflow.
    ...narrow("break in one run", bare24("aaaaaaaaaaaabbbbbbbbbbbb")),
    ...narrow("break at a boundary", bare24("aaaaaaaaaaaa") + bare24("bbbbbbbbbbbb")),
    ...narrow("break at a bare space", bare24("aaaaaaaaaaaa ") + bare24("bbbbbbbbbbbb")),
    ...narrow("break at an asked space", kept24("aaaaaaaaaaaa ") + bare24("bbbbbbbbbbbb")),
    // The same two, with the runs formatted differently either side of the boundary,
    // which is what the document in the wild holds: its space is Times New Roman
    // between two runs of Arial bold. Either a boundary between runs that are not
    // written alike is a place a line may break, or that document's space survived
    // where these did not.
    ...narrow("break at a bold boundary", bold24("aaaaaaaaaaaa") + bare24("bbbbbbbbbbbb")),
    ...narrow(
      "break at a bold bare space",
      bold24("aaaaaaaaaaaa") + bare24(" ") + bold24("bbbbbbbbbbbb"),
    ),
    // The document in the wild writes an empty run between the space and the text
    // after it, which is the last thing about its shape these cases do not have.
    ...narrow("break at an empty run", bold24("aaaaaaaaaaaa") + EMPTY_RUN + bold24("bbbbbbbbbbbb")),
    ...narrow(
      "break at a space and an empty run",
      bold24("aaaaaaaaaaaa") + bare24(" ") + EMPTY_RUN + bold24("bbbbbbbbbbbb"),
    ),

    // The same boundary where the word the two halves make fits on a line of its own.
    // Three answers are open and they are three different lines: the space survived
    // and the line ends where it was, it was dropped and the whole word moved down,
    // or it was dropped and the word was cut wherever it overflowed. **The
    // worst-placed document in the corpus turns on this**: Word ends its heading's
    // first line exactly at such a boundary, which the cases above say it has no
    // reason to.
    ...roomier(
      "room for the word alone, at a bare space",
      bold24(AHEAD + HALF) + bare24(" ") + EMPTY_RUN + bold24(REST),
    ),
    // The same with nothing between the two halves at all, which says what becomes of
    // a word that fits a line of its own and not the room left.
    ...roomier("room for the word alone, at a boundary", bold24(AHEAD + HALF) + bold24(REST)),
    // And with the space asking to stay, which is the line the first of the three
    // answers would draw.
    ...roomier(
      "room for the word alone, at an asked space",
      bold24(AHEAD + HALF) + kept24(" ") + EMPTY_RUN + bold24(REST),
    ),
    // The bare space written like the runs it stands between, since the one in the
    // wild is not: it names a face of its own and neither of the others.
    ...roomier(
      "room for the word alone, at a bold bare space",
      bold24(AHEAD + HALF) + bold24(" ") + EMPTY_RUN + bold24(REST),
    ),
    EMPTY,
  ].join("");
}

// A cell merged down a run of rows: where its text sits, what its rows are worth,
// and what becomes of the cells the merge swallowed.
//
// Ten of the 966 state `w:vMerge` and one of them holds 635 lines this project
// places 142 of, so the question is not academic. Nothing here is built: a
// continuation cell is laid out today as a cell of its own, and the restart cell's
// whole content is charged to the one row it opens.
//
// Every case is four rows of two cells with one 20pt line in the right hand one, so
// a row's height is the distance from one repeat to the next. The left hand column
// is what the case varies. Word's own pdf is the oracle: its report answers for the
// row, and every question here is about a cell.
export function mergedCellsDocument(): string {
  const CELL_TWIPS = 2880;
  const LINE_PT = 20;
  const ROWS = 4;

  const exactly = (pt: number): string =>
    `<w:spacing w:line="${String(pt * 20)}" w:lineRule="exact"/>`;

  const line = (name: string, pt = LINE_PT): string => paragraph(exactly(pt), run(name));

  // An authored document declares no table style, so a table stating no margins is
  // held off its walls by nothing at all, and a row is exactly as tall as its lines.
  const NO_MARGINS = `<w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>
    </w:tblCellMar>`;

  // Fixed, so that a cell stating no width of its own is the grid it stands on
  // rather than whatever Word would rather it were.
  const table = (columnTwips: readonly number[], rows: string): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="${String(columnTwips.reduce((a, b) => a + b, 0))}" w:type="dxa"/>
      <w:tblInd w:w="0" w:type="dxa"/><w:tblLayout w:type="fixed"/>${NO_MARGINS}</w:tblPr>
      <w:tblGrid>${columnTwips.map((twips) => `<w:gridCol w:w="${String(twips)}"/>`).join("")}</w:tblGrid>
      ${rows}</w:tbl>`;

  const cell = (properties: string, content: string): string =>
    `<w:tc><w:tcPr>${properties}</w:tcPr>${content}</w:tc>`;

  const width = (twips: number): string => `<w:tcW w:w="${String(twips)}" w:type="dxa"/>`;

  const RESTART = `<w:vMerge w:val="restart"/>`;
  const CONTINUE = `<w:vMerge/>`;

  // The right hand column of every case: one 20pt line a row, naming its row, which
  // is what every height here is read off.
  const marker = (name: string, at: number): string =>
    cell(width(CELL_TWIPS), line(`${name}${String(at)} right`));

  const merged = (name: string, lines: number, properties = ""): string =>
    cell(
      `${width(CELL_TWIPS)}${properties}${RESTART}`,
      Array.from({ length: lines }, (_, at) => line(`${name} m${String(at + 1)}`)).join(""),
    );

  const swallowed = (content = EMPTY): string => cell(`${width(CELL_TWIPS)}${CONTINUE}`, content);

  // A merge opening at the first row and running to the row named, the rows after it
  // holding an ordinary cell of their own.
  const mergedTable = (
    name: string,
    lines: number,
    through: number,
    options: { readonly restart?: string; readonly swallowedContent?: string } = {},
  ): string =>
    table(
      [CELL_TWIPS, CELL_TWIPS],
      Array.from({ length: ROWS }, (_, at) => {
        const left =
          at === 0
            ? merged(name, lines, options.restart ?? "")
            : at < through
              ? swallowed(options.swallowedContent)
              : cell(width(CELL_TWIPS), line(`${name}${String(at + 1)} left`));
        return `<w:tr>${left}${marker(name, at + 1)}</w:tr>`;
      }).join(""),
    );

  // A line of its own either side of each table, so where the table starts and where
  // it ends are read as well as the distance from row to row.
  const block = (name: string, content: string): string =>
    paragraph(`<w:pageBreakBefore/>${exactly(LINE_PT)}`, run(`case ${name}`)) +
    content +
    paragraph(exactly(LINE_PT), run(`${name} after`));

  // Three grid columns where the first two are spanned by one cell, so that the left
  // of the third column's text says what a span is worth. The second row spans
  // nothing, and is what the first is read against.
  const SPAN_TWIPS = [1440, 1440, 2880] as const;

  const spanTable = (name: string, statesItsWidth: boolean): string =>
    table(
      SPAN_TWIPS,
      `<w:tr>${cell(
        `${statesItsWidth ? width(SPAN_TWIPS[0] + SPAN_TWIPS[1]) : ""}<w:gridSpan w:val="2"/>`,
        line(`${name}1 spanning`),
      )}${cell(width(SPAN_TWIPS[2]), line(`${name}1 third`))}</w:tr>` +
        `<w:tr>${cell(width(SPAN_TWIPS[0]), line(`${name}2 first`))}${cell(
          width(SPAN_TWIPS[1]),
          line(`${name}2 second`),
        )}${cell(width(SPAN_TWIPS[2]), line(`${name}2 third`))}</w:tr>`,
    );

  return [
    // One line in a cell merged down the whole table. Says where the text of a merged
    // cell sits, and whether the cells it swallowed are worth anything at all: four
    // rows 20pt apart say they are not.
    block("a", mergedTable("a", 1, ROWS)),
    // Content that fits inside the merge with room to spare, which is the case a real
    // document is mostly made of.
    block("b", mergedTable("b", 3, ROWS)),
    // Six 20pt lines in a merge worth 80pt. The 40pt over says where the room a merge
    // is short comes from: rows at 20, 20, 20, 60 say it is the last row of the merge,
    // and rows at 30 each say it is shared out.
    block("c", mergedTable("c", 6, ROWS)),
    // The same again at ten lines, so that a rule read off one case is read again at
    // three times the deficit.
    block("d", mergedTable("d", 10, ROWS)),
    // A merge closing at the second row with two ordinary rows under it, and six lines
    // in it. Read against c, this says whether what a merge is short falls on the last
    // row of the merge or on the last row of the table.
    block("e", mergedTable("e", 6, 2)),
    // One line in a merge asking to be seated in the middle of itself, which says
    // whether a merged cell is seated in its own row or in the whole run of them.
    block("f", mergedTable("f", 1, ROWS, { restart: `<w:vAlign w:val="center"/>` })),
    // A swallowed cell holding a 40pt line of its own. If Word draws it, or if the row
    // grows to hold it, a continuation cell is not the nothing the rest of this
    // assumes.
    block("g", mergedTable("g", 1, ROWS, { swallowedContent: line("g spare", 40) })),
    // A cell spanning two grid columns and stating no width of its own, which says
    // whether a span is worth the columns under it.
    block("h", spanTable("h", false)),
    // And the same span stating the width of both columns, which is how a real
    // document writes it.
    block("i", spanTable("i", true)),
    // A table is the last thing in the body, which Word will not have.
    EMPTY,
  ].join("");
}

// How much room the first line of a numbered paragraph has.
//
// The number sits at the hanging position and the text after it starts wherever the
// level's suffix moves on to, which `startOfText` already reads off the paragraph's
// own tab stops. What nothing has asked is how wide the line that text stands on
// then is. What is built gives it the room between the two indents, on the reading
// that a number hanging in front of the text leaves the text starting at the left
// indent like every line under it.
//
// **A real document breaks that reading.** One of the 966 states stops at 907 and
// 1441 twips against a left indent of 1441 and a hanging of 744: its number tabs to
// the first stop, which is 26.7pt short of the indent, so its first line starts
// there and runs the whole way to the right indent. Word fits 29 characters on that
// line where this project fits 11.
//
// Each case is a marker, the case, and a plain line under it. The words are three
// characters each and numbered, so which of them Word put on the first line says
// where that line ended to within about 19pt, and the readings being compared are
// 36pt apart.
export function numberedFirstLineDocument(): string {
  // A column 162pt wide from the text start and 126 from the left indent, which is
  // the whole of what the cases disagree about.
  const LEFT_TWIPS = 2880;
  const HANGING_TWIPS = 1440;
  const RIGHT_TWIPS = 5400;
  // Short of the left indent by 36pt, and past where any of the numbers end.
  const STOP_TWIPS = 2160;

  const words = (name: string): string =>
    Array.from({ length: 24 }, (_, at) => `${name}${String(at + 1).padStart(2, "0")}`).join(" ");

  const indent = (hangingTwips: number, firstLineTwips: number | null): string =>
    `<w:ind w:left="${String(LEFT_TWIPS)}" w:right="${String(RIGHT_TWIPS)}"` +
    (firstLineTwips === null
      ? ` w:hanging="${String(hangingTwips)}"`
      : ` w:firstLine="${String(firstLineTwips)}"`) +
    `/>`;

  const stops = (positions: readonly number[]): string =>
    positions.length === 0
      ? ""
      : `<w:tabs>${positions.map((at) => `<w:tab w:val="left" w:pos="${String(at)}"/>`).join("")}</w:tabs>`;

  type Case = {
    readonly name: string;
    readonly numId: number | null;
    readonly stops: readonly number[];
    readonly firstLineTwips?: number;
  };

  const block = (of: Case): string =>
    paragraph(`<w:pageBreakBefore/>`, run(`case ${of.name}`)) +
    paragraph(
      (of.numId === null
        ? ""
        : `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${String(of.numId)}"/></w:numPr>`) +
        stops(of.stops) +
        indent(HANGING_TWIPS, of.firstLineTwips ?? null),
      run(words(of.name)),
    ) +
    paragraph("", run(`${of.name} after`));

  return [
    // A stop short of the left indent, which is what the real document holds: the
    // number tabs to it and the text starts 36pt in front of the indent. Whether
    // the line ends at the right indent or 36pt short of it is the whole question.
    block({ name: "a", numId: 1, stops: [STOP_TWIPS] }),
    // The same with no stop of its own, so the number's tab moves on to the implicit
    // one at the left indent. This is the case the built reading assumes, and it has
    // to come out unchanged.
    block({ name: "b", numId: 1, stops: [] }),
    // A number wider than the room in front of the indent, so the text is pushed
    // past the indent instead of standing short of it. Says whether a first line
    // loses room as readily as it gains it.
    block({ name: "c", numId: 2, stops: [] }),
    // A suffix that puts a single space after the number rather than a tab, so the
    // text starts wherever the number ended and no stop is consulted at all.
    block({ name: "d", numId: 3, stops: [] }),
    // And a suffix that puts nothing after it.
    block({ name: "e", numId: 4, stops: [] }),
    // The same indents with no number at all, which is the rule this one is read
    // against: a hanging indent leaves its first line wider than the rest.
    block({ name: "f", numId: null, stops: [] }),
    // And a first line indented the other way, which leaves it narrower.
    block({ name: "g", numId: null, stops: [], firstLineTwips: 720 }),
    EMPTY,
  ].join("");
}

// One list a case, so that a suffix and a width of number can be varied without
// disturbing the others. Every level states the same indents the paragraphs do, so
// nothing here turns on which of the two Word reads.
export const FIRST_LINE_NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${[
    { id: 0, text: "%1.", suffix: null },
    { id: 1, text: "Section %1.%1.%1:", suffix: null },
    { id: 2, text: "%1.", suffix: "space" },
    { id: 3, text: "%1.", suffix: "nothing" },
  ]
    .map(
      (each) => `<w:abstractNum w:abstractNumId="${String(each.id)}">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
      <w:lvlText w:val="${each.text}"/><w:lvlJc w:val="left"/>
      ${each.suffix === null ? "" : `<w:suff w:val="${each.suffix}"/>`}
      <w:pPr><w:ind w:left="2880" w:right="5400" w:hanging="1440"/></w:pPr></w:lvl>
  </w:abstractNum>`,
    )
    .join("")}
  ${[0, 1, 2, 3]
    .map(
      (id) => `<w:num w:numId="${String(id + 1)}"><w:abstractNumId w:val="${String(id)}"/></w:num>`,
    )
    .join("")}
</w:numbering>`;
