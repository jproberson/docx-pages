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
