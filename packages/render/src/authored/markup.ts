// The markup the authored documents share, written out once so that a change Word
// starts asking for lands in one place. What each document states for itself is
// what it is asking about; everything else here is the page the whole suite uses.

export const emu = (pt: number): string => String(Math.round(pt * 12700));

// A line told exactly how tall to be, which is what makes where the next one lands
// arithmetic rather than a measurement of a face.
export const exactLine = (pt: number): string =>
  `<w:spacing w:line="${String(pt * 20)}" w:lineRule="exact"/>`;

// The same, with whatever a style would add above and below closed up, for a
// document whose whole answer is the distance between one line and the next.
export const exactLineClosedUp = (pt: number): string =>
  `<w:spacing w:before="0" w:after="0" w:line="${String(pt * 20)}" w:lineRule="exact"/>`;

export type AnchoredTextBox = {
  readonly id: number;
  readonly name: string;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly wrap: string;
  readonly leftPt?: number;
  readonly offsetPt?: number;
  // What the box holds. A box of a stated size holding one line is what makes the
  // object readable at all: a picture says only where it was drawn, and a pdf says
  // nothing about which anchor drew it.
  readonly content: string;
};

// An anchored text box, which is how every document asking about a floating object
// states one. The whitespace inside the literal is part of what Word was measured
// against, so it is written out here exactly as each document used to write it.
export const anchoredTextBox = ({
  id,
  name,
  widthPt,
  heightPt,
  wrap,
  leftPt = 0,
  offsetPt = 0,
  content,
}: AnchoredTextBox): string =>
  `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
      <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${String(id)}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>${emu(leftPt)}</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(offsetPt)}</wp:posOffset></wp:positionV>
        <wp:extent cx="${emu(widthPt)}" cy="${emu(heightPt)}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        ${wrap}
        <wp:docPr id="${String(id)}" name="${name}"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:cNvSpPr txBox="1"/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(widthPt)}" cy="${emu(heightPt)}"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr>
            <wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>
            <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"/>
          </wps:wsp></a:graphicData></a:graphic>
      </wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>`;

// The single column of the page every authored document uses, half an inch of gap
// and all.
export const ONE_COLUMN = `<w:cols w:space="720"/>`;

export type BodySection = {
  readonly type?: string;
  readonly topTwips?: number;
  readonly leftTwips?: number;
  readonly bottomTwips?: number;
  readonly columns?: string;
};

// The body's own page, written out again: a section stating anything else would be
// asking a second question, so a document states only the twips it is about.
export const bodySectionProperties = ({
  type = "",
  topTwips = 720,
  leftTwips = 720,
  bottomTwips = 720,
  columns = "",
}: BodySection = {}): string =>
  `<w:sectPr>` +
  (type === "" ? "" : `<w:type w:val="${type}"/>`) +
  `<w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="${String(topTwips)}" w:right="720" w:bottom="${String(bottomTwips)}" w:left="${String(leftTwips)}" w:header="720" w:footer="720" w:gutter="0"/>` +
  columns +
  `</w:sectPr>`;
