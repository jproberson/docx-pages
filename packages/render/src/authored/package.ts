import { strToU8, zipSync } from "fflate";

// A complete word processing package, which the fixture builder in core is not:
// that one carries only what this project's own reader looks at, and Word will not
// open a file whose parts are not declared and related.

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WPG_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const WP14_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";

const DOCUMENT_NAMESPACES = [
  `xmlns:w="${W_NS}"`,
  `xmlns:r="${R_NS}"`,
  `xmlns:wp="${WP_NS}"`,
  `xmlns:a="${A_NS}"`,
  `xmlns:wps="${WPS_NS}"`,
  `xmlns:wpg="${WPG_NS}"`,
  `xmlns:mc="${MC_NS}"`,
  `xmlns:w14="${W14_NS}"`,
  `xmlns:wp14="${WP14_NS}"`,
  `mc:Ignorable="w14 wp14"`,
].join(" ");

const OFFICE = "application/vnd.openxmlformats-officedocument.wordprocessingml";

const contentTypes = (
  picture: boolean,
  footer: boolean,
): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${picture ? `<Default Extension="png" ContentType="image/png"/>` : ""}
  <Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="${OFFICE}.settings+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="${OFFICE}.numbering+xml"/>
  ${footer ? `<Override PartName="/word/footer1.xml" ContentType="${OFFICE}.footer+xml"/>` : ""}
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`;

// Every drawing in a document points at the one picture part, since what is being
// asked about a drawing is the room it takes rather than what is in it.
export const PICTURE_ID = "rId4";

const FOOTER_ID = "rId5";

const documentRels = (
  picture: boolean,
  footer: boolean,
): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${R_NS}/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="${R_NS}/numbering" Target="numbering.xml"/>
  ${picture ? `<Relationship Id="${PICTURE_ID}" Type="${R_NS}/image" Target="media/picture.png"/>` : ""}
  ${footer ? `<Relationship Id="${FOOTER_ID}" Type="${R_NS}/footer" Target="footer1.xml"/>` : ""}
</Relationships>`;

// One opaque pixel, stretched to whatever extent a drawing states. Written out
// rather than copied from anywhere: header, an 1x1 RGB IHDR, the pixel deflated,
// and the end marker.
const PICTURE_PNG =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR42mNwSFgAAAIkAUEotsXkAAAAAElFTkSuQmCC";

// Letter, half-inch margins all round, so a position on the page is easy to read:
// the text runs from 36pt to 576pt across and starts 36pt down.
export const PAGE = `<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>
  <w:cols w:space="720"/>
</w:sectPr>`;

export const LEFT_PT = 36;
export const RIGHT_PT = 576;
export const TOP_PT = 36;

// Every document is laid out in one face, so a measurement is never a question
// about which font Word chose. Calibri ships with Word and this project's builtin
// metrics answer for it.
export const FACE = "Calibri";

// What a document says about itself. Every authored document declares the
// compatibility mode Word writes today unless it is asking what changes without
// one, and the tab stop Word itself would have written unless it is asking whether
// that number is read at all.
export type AuthoredSettings = {
  readonly defaultTabStopTwips?: number;
  readonly compatibilityMode?: number | null;
};

export function settingsPart(settings: AuthoredSettings = {}): string {
  const mode = settings.compatibilityMode === undefined ? 15 : settings.compatibilityMode;
  const compat =
    mode === null
      ? ""
      : `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${String(mode)}"/></w:compat>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W_NS}">
  <w:defaultTabStop w:val="${String(settings.defaultTabStopTwips ?? 720)}"/>
  ${compat}
</w:settings>`;
}

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="${FACE}" w:hAnsi="${FACE}" w:cs="${FACE}"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <!--EXTRA-->
</w:styles>`;

const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W_NS}"/>`;

export type AuthoredParts = {
  // Everything between <w:body> and its section properties.
  readonly body: string;
  // The blocks of a footer drawn on every page, which is what makes the foot of
  // the text a quantity of the document's rather than the margin's: Word holds the
  // text off a footer taller than the room the margin left it.
  readonly footer?: string;
  // Styles of the document's own, which stand beside the defaults every authored
  // document shares rather than replacing them.
  readonly extraStyles?: string;
  readonly settings?: string;
  readonly numbering?: string;
  // Whether the document carries the picture part its drawings point at.
  readonly picture?: boolean;
  // Whether the document states on its own root that the whitespace under it is the
  // text's own. A document in the wild states it there and on no `w:t` at all.
  readonly preservesSpace?: boolean;
};

export function buildAuthoredDocx(parts: AuthoredParts): Uint8Array {
  const footer = parts.footer !== undefined;
  const page = footer
    ? PAGE.replace(
        "<w:sectPr>",
        `<w:sectPr><w:footerReference w:type="default" r:id="${FOOTER_ID}"/>`,
      )
    : PAGE;
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DOCUMENT_NAMESPACES}${parts.preservesSpace === true ? ` xml:space="preserve"` : ""}>
  <w:body>${parts.body}${page}</w:body>
</w:document>`;

  const picture = parts.picture === true;
  return zipSync({
    "[Content_Types].xml": strToU8(contentTypes(picture, footer)),
    "_rels/.rels": strToU8(ROOT_RELS),
    "word/_rels/document.xml.rels": strToU8(documentRels(picture, footer)),
    "word/document.xml": strToU8(document),
    "word/styles.xml": strToU8(STYLES.replace("<!--EXTRA-->", parts.extraStyles ?? "")),
    "word/settings.xml": strToU8(parts.settings ?? settingsPart()),
    "word/numbering.xml": strToU8(parts.numbering ?? NUMBERING),
    ...(picture ? { "word/media/picture.png": Buffer.from(PICTURE_PNG, "base64") } : {}),
    ...(parts.footer === undefined
      ? {}
      : {
          "word/footer1.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr ${DOCUMENT_NAMESPACES}>${parts.footer}</w:ftr>`),
        }),
  });
}
