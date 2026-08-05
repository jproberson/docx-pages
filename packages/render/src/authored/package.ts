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

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="${OFFICE}.settings+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="${OFFICE}.numbering+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${R_NS}/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="${R_NS}/numbering" Target="numbering.xml"/>
</Relationships>`;

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

const SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W_NS}">
  <w:defaultTabStop w:val="720"/>
  <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>`;

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
</w:styles>`;

const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W_NS}"/>`;

export type AuthoredParts = {
  // Everything between <w:body> and its section properties.
  readonly body: string;
  readonly styles?: string;
  readonly settings?: string;
  readonly numbering?: string;
};

export function buildAuthoredDocx(parts: AuthoredParts): Uint8Array {
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DOCUMENT_NAMESPACES}>
  <w:body>${parts.body}${PAGE}</w:body>
</w:document>`;

  return zipSync({
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "_rels/.rels": strToU8(ROOT_RELS),
    "word/_rels/document.xml.rels": strToU8(DOCUMENT_RELS),
    "word/document.xml": strToU8(document),
    "word/styles.xml": strToU8(parts.styles ?? STYLES),
    "word/settings.xml": strToU8(parts.settings ?? SETTINGS),
    "word/numbering.xml": strToU8(parts.numbering ?? NUMBERING),
  });
}
