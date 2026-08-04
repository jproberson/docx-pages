import { zipSync, strToU8 } from "fflate";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

export const WORDPROCESSING_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export function wordDocument(bodyXml: string, prefix = "w"): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<${prefix}:document xmlns:${prefix}="${WORDPROCESSING_NS}">
  <${prefix}:body>${bodyXml}</${prefix}:body>
</${prefix}:document>`;
}

export type DocxPart = string | Uint8Array;

export function buildDocx(parts: Readonly<Record<string, DocxPart>>): Uint8Array {
  const entries: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "_rels/.rels": strToU8(ROOT_RELS),
  };
  for (const [name, content] of Object.entries(parts)) {
    entries[name] = typeof content === "string" ? strToU8(content) : content;
  }
  return zipSync(entries);
}
