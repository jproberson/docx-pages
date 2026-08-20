import { strFromU8, unzipSync } from "fflate";

import { DocxPagesError } from "../errors.js";
import { parseXml, type XmlElement } from "./xml.js";

export const MAIN_DOCUMENT_PART = "word/document.xml";

export type DocxPackage = {
  readonly parts: ReadonlyMap<string, Uint8Array>;
};

export function openDocx(bytes: Uint8Array): DocxPackage {
  let entries: Record<string, Uint8Array>;
  try {
    entries = unzipSync(bytes);
  } catch (error: unknown) {
    throw new DocxPagesError({
      code: "docx-unreadable",
      message: "the bytes are not a readable zip archive",
      at: "core/docx/package.openDocx",
      context: { byteLength: bytes.byteLength },
      cause: error,
    });
  }

  const parts = new Map(Object.entries(entries));
  if (!parts.has(MAIN_DOCUMENT_PART)) {
    throw new DocxPagesError({
      code: "docx-malformed",
      message: "the archive has no main document part",
      at: "core/docx/package.openDocx",
      context: { part: MAIN_DOCUMENT_PART, partNames: [...parts.keys()].sort() },
    });
  }

  return { parts };
}

export function partText(pkg: DocxPackage, part: string): string {
  const bytes = pkg.parts.get(part);
  if (bytes === undefined) {
    throw new DocxPagesError({
      code: "docx-malformed",
      message: "the archive has no such part",
      at: "core/docx/package.partText",
      context: { part, partNames: [...pkg.parts.keys()].sort() },
    });
  }
  return strFromU8(bytes);
}

export function partXml(pkg: DocxPackage, part: string): XmlElement {
  const text = partText(pkg, part);
  // The parser answers nothing for some malformed input and throws for the rest, a
  // part cut off in the middle of a tag among them. Both are the same fault in bytes
  // a caller was handed rather than wrote, and both leave here as the same error, or
  // a caller telling a document it refuses from a fault of its own reads one as the
  // other.
  let root: XmlElement | null;
  try {
    root = parseXml(text);
  } catch (error: unknown) {
    throw new DocxPagesError({
      code: "docx-malformed",
      message: "the part is not readable xml",
      at: "core/docx/package.partXml",
      context: { part },
      cause: error,
    });
  }
  if (root === null) {
    throw new DocxPagesError({
      code: "docx-malformed",
      message: "the part has no root element",
      at: "core/docx/package.partXml",
      context: { part },
    });
  }
  return root;
}
