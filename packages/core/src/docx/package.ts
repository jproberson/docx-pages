import { strFromU8, unzipSync } from "fflate";

import { OnePagerError } from "../errors.js";
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
    throw new OnePagerError({
      code: "docx-unreadable",
      message: "the bytes are not a readable zip archive",
      at: "core/docx/package.openDocx",
      context: { byteLength: bytes.byteLength },
      cause: error,
    });
  }

  const parts = new Map(Object.entries(entries));
  if (!parts.has(MAIN_DOCUMENT_PART)) {
    throw new OnePagerError({
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
    throw new OnePagerError({
      code: "docx-malformed",
      message: "the archive has no such part",
      at: "core/docx/package.partText",
      context: { part, partNames: [...pkg.parts.keys()].sort() },
    });
  }
  return strFromU8(bytes);
}

export function partXml(pkg: DocxPackage, part: string): XmlElement {
  const root = parseXml(partText(pkg, part));
  if (root === null) {
    throw new OnePagerError({
      code: "docx-malformed",
      message: "the part has no root element",
      at: "core/docx/package.partXml",
      context: { part },
    });
  }
  return root;
}
