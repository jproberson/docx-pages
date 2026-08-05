import { DocxPagesError } from "../errors.js";
import { MAIN_DOCUMENT_PART, partXml, type DocxPackage } from "./package.js";
import { attribute, descendantsNamed, firstNamed, type XmlElement } from "./xml.js";

export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export type PageMargin = {
  readonly topTwips: number;
  readonly rightTwips: number;
  readonly bottomTwips: number;
  readonly leftTwips: number;
  readonly headerTwips: number;
  readonly footerTwips: number;
};

export type SectionGeometry = {
  readonly widthTwips: number;
  readonly heightTwips: number;
  readonly margin: PageMargin;
};

const AT = "core/docx/section.readSectionGeometry";

function twips(element: XmlElement | null, name: string, fallback: number): number {
  if (element === null) return fallback;
  const raw = attribute(element, W_NS, name);
  if (raw === undefined) return fallback;
  const value = Number(raw);
  if (!Number.isFinite(value)) {
    throw new DocxPagesError({
      code: "docx-malformed",
      message: "section dimension is not a number",
      at: AT,
      context: { part: MAIN_DOCUMENT_PART, element: element.name, attribute: name, value: raw },
    });
  }
  return value;
}

// The page one section makes, as everything here reads it: the sheet, the margins
// round the text and the room kept for a header and a footer. Two sections whose
// signatures match lay their pages out identically, whatever else they differ in,
// so reading the last of them loses nothing.
//
// A section break far more often changes a header or a column count than the page
// itself, and this is what tells those apart.
export function pageGeometrySignature(section: XmlElement): string {
  const size = firstNamed(section, W_NS, "pgSz");
  const margin = firstNamed(section, W_NS, "pgMar");
  const of = (element: XmlElement | null, names: readonly string[]): string =>
    names.map((name) => (element === null ? "" : (attribute(element, W_NS, name) ?? ""))).join(",");

  return [
    of(size, ["w", "h", "orient"]),
    of(margin, ["top", "right", "bottom", "left", "header", "footer", "gutter"]),
  ].join("|");
}

export function readSectionGeometry(pkg: DocxPackage): SectionGeometry {
  const root = partXml(pkg, MAIN_DOCUMENT_PART);
  const body = firstNamed(root, W_NS, "body");
  const sections = body === null ? [] : descendantsNamed(body, W_NS, "sectPr");
  const section = sections.at(-1);
  if (section === undefined) {
    throw new DocxPagesError({
      code: "docx-malformed",
      message: "the body has no section properties",
      at: AT,
      context: { part: MAIN_DOCUMENT_PART },
    });
  }

  const size = firstNamed(section, W_NS, "pgSz");
  const margin = firstNamed(section, W_NS, "pgMar");

  return {
    widthTwips: twips(size, "w", 0),
    heightTwips: twips(size, "h", 0),
    margin: {
      topTwips: twips(margin, "top", 0),
      rightTwips: twips(margin, "right", 0),
      bottomTwips: twips(margin, "bottom", 0),
      leftTwips: twips(margin, "left", 0),
      headerTwips: twips(margin, "header", 0),
      footerTwips: twips(margin, "footer", 0),
    },
  };
}
