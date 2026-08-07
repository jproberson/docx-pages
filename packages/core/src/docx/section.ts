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

// How a section stands against the one before it. Word writes no `w:type` for the
// commonest of them, so a section saying nothing is one starting a page.
export type SectionBreak = "nextPage" | "continuous" | "evenPage" | "oddPage" | "nextColumn";

const BREAKS: readonly SectionBreak[] = [
  "nextPage",
  "continuous",
  "evenPage",
  "oddPage",
  "nextColumn",
];

export type DocumentSection = {
  readonly geometry: SectionGeometry;
  readonly breakKind: SectionBreak;
  // How many columns the section's text runs in. One unless it says otherwise.
  readonly columns: number;
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

// Every section a document is made of, in the order they run.
//
// A `w:sectPr` on a paragraph ends the section that paragraph is the last of, and
// the one standing at the end of the body governs the text after the last of those.
// So the sections are the `sectPr` elements in document order, and the body's own
// is the final one.
//
// Reading only the last of them, which is all this did until now, gives the whole
// document the geometry of its final section. Where the sections differ that puts
// every page above the last break on the wrong page: measured against Word's own
// pdf over a corpus, documents with more than one section place one line in eleven
// where documents with one place two in three.
export function readSections(pkg: DocxPackage): readonly DocumentSection[] {
  const root = partXml(pkg, MAIN_DOCUMENT_PART);
  const body = firstNamed(root, W_NS, "body");
  const sections = body === null ? [] : descendantsNamed(body, W_NS, "sectPr");
  if (sections.length === 0) throw noSection();

  return sections.map((section) => ({
    geometry: geometryOf(section),
    breakKind: breakOf(section),
    columns: columnsOf(section),
  }));
}

// A paragraph carrying section properties is the last paragraph of its section.
// Only a paragraph standing in the body itself can be: one inside a table cell
// carries them for the story in that cell and ends nothing.
export const endsASection = (paragraph: XmlElement): boolean => {
  return sectionClosedBy(paragraph) !== null;
};

const sectionClosedBy = (paragraph: XmlElement): XmlElement | null => {
  const properties = firstNamed(paragraph, W_NS, "pPr");
  return properties === null ? null : firstNamed(properties, W_NS, "sectPr");
};

// What the break a paragraph carries does to the page under it.
export type SectionClose = {
  // Whether a page opens after the paragraph. That is the type stated by the
  // section beginning under it and never the one stated by the section it closes.
  readonly opensAPage: boolean;
};

// Every paragraph of the body closing a section, and what its break does.
//
// A break is read at the section it opens rather than at the one it closes: a
// section's `w:type` says how that section begins against the one before it, so
// the type deciding a break is the one stated by the section under it. The last
// paragraph carrying section properties is followed by the body's own section,
// which is the final one and states a type like any other.
//
// Measured on 2026-08-07 by the authored `section-pages` document, whose every
// section is the same page to the twip so that only the break can move a page.
// Of the five types, only `continuous` carries on down the page already open.
// `nextPage` and a section stating no type at all each open one, and so does
// `nextColumn` where the section runs in a single column. `evenPage` and
// `oddPage` open one too, and leave a blank page behind where the page they
// reach for is not the next one: the section after a page 4 asking for an even
// page opened page 6.
export function sectionsClosedBy(
  pkg: DocxPackage,
  bodyParagraphs: readonly XmlElement[],
): ReadonlyMap<XmlElement, SectionClose> {
  const body = firstNamed(partXml(pkg, MAIN_DOCUMENT_PART), W_NS, "body");
  const closers = bodyParagraphs.filter(endsASection);
  const closes = new Map<XmlElement, SectionClose>();

  for (const [at, closer] of closers.entries()) {
    const next = closers[at + 1];
    const opening =
      next === undefined
        ? body === null
          ? null
          : firstNamed(body, W_NS, "sectPr")
        : sectionClosedBy(next);
    closes.set(closer, {
      opensAPage: opening !== null && breakOf(opening) !== "continuous",
    });
  }
  return closes;
}

function breakOf(section: XmlElement): SectionBreak {
  const stated = firstNamed(section, W_NS, "type");
  const value = stated === null ? undefined : attribute(stated, W_NS, "val");
  return BREAKS.find((each) => each === value) ?? "nextPage";
}

function columnsOf(section: XmlElement): number {
  const columns = firstNamed(section, W_NS, "cols");
  if (columns === null) return 1;
  const stated = Number(attribute(columns, W_NS, "num") ?? "1");
  return Number.isFinite(stated) && stated >= 1 ? Math.floor(stated) : 1;
}

const noSection = (): DocxPagesError =>
  new DocxPagesError({
    code: "docx-malformed",
    message: "the body has no section properties",
    at: AT,
    context: { part: MAIN_DOCUMENT_PART },
  });

function geometryOf(section: XmlElement): SectionGeometry {
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

// The page the document's final section makes, which is the one every story that
// is not the body is measured against.
export function readSectionGeometry(pkg: DocxPackage): SectionGeometry {
  const sections = readSections(pkg);
  const last = sections.at(-1);
  if (last === undefined) throw noSection();
  return last.geometry;
}
