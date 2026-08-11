import { DocxPagesError } from "../errors.js";
import { MAIN_DOCUMENT_PART, partXml, type DocxPackage } from "./package.js";
import { readRelationships, R_NS } from "./relationships.js";
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

// How a section divides its frame between its columns. Word writes one of two
// things: a count and one gap for every column, or a width and a gap for each
// column in turn. Every one of the sixteen corpus documents that hold columns
// writes the second, and none of their columns is the same width as its neighbour.
export type SectionColumns = {
  // How many columns the section's text runs in. One unless it says otherwise.
  readonly count: number;
  // The width and the gap after each, where the section stated them; empty where
  // it asked for equal widths and left them to be divided.
  readonly widthsTwips: readonly number[];
  readonly gapsTwips: readonly number[];
  // The gap a section of equal columns keeps between all of them.
  readonly spaceTwips: number;
};

export type DocumentSection = {
  readonly geometry: SectionGeometry;
  readonly breakKind: SectionBreak;
  readonly columns: SectionColumns;
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
  const closes = new Map<XmlElement, SectionClose>();
  const sections = bodySections(pkg, bodyParagraphs);

  for (const [at, section] of sections.entries()) {
    const under = sections[at + 1];
    if (section.endsAt === null || under === undefined) continue;
    closes.set(section.endsAt, { opensAPage: under.breakKind !== "continuous" });
  }
  return closes;
}

// The header or footer parts a section names, by the page of it each is drawn on.
// Null where the section names none of that kind, which is a page that draws no
// header at all rather than one falling back to another.
export type SectionStories = {
  readonly first: string | null;
  readonly default: string | null;
  readonly even: string | null;
};

export const NO_STORIES: SectionStories = { first: null, default: null, even: null };

// One of the body's own sections, and the paragraph it ends at. The final section
// ends at nothing: the body's own `w:sectPr` governs whatever follows the last
// paragraph carrying one.
export type BodySection = DocumentSection & {
  readonly endsAt: XmlElement | null;
  readonly headers: SectionStories;
  readonly footers: SectionStories;
  // Whether the section draws something of its own on the page it opens, which is
  // what `w:titlePg` says. 408 of the 718 corpus documents state it.
  readonly titlePage: boolean;
};

// **A page draws its own section's header, and the page a section opens draws that
// section's first-page one where it says `w:titlePg`.** What stood here read the
// last `w:headerReference` of type `default` anywhere in the part, which is one
// section's answer given to every page of the document: a corpus document whose
// only page is in a section naming nothing but a first-page header was drawn under
// the *next* section's default header, a full-page background image, and lost the
// logo and rule its own header holds.
//
// **A section stating no reference of a kind inherits that kind from the section
// before it**, which is how a document writes one header once and runs it through
// to the end.
//
// `even` is read and never chosen. Not one of the 718 corpus documents states
// `w:evenAndOddHeaders`, though 521 of them name an even-page part, so choosing one
// would be acting on a setting nothing has turned on and nothing has measured.
export function storyFor(
  stories: SectionStories,
  opensItsSection: boolean,
  titlePage: boolean,
): string | null {
  return opensItsSection && titlePage ? stories.first : stories.default;
}

const storiesOf = (
  section: XmlElement,
  reference: string,
  partOf: (relationshipId: string) => string | null,
  inherited: SectionStories,
): SectionStories => {
  const stated = new Map<string, string | null>();
  for (const node of descendantsNamed(section, W_NS, reference)) {
    const id = attribute(node, R_NS, "id");
    if (id === undefined) continue;
    stated.set(attribute(node, W_NS, "type") ?? "default", partOf(id));
  }
  const of = (kind: keyof SectionStories): string | null =>
    stated.has(kind) ? (stated.get(kind) ?? null) : inherited[kind];
  return { first: of("first"), default: of("default"), even: of("even") };
};

// The body's sections in the order they run, read off the paragraphs standing in
// the body itself rather than off every `w:sectPr` in the part: one inside a table
// cell governs the story in that cell and closes no section of the body's.
export function bodySections(
  pkg: DocxPackage,
  bodyParagraphs: readonly XmlElement[],
): readonly BodySection[] {
  const body = firstNamed(partXml(pkg, MAIN_DOCUMENT_PART), W_NS, "body");
  const own = body === null ? null : firstNamed(body, W_NS, "sectPr");
  const closers = bodyParagraphs.filter(endsASection);

  const relationships = readRelationships(pkg, MAIN_DOCUMENT_PART);
  const partOf = (relationshipId: string): string | null => {
    const target = relationships.get(relationshipId)?.part;
    return target !== undefined && pkg.parts.has(target) ? target : null;
  };

  let headers = NO_STORIES;
  let footers = NO_STORIES;
  const of = (element: XmlElement, endsAt: XmlElement | null): BodySection => {
    headers = storiesOf(element, "headerReference", partOf, headers);
    footers = storiesOf(element, "footerReference", partOf, footers);
    return {
      geometry: geometryOf(element),
      breakKind: breakOf(element),
      columns: columnsOf(element),
      endsAt,
      headers,
      footers,
      titlePage: firstNamed(element, W_NS, "titlePg") !== null,
    };
  };

  const sections = closers.flatMap((closer) => {
    const properties = sectionClosedBy(closer);
    return properties === null ? [] : [of(properties, closer)];
  });
  return own === null ? sections : [...sections, of(own, null)];
}

function breakOf(section: XmlElement): SectionBreak {
  const stated = firstNamed(section, W_NS, "type");
  const value = stated === null ? undefined : attribute(stated, W_NS, "val");
  return BREAKS.find((each) => each === value) ?? "nextPage";
}

const ONE_COLUMN: SectionColumns = {
  count: 1,
  widthsTwips: [],
  gapsTwips: [],
  spaceTwips: 0,
};

function columnsOf(section: XmlElement): SectionColumns {
  const columns = firstNamed(section, W_NS, "cols");
  if (columns === null) return ONE_COLUMN;

  const stated = Number(attribute(columns, W_NS, "num") ?? "1");
  const count = Number.isFinite(stated) && stated >= 1 ? Math.floor(stated) : 1;
  const each = columns.children.filter((child) => child.namespace === W_NS && child.name === "col");
  const number = (element: XmlElement, name: string): number => {
    const value = Number(attribute(element, W_NS, name) ?? Number.NaN);
    return Number.isFinite(value) ? value : 0;
  };

  // A section asking for equal widths is divided rather than read, whatever widths
  // it wrote beside the request.
  const divided = attribute(columns, W_NS, "equalWidth") === "0" && each.length >= count;

  return {
    count,
    widthsTwips: divided ? each.slice(0, count).map((column) => number(column, "w")) : [],
    gapsTwips: divided ? each.slice(0, count).map((column) => number(column, "space")) : [],
    spaceTwips: number(columns, "space"),
  };
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
