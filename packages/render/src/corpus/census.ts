import { createHash } from "node:crypto";

import {
  attribute,
  openDocx,
  pageGeometrySignature,
  partXml,
  MAIN_DOCUMENT_PART,
  W_NS,
  WP_NS,
  type DocxPackage,
  type XmlElement,
} from "@docx-pages/core";

// What is actually inside a corpus of documents, as against what this project
// cannot honour in one.
//
// The fidelity report answers the second question and says nothing about the
// first: a document full of tables, floats and numbered lists that this project
// draws correctly reports nothing at all. So a gap met in two hundred documents
// cannot be weighed without knowing how many documents had the feature to begin
// with, and a sample cannot be chosen to cover what the corpus holds.
//
// A census needs no fonts and no layout: it opens the package and counts. That
// makes it cheap enough to run over a whole corpus whenever the question comes up.
//
// **It records counts and nothing else.** No text, no file names, no font names:
// how many distinct faces a document asks for is a number, and which ones it asks
// for is the document's business. See the note in `sweep.ts` about identity.

export type DocumentCensus = {
  readonly id: string;
  readonly bytes: number;
  // Every feature met, and how many times. A feature met nowhere is left out
  // rather than written as nought, so a profile is the keys of this.
  readonly counts: Readonly<Record<string, number>>;
};

// The namespaces a census has to tell apart. A drawing says what it is by the
// namespace of the thing inside it rather than by any name of its own.
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WPG_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

type Tally = Map<string, number>;

const count = (tally: Tally, feature: string, by = 1): void => {
  tally.set(feature, (tally.get(feature) ?? 0) + by);
};

// Word's own attributes are in its namespace; the ones a drawing states about its
// wrapping are in none.
const stated = (element: XmlElement, name: string): string | undefined =>
  attribute(element, W_NS, name) ?? attribute(element, "", name);

const toggledOn = (element: XmlElement): boolean => {
  const value = stated(element, "val");
  return value === undefined || (value !== "0" && value !== "false" && value !== "off");
};

// What one element of the flow says about the document. Named for what a reader of
// the report would want to know rather than for the tag: `w:tbl` inside a cell is
// a nested table, which is a different problem from a table.
function census(element: XmlElement, tally: Tally, inCell: boolean, gathered: Gathered): void {
  if (element.namespace === WP_NS) {
    if (element.name === "inline") count(tally, "inline-drawing");
    if (element.name === "anchor") count(tally, "floating-drawing");
    if (WRAPS.has(element.name)) {
      count(tally, `wrap-${WRAPS.get(element.name) ?? "square"}`);
      const side = stated(element, "wrapText");
      if (side !== undefined && side !== "bothSides") count(tally, `wrap-side-${side}`);
    }
    return;
  }

  if (element.namespace === PIC_NS && element.name === "pic") count(tally, "picture");
  if (element.namespace === WPS_NS && element.name === "wsp") count(tally, "shape");
  if (element.namespace === WPS_NS && element.name === "txbx") count(tally, "text-box");
  if (element.namespace === WPG_NS && element.name === "wgp") count(tally, "grouped-shapes");
  if (element.namespace !== W_NS) return;

  switch (element.name) {
    case "p":
      count(tally, "paragraphs");
      return;
    case "tbl":
      count(tally, inCell ? "nested-tables" : "tables");
      return;
    case "tr":
      count(tally, "table-rows");
      return;
    case "tc":
      count(tally, "table-cells");
      return;
    case "gridSpan":
      count(tally, "cells-spanning-columns");
      return;
    case "vMerge":
      count(tally, "cells-spanning-rows");
      return;
    case "sectPr":
      count(tally, "sections");
      gathered.geometries.add(pageGeometrySignature(element));
      return;
    case "cols":
      if (Number(stated(element, "num") ?? 1) > 1) count(tally, "multiple-columns");
      return;
    case "numPr":
      count(tally, "numbered-paragraphs");
      return;
    case "tabs":
      count(tally, "stated-tab-stops");
      return;
    case "pBdr":
      count(tally, "paragraph-borders");
      return;
    case "tblBorders":
      count(tally, "table-borders");
      return;
    case "tcBorders":
      count(tally, "cell-borders");
      return;
    case "shd":
      count(tally, "shading");
      return;
    case "hyperlink":
      count(tally, "hyperlinks");
      return;
    case "fldSimple":
    case "instrText":
      count(tally, "fields");
      return;
    case "footnoteReference":
    case "endnoteReference":
      count(tally, "notes");
      return;
    case "pageBreakBefore":
      if (toggledOn(element)) count(tally, "page-breaks");
      return;
    case "br":
      countBreak(element, tally);
      return;
    case "spacing":
      countSpacing(element, tally);
      return;
    case "jc":
      count(tally, `aligned-${stated(element, "val") ?? "left"}`);
      return;
    case "rFonts":
      for (const which of ["ascii", "hAnsi", "cs", "eastAsia"]) {
        const face = stated(element, which);
        // The name is kept only long enough to know it is a different one. The
        // count leaves the file; the names never do.
        if (face !== undefined && face !== "") gathered.faces.add(face.toLowerCase());
      }
      return;
    case "drawing":
      count(tally, "drawings");
      return;
    case "pict":
    case "object":
      count(tally, "legacy-drawings");
      return;
    case "bidi":
    case "rtl":
      if (toggledOn(element)) count(tally, "right-to-left");
      return;
    default:
      return;
  }
}

type Gathered = {
  readonly faces: Set<string>;
  readonly geometries: Set<string>;
};

const WRAPS: ReadonlyMap<string, string> = new Map([
  ["wrapNone", "none"],
  ["wrapSquare", "square"],
  ["wrapTight", "tight"],
  ["wrapThrough", "through"],
  ["wrapTopAndBottom", "top-and-bottom"],
]);

function countBreak(element: XmlElement, tally: Tally): void {
  const type = stated(element, "type");
  if (type === "page") count(tally, "page-breaks");
  else if (type === "column") count(tally, "column-breaks");
  else count(tally, "line-breaks");
}

// Which line rule a paragraph asks for, which is one of the few things a census
// can see that decides how tall every line in it comes out.
function countSpacing(element: XmlElement, tally: Tally): void {
  const rule = stated(element, "lineRule");
  if (rule !== undefined) count(tally, `line-rule-${rule}`);
}

function walk(element: XmlElement, tally: Tally, inCell: boolean, gathered: Gathered): void {
  // The fallback beside a drawing is the copy Word itself ignores, and counting it
  // would report every shape twice.
  if (element.namespace === MC_NS && element.name === "Fallback") return;

  census(element, tally, inCell, gathered);
  const inside = inCell || (element.namespace === W_NS && element.name === "tc");
  for (const child of element.children) walk(child, tally, inside, gathered);
}

const HEADER_OR_FOOTER = /^word\/(header|footer)\d*\.xml$/;

export function censusOfPackage(pkg: DocxPackage, bytes: number, id: string): DocumentCensus {
  const tally: Tally = new Map();
  const gathered: Gathered = { faces: new Set(), geometries: new Set() };

  const parts = [
    MAIN_DOCUMENT_PART,
    ...[...pkg.parts.keys()].filter((part) => HEADER_OR_FOOTER.test(part)).sort(),
  ];

  for (const part of parts) {
    if (!pkg.parts.has(part)) continue;
    if (part !== MAIN_DOCUMENT_PART) count(tally, "headers-and-footers");
    walk(partXml(pkg, part), tally, false, gathered);
  }

  count(tally, "distinct-faces", gathered.faces.size);
  // Only the last section's geometry is read, so a document whose sections differ
  // in page size or margins is laid out on the wrong page above the last break.
  // A document that breaks a section only to change a header or a column count
  // asks nothing of that rule, and counting sections alone cannot tell the two
  // apart.
  count(tally, "distinct-page-geometries", gathered.geometries.size);
  if (gathered.geometries.size > 1) count(tally, "more-than-one-page-geometry");
  return { id, bytes, counts: Object.fromEntries([...tally.entries()].sort()) };
}

export function censusOf(bytes: Uint8Array): DocumentCensus {
  const id = createHash("sha256").update(bytes).digest("hex").slice(0, 12);
  try {
    return censusOfPackage(openDocx(bytes), bytes.length, id);
  } catch {
    // A package that will not open is still one of the corpus, and the sweep is
    // what says why.
    return { id, bytes: bytes.length, counts: {} };
  }
}

// What a document holds, without regard to how much of it: this is what a sample
// has to cover.
// A magnitude is not a feature: sampling on one would ask for a document per
// distinct count of it.
const MAGNITUDES = new Set(["distinct-faces", "distinct-page-geometries"]);

export const profileOf = (each: DocumentCensus): readonly string[] =>
  Object.keys(each.counts).filter((feature) => !MAGNITUDES.has(feature));

/**
 * The smallest set of documents that holds every feature the corpus has, each at
 * least `atLeast` times over.
 *
 * Greedy: take the document that adds the most still-wanted features, and keep
 * going until nothing is wanted. Greedy does not promise the smallest set there
 * is, but it is within a whisker of it and it is the difference between looking at
 * fifty documents and looking at a thousand.
 *
 * Documents are considered in a fixed order so that two runs over the same corpus
 * choose the same sample.
 */
export function coveringDocuments(
  censuses: readonly DocumentCensus[],
  atLeast = 3,
): readonly DocumentCensus[] {
  const wanted = new Map<string, number>();
  for (const each of censuses) {
    for (const feature of profileOf(each)) {
      wanted.set(feature, Math.min(atLeast, (wanted.get(feature) ?? 0) + 1));
    }
  }

  const left = [...censuses].sort((one, other) => one.id.localeCompare(other.id));
  const chosen: DocumentCensus[] = [];

  while (wanted.size > 0) {
    let best: DocumentCensus | null = null;
    let bestAdds = 0;

    for (const each of left) {
      const adds = profileOf(each).filter((feature) => wanted.has(feature)).length;
      if (adds > bestAdds) {
        best = each;
        bestAdds = adds;
      }
    }
    if (best === null) break;

    chosen.push(best);
    left.splice(left.indexOf(best), 1);
    for (const feature of profileOf(best)) {
      const still = (wanted.get(feature) ?? 0) - 1;
      if (still <= 0) wanted.delete(feature);
      else wanted.set(feature, still);
    }
  }

  return chosen;
}

export type FeatureTally = {
  readonly feature: string;
  readonly documents: number;
  readonly occurrences: number;
};

// Every feature, by how many documents hold it. This is the denominator the gap
// ranking wants: a gap met in two hundred documents is one thing where nine
// hundred have the feature and another where two hundred and ten do.
export function featuresIn(censuses: readonly DocumentCensus[]): readonly FeatureTally[] {
  const documents = new Map<string, number>();
  const occurrences = new Map<string, number>();

  for (const each of censuses) {
    for (const [feature, met] of Object.entries(each.counts)) {
      if (met === 0) continue;
      documents.set(feature, (documents.get(feature) ?? 0) + 1);
      occurrences.set(feature, (occurrences.get(feature) ?? 0) + met);
    }
  }

  return [...documents.entries()]
    .map(([feature, met]) => ({
      feature,
      documents: met,
      occurrences: occurrences.get(feature) ?? 0,
    }))
    .sort(
      (one, other) => other.documents - one.documents || one.feature.localeCompare(other.feature),
    );
}
