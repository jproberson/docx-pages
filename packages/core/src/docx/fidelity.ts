import { blockParagraphs, blocksIn } from "./blocks.js";
import { readDrawingContent } from "./drawing.js";
import { WP_NS } from "./inlines.js";
import { MAIN_DOCUMENT_PART, partXml, type DocxPackage } from "./package.js";
import { defaultFooterPart, defaultHeaderPart } from "./relationships.js";
import { W_NS } from "./section.js";
import { SETTINGS_PART } from "./settings.js";
import { attribute, type XmlElement } from "./xml.js";

// What this project met in a document and did not honour.
//
// A renderer that quietly passes over what it does not understand is the failure
// this whole project exists to avoid: a page that looks plausible and is not the
// page Word draws is worse than one that says where it is not to be trusted. So
// everything known to be unread is named here, and the suites hold each document
// to the list it is expected to produce.
//
// This reads the document rather than watching the layout, so it says what is in
// the file, not what the layout would have done with it. The two are kept together
// by the suites: a feature that stops being ignored has to leave this list.

// How much of the page an entry puts in doubt. Text that moved is a page that
// cannot be trusted anywhere below it; paint that changed is wrong only where it
// stands.
export type UnhonouredEffect = "moves-text" | "changes-paint";

// Where a feature was met: which part of the package, and which of that part's
// paragraphs, numbered as this project numbers them. Something a paragraph does
// not hold, a section's own columns among them, answers with no paragraph.
export type UnhonouredPlace = {
  readonly part: string;
  readonly paragraphIndex: number | null;
};

export type Unhonoured = {
  readonly kind: UnhonouredKind;
  readonly effect: UnhonouredEffect;
  // Every place it was met, in the order the parts were read.
  readonly places: readonly UnhonouredPlace[];
};

export type UnhonouredKind =
  | "more-than-one-section"
  | "text-columns"
  | "merged-cells"
  | "table-style-conditional-formatting"
  | "unsplittable-row"
  | "keep-with-next"
  | "keep-lines-together"
  | "character-kerning"
  | "character-spacing"
  | "capitals"
  | "raised-or-lowered-text"
  | "hidden-text"
  | "automatic-hyphenation"
  | "right-to-left"
  | "footnote"
  | "column-break"
  | "bar-tab-stop"
  | "highlighting"
  | "page-background"
  | "unknown-drawing"
  | "alternate-first-or-even-page"
  | "substituted-face";

const EFFECTS: Readonly<Record<UnhonouredKind, UnhonouredEffect>> = {
  // Only the last section's geometry is read, so a document that changes page
  // size, margins or columns part way through lays the rest out on the wrong page.
  "more-than-one-section": "moves-text",
  "text-columns": "moves-text",
  // A cell spanning its neighbours is laid out at its own width, and the cells
  // beside it resolve their borders against the wrong neighbour.
  "merged-cells": "moves-text",
  "table-style-conditional-formatting": "changes-paint",
  "unsplittable-row": "moves-text",
  "keep-with-next": "moves-text",
  "keep-lines-together": "moves-text",
  "character-kerning": "moves-text",
  "character-spacing": "moves-text",
  capitals: "moves-text",
  "raised-or-lowered-text": "moves-text",
  // Hidden text is measured and drawn here as any other run, so it takes room
  // Word gives it none of.
  "hidden-text": "moves-text",
  "automatic-hyphenation": "moves-text",
  "right-to-left": "moves-text",
  // A note takes room at the foot of its page, which nothing here keeps for it.
  footnote: "moves-text",
  "column-break": "moves-text",
  "bar-tab-stop": "changes-paint",
  highlighting: "changes-paint",
  "page-background": "changes-paint",
  // A drawing that is neither a picture nor a shape, a chart being the one met so
  // far: its room is held and nothing is drawn in it.
  "unknown-drawing": "changes-paint",
  // Only one header and one footer are drawn, on every page alike, so a document
  // that asks for another on its first page or on its even ones draws the wrong
  // one there.
  "alternate-first-or-even-page": "moves-text",
  // A face the document asked for that another one answered for: every line drawn
  // in it may break where Word did not break it.
  "substituted-face": "moves-text",
};

// Whether an element that is a toggle is on. Word writes the toggle bare to turn
// it on, so an element with no value at all counts.
function toggled(element: XmlElement): boolean {
  const value = attribute(element, W_NS, "val");
  return value === undefined || (value !== "0" && value !== "false" && value !== "off");
}

const numbered = (element: XmlElement): number => {
  const value = Number(attribute(element, W_NS, "val"));
  return Number.isFinite(value) ? value : 0;
};

// What an element says about itself, where what it says is something this project
// passes over. A name alone is not enough: `w:caps` is written both ways round,
// and a document that turns a feature off is asking for what it already gets.
function unhonouredBy(element: XmlElement, parent: XmlElement | null): UnhonouredKind | null {
  // A drawing answers for itself, by the same reader the layout uses: whatever
  // that cannot make a picture or a shape of is drawn nowhere.
  if (element.namespace === WP_NS && (element.name === "anchor" || element.name === "inline")) {
    return readDrawingContent(element).kind === "unknown" ? "unknown-drawing" : null;
  }
  if (element.namespace !== W_NS) return null;

  switch (element.name) {
    case "gridSpan":
    case "vMerge":
      return "merged-cells";
    case "tblStylePr":
      return "table-style-conditional-formatting";
    case "cantSplit":
      return toggled(element) ? "unsplittable-row" : null;
    case "keepNext":
      return toggled(element) ? "keep-with-next" : null;
    case "keepLines":
      return toggled(element) ? "keep-lines-together" : null;
    case "kern":
      return numbered(element) > 0 ? "character-kerning" : null;
    // Inside run properties `w:spacing` is the room between letters; a paragraph's
    // own `w:spacing`, which is read, states no `w:val` at all.
    case "spacing":
      return parent?.name === "rPr" && numbered(element) !== 0 ? "character-spacing" : null;
    case "caps":
    case "smallCaps":
      return toggled(element) ? "capitals" : null;
    case "position":
      return numbered(element) !== 0 ? "raised-or-lowered-text" : null;
    case "vanish":
      return toggled(element) ? "hidden-text" : null;
    case "autoHyphenation":
      return toggled(element) ? "automatic-hyphenation" : null;
    case "rtl":
    case "bidi":
      return toggled(element) ? "right-to-left" : null;
    case "footnoteReference":
    case "endnoteReference":
      return "footnote";
    case "br":
      return attribute(element, W_NS, "type") === "column" ? "column-break" : null;
    case "tab":
      return attribute(element, W_NS, "val") === "bar" ? "bar-tab-stop" : null;
    case "highlight":
      return "highlighting";
    case "background":
      return "page-background";
    case "titlePg":
    case "evenAndOddHeaders":
      return toggled(element) ? "alternate-first-or-even-page" : null;
    case "cols":
      return numbered(element) > 1 || Number(attribute(element, W_NS, "num") ?? 1) > 1
        ? "text-columns"
        : null;
    default:
      return null;
  }
}

type Found = { readonly kind: UnhonouredKind; readonly place: UnhonouredPlace };

export function readUnhonoured(pkg: DocxPackage): readonly Unhonoured[] {
  const found: Found[] = [];
  const parts = [MAIN_DOCUMENT_PART, defaultHeaderPart(pkg), defaultFooterPart(pkg)].filter(
    (part): part is string => part !== null && pkg.parts.has(part),
  );

  for (const part of parts) {
    const root = partXml(pkg, part);
    const paragraphs = new Map(
      // Numbered off the tree being walked, so that a paragraph met here is the
      // same object the layout will lay out rather than a second reading of it.
      blockParagraphs(blocksIn(root)).map((each) => [each.element, each.index]),
    );
    walk(root, part, null, paragraphs, found);
    countSections(root, part, found);
  }

  // A style's conditional formats and the settings' own switches belong to no part
  // of the flow, so they are read where they are written.
  for (const part of [STYLES_PART, SETTINGS_PART]) {
    if (pkg.parts.has(part)) walk(partXml(pkg, part), part, null, new Map(), found);
  }

  return gathered(found);
}

const STYLES_PART = "word/styles.xml";

const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

// A text box's own content is laid out from its frame rather than from the part's
// flow, and its paragraphs are numbered inside it, so what is met in one answers
// for the paragraph that anchors it.
function walk(
  node: XmlElement,
  part: string,
  paragraphIndex: number | null,
  paragraphs: ReadonlyMap<XmlElement, number>,
  found: Found[],
): void {
  for (const child of node.children) {
    const kind = unhonouredBy(child, node);
    const index = paragraphs.get(child) ?? paragraphIndex;
    if (kind !== null) found.push({ kind, place: { part, paragraphIndex: index } });
    // A text box's own content is laid out and counts; the fallback beside a
    // drawing is the copy Word itself ignores, and is passed over here too.
    if (child.namespace === MC_NS && child.name === "Fallback") continue;
    walk(child, part, index, paragraphs, found);
  }
}

// Word lays a document out under the section its text ends in, and this project
// reads the last one alone. One section is the ordinary document; a second is a
// page whose geometry nothing here has answered for.
function countSections(root: XmlElement, part: string, found: Found[]): void {
  const sections = allNamed(root, "sectPr");
  for (let at = 1; at < sections.length; at += 1) {
    found.push({ kind: "more-than-one-section", place: { part, paragraphIndex: null } });
  }
}

function allNamed(node: XmlElement, name: string): readonly XmlElement[] {
  const found: XmlElement[] = [];
  const visit = (each: XmlElement): void => {
    if (each.namespace === W_NS && each.name === name) found.push(each);
    for (const child of each.children) visit(child);
  };
  visit(node);
  return found;
}

// One entry a kind, holding every place it was met, in the order the kinds are
// named so that a report reads the same way twice.
function gathered(found: readonly Found[]): readonly Unhonoured[] {
  const places = new Map<UnhonouredKind, UnhonouredPlace[]>();
  for (const each of found) {
    const already = places.get(each.kind);
    if (already === undefined) places.set(each.kind, [each.place]);
    else already.push(each.place);
  }

  return [...places.entries()]
    .map(([kind, met]) => ({ kind, effect: EFFECTS[kind], places: met }))
    .sort((one, other) => one.kind.localeCompare(other.kind));
}

// A face stood in for is not in the document's own words, so it is not read out of
// the package: the layout is what finds one. `substitutingMetrics` collects them,
// and this puts them in the same list as everything else.
export const unhonouredFaces = (
  substitutions: readonly { readonly requested: { readonly name: string } }[],
): readonly Unhonoured[] =>
  substitutions.length === 0
    ? []
    : [
        {
          kind: "substituted-face" as const,
          effect: EFFECTS["substituted-face"],
          places: substitutions.map(() => ({ part: MAIN_DOCUMENT_PART, paragraphIndex: null })),
        },
      ];
