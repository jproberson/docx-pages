import type { MetricsResolver } from "../layout/lines.js";
import { blockParagraphs, blocksIn } from "./blocks.js";
import { drawnAsStated } from "./borders.js";
import { readDrawingContent, type DrawingContent } from "./drawing.js";
import { MATH_NS, readEquation, runsAlone } from "./equations.js";
import { WP_NS } from "./inlines.js";
import { MAIN_DOCUMENT_PART, partXml, type DocxPackage } from "./package.js";
import { drawablePicture } from "./pictures.js";
import { defaultFooterPart, defaultHeaderPart, readRelationships } from "./relationships.js";
import { pageGeometrySignature, W_NS } from "./section.js";
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
//
// **A metafile is the one thing in the file that does not answer for itself.**
// Whether it draws is whether it plays, and playing it needs the faces this machine
// has, so a caller holding a resolver hands one over and a caller reading the
// package alone gets the answer the package can give (see `drawablePicture`).

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
  | "keep-lines-together"
  | "hidden-text"
  | "automatic-hyphenation"
  | "right-to-left"
  | "footnote"
  | "column-break"
  | "bar-tab-stop"
  | "page-background"
  | "equation"
  | "unknown-drawing"
  | "custom-geometry"
  | "undrawable-picture"
  | "approximated-border"
  | "alternate-first-or-even-page"
  | "substituted-face"
  | "character-from-another-face"
  | "missing-glyph";

const EFFECTS: Readonly<Record<UnhonouredKind, UnhonouredEffect>> = {
  // A page draws its own section's header and footer and hangs them where its own
  // section keeps them; how wide they are is still the document's answer, so a
  // section keeping a different left or right margin wraps its header at the wrong
  // width.
  "more-than-one-section": "moves-text",
  "text-columns": "moves-text",
  // Where a span and a merge put their text is built (see `planCells`), and what is
  // left is the lines round them: borders are settled by a cell's place in its row
  // rather than by the grid column it stands on, so a cell beside a span agrees with
  // the wrong neighbour, and the room half a line takes moves the text with it.
  "merged-cells": "moves-text",
  // What such a format says about a paragraph or a run is read; what it says about a
  // cell, a row or the table is not, so a first row shaded by its style comes out
  // unshaded and one lined by its style moves the text under it.
  "table-style-conditional-formatting": "moves-text",
  "keep-lines-together": "moves-text",
  // Hidden text is measured and drawn here as any other run, so it takes room
  // Word gives it none of.
  "hidden-text": "moves-text",
  "automatic-hyphenation": "moves-text",
  "right-to-left": "moves-text",
  // A note takes room at the foot of its page, which nothing here keeps for it.
  footnote: "moves-text",
  "column-break": "moves-text",
  // An equation this cannot read draws nothing at all and is measured as though the
  // paragraph holding it were empty, so the text of it is missing and everything below
  // it on the page has moved up. Found on 2026-08-12 by reading the top of the deformed
  // ranking: two documents of one template lose 16pt a page per equation and 6 of their
  // 14 pages between them cannot be shown, and **neither document stated a single gap**
  // until this was named. What is left under the name is the structures `readEquation`
  // refuses, the fraction above all.
  equation: "moves-text",
  "bar-tab-stop": "changes-paint",
  "page-background": "changes-paint",
  // A drawing that is neither a picture nor a shape, a chart being the one met so
  // far: its room is held and nothing is drawn in it.
  "unknown-drawing": "changes-paint",
  // A shape whose outline the file draws point by point in a way this cannot play,
  // which is an arc or a quadratic: the path is refused whole rather than drawn in
  // part, so its room is held and nothing is drawn in it. The box it fits in is not
  // a fallback, since a path that rules a page fits a box the size of the page.
  "custom-geometry": "changes-paint",
  // A picture nothing here draws: one in a format nothing decodes, WMF being what
  // Word writes beside the metafile this project plays, or a metafile the player
  // refuses. Its room is held and it is marked rather than drawn.
  //
  // A refused metafile is the quieter of the two and the reason the picture is put
  // to the player at all. A format nothing decodes is named by its own name, so the
  // report could always see it; a metafile that will not play is named by nothing
  // in the package, and a document holding one drew a blank rectangle and reported
  // no gap of any kind.
  "undrawable-picture": "changes-paint",
  "approximated-border": "changes-paint",
  // A document asking for one header on its even pages and another on its odd ones
  // draws the odd one everywhere, and a header of another height moves every line
  // under it.
  "alternate-first-or-even-page": "moves-text",
  // A face the document asked for that another one answered for: every line drawn
  // in it may break where Word did not break it.
  "substituted-face": "moves-text",
  // A character measured out of the face Word drew it from, so the room it takes
  // is Word's and nothing moves. What is drawn in that room is whatever the viewer
  // finds for a character its stated face cannot draw, which is not Word's glyph.
  "character-from-another-face": "changes-paint",
  // A character no face on hand maps at all, drawn as a missing-glyph box at the
  // box's own advance. Word would have reached for a face of its own before
  // drawing one, so the width here is not Word's width.
  "missing-glyph": "moves-text",
};

// Whether an element that is a toggle is on. Word writes the toggle bare to turn
// it on, so an element with no value at all counts.
function toggled(element: XmlElement): boolean {
  const value = attribute(element, W_NS, "val");
  return value === undefined || (value !== "0" && value !== "false" && value !== "off");
}

// Which part a drawing's picture is held in, and what that part holds. Reading a
// part the package does not carry is a broken package rather than a feature passed
// over, so nothing is said about one. **The bytes are wanted as well as the name**,
// since whether a metafile can be drawn is what is recorded inside it.
type PartResolver = (relationshipId: string) => {
  readonly part: string;
  readonly bytes: Uint8Array | undefined;
} | null;

// Whether anything in a drawing, at any depth of a group, states an outline this
// project cannot play. **A path it can play is no longer a gap**: what is left under
// this name is a path holding an arc or a quadratic, which the reader refuses whole
// rather than drawing part of, and neither appears in any of the 718 documents.
function drawsACustomPath(content: DrawingContent): boolean {
  if (content.kind === "group")
    return content.children.some((each) => drawsACustomPath(each.content));
  return (
    content.kind !== "unknown" && content.paint.geometry === "custom" && content.paint.path === null
  );
}

// What an element says about itself, where what it says is something this project
// passes over. A name alone is not enough: `w:caps` is written both ways round,
// and a document that turns a feature off is asking for what it already gets.
function unhonouredBy(
  element: XmlElement,
  parent: XmlElement | null,
  paragraph: XmlElement | null,
  resolvePart: PartResolver,
  metricsFor: MetricsResolver | undefined,
): UnhonouredKind | null {
  // A drawing answers for itself, by the same reader the layout uses: whatever
  // that cannot make a picture or a shape of is drawn nowhere, and a picture is
  // drawn only where something here decodes the format it is held in.
  if (element.namespace === WP_NS && (element.name === "anchor" || element.name === "inline")) {
    const content = readDrawingContent(element);
    if (content.kind === "unknown") return "unknown-drawing";
    if (drawsACustomPath(content)) return "custom-geometry";
    if (content.kind !== "picture") return null;
    const held = resolvePart(content.relationshipId);
    return held === null || drawablePicture(held.part, held.bytes, metricsFor)
      ? null
      : "undrawable-picture";
  }
  // An equation answers for itself, by the same reader the layout uses. **What is named
  // here is what is not drawn, which is more than what is not read**: the shape of a
  // fraction and of a delimiter is read, and how tall Word sets one is unmeasured, so
  // until that is answered an equation holding either draws nothing and is named. An
  // equation of runs alone needs no setting and is drawn where the paragraph's own runs
  // are.
  //
  // It is `m:oMath` wherever it stands, and `m:oMathPara` around it where it stands alone
  // in its paragraph, so the inner one answers for both and a paragraph of two equations
  // says so twice.
  if (element.namespace === MATH_NS) {
    if (element.name !== "oMath") return null;
    return runsAlone(readEquation(element)) === null ? "equation" : null;
  }
  if (element.namespace !== W_NS) return null;

  switch (element.name) {
    case "gridSpan":
    case "vMerge":
      return "merged-cells";
    case "tblStylePr":
      return conditionalFormattingUnread(element);
    case "keepLines":
      return toggled(element) ? "keep-lines-together" : null;
    // Kerning was built on 2026-08-13, in the pairs the face itself states, and
    // capitals the same day: every letter at the run's own size under `w:caps`, or a
    // small capital at four fifths of it rounded to the nearest half point.
    case "kern":
      return null;
    case "caps":
    case "smallCaps":
      return null;
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
    // A column break is honoured where it stands alone in its paragraph or opens
    // one, which is where every one of the 25 in the corpus stands. One with text of
    // its own paragraph in front of it is a place inside a block, and the division
    // into columns is made between them.
    case "br":
      return attribute(element, W_NS, "type") === "column" && drawnBefore(element, paragraph)
        ? "column-break"
        : null;
    case "tab":
      return attribute(element, W_NS, "val") === "bar" ? "bar-tab-stop" : null;
    // Built on 2026-08-13: the run's own advance across, the line's box down, in the
    // colour the name stands for. A colour outside the sixteen is nothing Word
    // paints either, so no name is left to raise.
    case "highlight":
      return null;
    // Every side of every border is written the same way, so one case answers for
    // a table's, a cell's and a paragraph's alike.
    case "top":
    case "bottom":
    case "left":
    case "right":
    case "insideH":
    case "insideV":
    case "start":
    case "end":
      return borderPattern(element, parent);
    case "background":
      return "page-background";
    // A first-page header and footer were built on 2026-08-10: the page a section
    // opens draws that section's own first-page pair where it states w:titlePg, so
    // stating it stands in for nothing. What is left of this gap is the even-page
    // pair, which is read and never chosen.
    case "titlePg":
      return null;
    case "evenAndOddHeaders":
      return toggled(element) ? "alternate-first-or-even-page" : null;
    // Columns were built on 2026-08-08, so a section running its text in more than
    // one of them no longer stands in for anything on its own.
    case "cols":
      return null;
    default:
      return null;
  }
}

// Whether the paragraph holding a break has drawn anything before it. Word's own
// runs are read rather than the paragraph's, since a break in the second run of a
// paragraph is what this is looking for.
function drawnBefore(element: XmlElement, paragraph: XmlElement | null): boolean {
  if (paragraph === null) return false;
  let seen = false;
  let reached = false;
  const walk = (node: XmlElement): void => {
    for (const child of node.children) {
      if (reached) return;
      if (child === element) {
        reached = true;
        return;
      }
      if (child.namespace === W_NS && child.name === "t" && child.text !== "") seen = true;
      walk(child);
    }
  };
  walk(paragraph);
  return seen;
}

// **What a table style says about one place in the table is read for the paragraphs
// and the runs standing there and for nothing else.** A `w:tblStylePr` holding a
// `w:pPr` or a `w:rPr` alone is honoured (see `CONDITIONAL_ORDER`); one that also
// states a `w:tcPr`, a `w:trPr` or a `w:tblPr` is asking for shading, a row height
// or a border this project still settles from the table's own properties.
const READ_IN_A_CONDITIONAL_FORMAT = new Set(["pPr", "rPr"]);

function conditionalFormattingUnread(element: XmlElement): UnhonouredKind | null {
  for (const child of element.children)
    if (child.namespace === W_NS && !READ_IN_A_CONDITIONAL_FORMAT.has(child.name))
      return "table-style-conditional-formatting";
  return null;
}

// A border names its pattern in `w:val`, and only inside one of the elements that
// hold borders: `w:top` means something else entirely in a cell's margins.
function borderPattern(element: XmlElement, parent: XmlElement | null): UnhonouredKind | null {
  if (
    parent === null ||
    (parent.name !== "pBdr" && parent.name !== "tblBorders" && parent.name !== "tcBorders")
  ) {
    return null;
  }
  const value = attribute(element, W_NS, "val");
  return value === undefined || drawnAsStated(value) ? null : "approximated-border";
}

type Found = { readonly kind: UnhonouredKind; readonly place: UnhonouredPlace };

/**
 * What the document asks for and does not get.
 *
 * The faces are the one thing this cannot read out of the package, and it needs
 * them for one answer alone: a metafile draws what it plays, and it plays against
 * the metrics of the faces it selects. A caller laying the document out holds a
 * resolver already and hands it over. One asking what a document holds before any
 * face is to hand may leave it out, and its metafiles are taken on trust while
 * everything else in the list still answers.
 */
export function readUnhonoured(
  pkg: DocxPackage,
  metricsFor?: MetricsResolver,
): readonly Unhonoured[] {
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
    const relationships = readRelationships(pkg, part);
    walk(root, null, {
      part,
      paragraphs,
      found,
      metricsFor,
      resolvePart: (relationshipId) => {
        const held = relationships.get(relationshipId)?.part;
        return held === undefined ? null : { part: held, bytes: pkg.parts.get(held) };
      },
    });
    countSections(root, part, found, parts.length > 1);
  }

  // A style's conditional formats and the settings' own switches belong to no part
  // of the flow, so they are read where they are written. Only the styles the flow
  // reaches are read: see `stylesTheFlowReaches`.
  if (pkg.parts.has(STYLES_PART)) {
    const styles = partXml(pkg, STYLES_PART);
    const reached = stylesTheFlowReaches(pkg, styles, parts);
    for (const child of styles.children) {
      if (child.namespace === W_NS && child.name === "style" && !reached.has(child)) continue;
      walk(child, null, {
        part: STYLES_PART,
        paragraphs: new Map(),
        found,
        metricsFor,
        resolvePart: () => null,
      });
    }
  }

  if (pkg.parts.has(SETTINGS_PART)) {
    walk(partXml(pkg, SETTINGS_PART), null, {
      part: SETTINGS_PART,
      paragraphs: new Map(),
      found,
      metricsFor,
      resolvePart: () => null,
    });
  }

  return gathered(found);
}

// The styles some paragraph, run or table in the flow actually resolves to, with
// everything each of them is based on.
//
// **A style nothing is written in is not something the document asked for.** Reading
// every `w:style` in the part said 509 of the 718 corpus documents wanted kerning
// where 44 state `w:kern` in their own flow and 38 more on the defaults every
// paragraph inherits; the other 481 carry it on a named style, and the ranking that
// number led was pointing at nothing. `keep-with-next` was rank 1 on the same
// mistake, at 707 documents against the 44 that have a paragraph resolving to it.
//
// `w:docDefaults` is not a style and is always read: what it states, every paragraph
// in the document gets.
function stylesTheFlowReaches(
  pkg: DocxPackage,
  styles: XmlElement,
  parts: readonly string[],
): ReadonlySet<XmlElement> {
  const byId = new Map<string, XmlElement>();
  const wanted = new Set<string>();

  for (const style of styles.children) {
    if (style.namespace !== W_NS || style.name !== "style") continue;
    const id = attribute(style, W_NS, "styleId");
    if (id === undefined) continue;
    byId.set(id, style);
    // A paragraph naming no style of its own is written in the default one, and so
    // is a run and a table.
    if (attribute(style, W_NS, "default") === "1") wanted.add(id);
  }

  const NAMES = new Set(["pStyle", "rStyle", "tblStyle"]);
  const gather = (node: XmlElement): void => {
    for (const child of node.children) {
      if (child.namespace === W_NS && NAMES.has(child.name)) {
        const id = attribute(child, W_NS, "val");
        if (id !== undefined) wanted.add(id);
      }
      gather(child);
    }
  };
  for (const part of parts) gather(partXml(pkg, part));

  const reached = new Set<XmlElement>();
  for (const id of wanted) {
    let at: string | undefined = id;
    // A cycle in `w:basedOn` is a broken package rather than a feature passed over,
    // and walking one is what would hang here.
    const walked = new Set<string>();
    while (at !== undefined && !walked.has(at)) {
      walked.add(at);
      const style = byId.get(at);
      if (style === undefined) break;
      reached.add(style);
      const basedOn = basedOnOf(style);
      at = basedOn === null ? undefined : (attribute(basedOn, W_NS, "val") ?? undefined);
    }
  }

  return reached;
}

function basedOnOf(style: XmlElement): XmlElement | null {
  for (const child of style.children)
    if (child.namespace === W_NS && child.name === "basedOn") return child;
  return null;
}

const STYLES_PART = "word/styles.xml";

const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

// What one part of the package answers with: where a paragraph met in it is
// numbered, where a drawing's picture is held, which faces are to hand to play a
// metafile with, and what has been met so far.
type Reading = {
  readonly part: string;
  readonly paragraphs: ReadonlyMap<XmlElement, number>;
  readonly resolvePart: PartResolver;
  readonly metricsFor: MetricsResolver | undefined;
  readonly found: Found[];
};

// A text box's own content is laid out from its frame rather than from the part's
// flow, and its paragraphs are numbered inside it, so what is met in one answers
// for the paragraph that anchors it.
function walk(
  node: XmlElement,
  paragraphIndex: number | null,
  reading: Reading,
  paragraph: XmlElement | null = null,
): void {
  for (const child of node.children) {
    const standing = reading.paragraphs.has(child) ? child : paragraph;
    const kind = unhonouredBy(child, node, standing, reading.resolvePart, reading.metricsFor);
    const index = reading.paragraphs.get(child) ?? paragraphIndex;
    if (kind !== null) {
      reading.found.push({ kind, place: { part: reading.part, paragraphIndex: index } });
    }
    // A text box's own content is laid out and counts; the fallback beside a
    // drawing is the copy Word itself ignores, and is passed over here too.
    if (child.namespace === MC_NS && child.name === "Fallback") continue;
    walk(child, index, reading, standing);
  }
}

// Every page is now broken and drawn against the geometry of the section whose
// text opened it, so a second page size or a second set of margins costs the body
// nothing, and as of 2026-08-10 a page hangs its header and its footer from the
// room its own section keeps for them as well.
//
// **What is still read from the last section alone is how wide they are.** A
// header is measured across the document's own text frame, so a section keeping a
// different left or right margin draws its header at the wrong left and wraps it
// at the wrong width. Nothing has measured what that is worth.
//
// So what is named here is a second *page* with a header or a footer to put in the
// wrong place. A break that changes only a column count leaves the geometry alone
// and is named elsewhere, and a document that draws neither header nor footer has
// nothing left standing in for it. The count is deliberately the loose one: a
// second page size is named whether or not its margins are the ones that still
// matter, which leaves the row wider than the fault.
function countSections(
  root: XmlElement,
  part: string,
  found: Found[],
  drawsAHeaderOrFooter: boolean,
): void {
  if (!drawsAHeaderOrFooter) return;
  const pages = new Set(allNamed(root, "sectPr").map(pageGeometrySignature));
  for (let at = 1; at < pages.size; at += 1) {
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
// and this puts them in the list the document itself reported, in the place the
// order of kinds gives them.
//
// One place a face, naming the document's own part: the layout knows which faces
// it stood in for but not which paragraph asked for each of them. The same is true
// of a character drawn from another face.
export function withSubstitutedFaces(
  unhonoured: readonly Unhonoured[],
  substitutions: readonly { readonly requested: { readonly name: string } }[],
): readonly Unhonoured[] {
  return withEntryFor("substituted-face", unhonoured, substitutions.length);
}

// The other of the two, and the same argument: which characters a face had no
// glyph for is only known once the layout has asked it for them. One place a
// character, so a document drawing two of them says so twice.
export function withFallbackCharacters(
  unhonoured: readonly Unhonoured[],
  characters: readonly { readonly codePoint: number }[],
): readonly Unhonoured[] {
  return withEntryFor("character-from-another-face", unhonoured, characters.length);
}

// The last of the three, from the same resolver: a character nothing on hand
// could draw, stood in for by the missing-glyph box.
export function withMissingGlyphs(
  unhonoured: readonly Unhonoured[],
  characters: readonly { readonly codePoint: number }[],
): readonly Unhonoured[] {
  return withEntryFor("missing-glyph", unhonoured, characters.length);
}

function withEntryFor(
  kind: UnhonouredKind,
  unhonoured: readonly Unhonoured[],
  met: number,
): readonly Unhonoured[] {
  if (met === 0) return unhonoured;
  return [
    ...unhonoured,
    {
      kind,
      effect: EFFECTS[kind],
      places: Array.from({ length: met }, () => ({
        part: MAIN_DOCUMENT_PART,
        paragraphIndex: null,
      })),
    },
  ].sort((one, other) => one.kind.localeCompare(other.kind));
}
