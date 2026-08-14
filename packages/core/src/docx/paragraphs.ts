import { isDetachedContent, type Paragraph } from "./blocks.js";
import { MATH_NS, readEquation, runElementsOf } from "./equations.js";
import { W_NS } from "./section.js";
import { holdsALegacyPicture, inlinePictureOf } from "./vml.js";
import { descendantsNamed, type XmlElement } from "./xml.js";

export function paragraphText(paragraph: Paragraph): string {
  const kept: XmlElement[] = [];
  collectRuns(paragraph.element, kept);
  return kept.map((node) => node.text).join("");
}

function collectRuns(node: XmlElement, into: XmlElement[]): void {
  for (const child of node.children) {
    if (isDetachedContent(child)) continue;
    if (child.namespace === W_NS && child.name === "p" && child !== node) continue;
    if (child.namespace === W_NS && child.name === "t") {
      into.push(child);
      continue;
    }
    collectRuns(child, into);
  }
}

export const paragraphDescendants = (
  paragraph: Paragraph,
  namespace: string,
  name: string,
): readonly XmlElement[] => descendantsNamed(paragraph.element, namespace, name);

// An equation stands where the paragraph holds it, so its runs are gathered by the
// same walk rather than appended after, and every run it holds comes out however deep
// it stands. **Descending into a fraction would put its halves on the line side by
// side, which is why it was refused until there was geometry for one**; what makes it
// safe now is that `readRuns` gathers an equation's runs back into one piece before
// anything reaches a line, so a half is never a thing a line can see. The runs are
// wanted here all the same, because a run is marked by the cascade only where the
// paragraph hands it out, and a half cannot be measured without its mark. **A run
// carrying nothing but a break is one of them**, since the line it ends is measured
// from that run.
function collectNamed(node: XmlElement, name: string, into: XmlElement[]): void {
  for (const child of node.children) {
    if (isDetachedContent(child)) continue;
    if (child.namespace === W_NS && child.name === "p") continue;
    if (child.namespace === MATH_NS && child.name === "oMath") {
      if (name === "r") {
        for (const run of runElementsOf(readEquation(child))) into.push(run);
      }
      continue;
    }
    if (child.namespace === W_NS && child.name === name) into.push(child);
    collectNamed(child, name, into);
  }
}

const WP_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

// A run holding only a floating anchor places nothing on the line, so it does not
// contribute to the line's height. Word lays floats out of flow. A run holding
// only a break is not that: it ends the line it sits on, and dropping it runs the
// text either side of it together.
const LINE_CONTENT = new Set(["t", "tab", "br"]);

function holdsLineContent(run: XmlElement, mustDraw: boolean): boolean {
  let found = false;
  const visit = (node: XmlElement): void => {
    if (found) return;
    for (const child of node.children) {
      if (isDetachedContent(child)) continue;
      const isText =
        (child.namespace === W_NS && LINE_CONTENT.has(child.name)) ||
        (child.namespace === MATH_NS && child.name === "t");
      const isInline = child.namespace === WP_DRAWING_NS && child.name === "inline";
      const isLegacy =
        holdsALegacyPicture(child.namespace, child.name) && inlinePictureOf(child) !== null;
      if (isText && mustDraw && child.name === "t" && child.text === "") {
        visit(child);
        continue;
      }
      if (isText || isInline || isLegacy) {
        found = true;
        return;
      }
      visit(child);
    }
  };
  visit(run);
  return found;
}

const placesContentInLine = (run: XmlElement): boolean => holdsLineContent(run, false);

// **Whether the run draws anything, which is not whether it stands on the line.** A
// `w:t` holding nothing takes the line's own height with it and puts no ink on the
// page, and the two questions part company over a display equation: Word centred one
// beside an empty run and laid the same one in the flow beside a single space,
// measured on 2026-08-13 over the authored probe.
export const drawsInLine = (run: XmlElement): boolean => holdsLineContent(run, true);

export function paragraphRuns(paragraph: Paragraph): readonly XmlElement[] {
  const found: XmlElement[] = [];
  collectNamed(paragraph.element, "r", found);
  return found.filter(placesContentInLine);
}

export type ElementName = { readonly namespace: string; readonly name: string };

export function paragraphOwnDrawings(
  paragraph: Paragraph,
  wanted: readonly ElementName[],
): readonly XmlElement[] {
  const found: XmlElement[] = [];
  const visit = (node: XmlElement): void => {
    for (const child of node.children) {
      if (isDetachedContent(child)) continue;
      if (wanted.some((each) => child.namespace === each.namespace && child.name === each.name)) {
        found.push(child);
      }
      visit(child);
    }
  };
  visit(paragraph.element);
  return found;
}
