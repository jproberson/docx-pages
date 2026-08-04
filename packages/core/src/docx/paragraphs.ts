import { isDetachedContent, type Paragraph } from "./blocks.js";
import { W_NS } from "./section.js";
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

function collectNamed(node: XmlElement, name: string, into: XmlElement[]): void {
  for (const child of node.children) {
    if (isDetachedContent(child)) continue;
    if (child.namespace === W_NS && child.name === "p") continue;
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

function placesContentInLine(run: XmlElement): boolean {
  let found = false;
  const visit = (node: XmlElement): void => {
    if (found) return;
    for (const child of node.children) {
      if (isDetachedContent(child)) continue;
      const isText = child.namespace === W_NS && LINE_CONTENT.has(child.name);
      const isInline = child.namespace === WP_DRAWING_NS && child.name === "inline";
      if (isText || isInline) {
        found = true;
        return;
      }
      visit(child);
    }
  };
  visit(run);
  return found;
}

export function paragraphRuns(paragraph: Paragraph): readonly XmlElement[] {
  const found: XmlElement[] = [];
  collectNamed(paragraph.element, "r", found);
  return found.filter(placesContentInLine);
}

export function paragraphOwnDrawings(
  paragraph: Paragraph,
  namespace: string,
  name: string,
): readonly XmlElement[] {
  const found: XmlElement[] = [];
  const visit = (node: XmlElement): void => {
    for (const child of node.children) {
      if (isDetachedContent(child)) continue;
      if (child.namespace === namespace && child.name === name) found.push(child);
      visit(child);
    }
  };
  visit(paragraph.element);
  return found;
}
