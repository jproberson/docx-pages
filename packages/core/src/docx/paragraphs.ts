import { MAIN_DOCUMENT_PART, partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import { descendantsNamed, type XmlElement } from "./xml.js";

export const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

export type Paragraph = {
  readonly index: number;
  readonly element: XmlElement;
};

const isTextBoxContent = (node: XmlElement): boolean =>
  node.namespace === W_NS && node.name === "txbxContent";

const isFallback = (node: XmlElement): boolean =>
  node.namespace === MC_NS && node.name === "Fallback";

function collect(node: XmlElement, into: XmlElement[]): void {
  for (const child of node.children) {
    if (isTextBoxContent(child) || isFallback(child)) continue;
    if (child.namespace === W_NS && child.name === "p") into.push(child);
    collect(child, into);
  }
}

export function readParagraphs(
  pkg: DocxPackage,
  part: string = MAIN_DOCUMENT_PART,
): readonly Paragraph[] {
  const root = partXml(pkg, part);
  const elements: XmlElement[] = [];
  collect(root, elements);
  return elements.map((element, index) => ({ index, element }));
}

export function paragraphText(paragraph: Paragraph): string {
  const kept: XmlElement[] = [];
  collectRuns(paragraph.element, kept);
  return kept.map((node) => node.text).join("");
}

function collectRuns(node: XmlElement, into: XmlElement[]): void {
  for (const child of node.children) {
    if (isTextBoxContent(child) || isFallback(child)) continue;
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
    if (isTextBoxContent(child) || isFallback(child)) continue;
    if (child.namespace === W_NS && child.name === "p") continue;
    if (child.namespace === W_NS && child.name === name) into.push(child);
    collectNamed(child, name, into);
  }
}

const WP_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

// A run holding only a floating anchor places nothing on the line, so it does not
// contribute to the line's height. Word lays floats out of flow.
function placesContentInLine(run: XmlElement): boolean {
  let found = false;
  const visit = (node: XmlElement): void => {
    if (found) return;
    for (const child of node.children) {
      if (isTextBoxContent(child) || isFallback(child)) continue;
      const isText = child.namespace === W_NS && (child.name === "t" || child.name === "tab");
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
      if (isTextBoxContent(child) || isFallback(child)) continue;
      if (child.namespace === namespace && child.name === name) found.push(child);
      visit(child);
    }
  };
  visit(paragraph.element);
  return found;
}
