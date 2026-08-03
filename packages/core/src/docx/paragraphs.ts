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

export function readParagraphs(pkg: DocxPackage): readonly Paragraph[] {
  const root = partXml(pkg, MAIN_DOCUMENT_PART);
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
