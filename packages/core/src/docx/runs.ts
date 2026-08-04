import { isDetachedContent, type Paragraph } from "./blocks.js";
import { WP_NS } from "./inlines.js";
import { W_NS } from "./section.js";
import { resolveRuns, type ParagraphMark, type StyleTable } from "./styles.js";
import { attribute, firstNamed, type XmlElement } from "./xml.js";

export type RunPiece =
  | { readonly kind: "text"; readonly text: string }
  | { readonly kind: "tab" }
  | { readonly kind: "break" }
  | { readonly kind: "drawing"; readonly widthEmu: number; readonly heightEmu: number };

export type TextRun = {
  readonly mark: ParagraphMark;
  readonly pieces: readonly RunPiece[];
};

// The whitespace xml leaves insignificant, which is the only kind the edges of a
// w:t lose. A no-break space is a character the text is made of rather than
// whitespace around it, so it stays whether or not the run asks for it.
const INSIGNIFICANT = /^[ \t\r\n]+|[ \t\r\n]+$/g;

// Whitespace at the edges of a w:t is insignificant unless the run asks for it,
// which is why Word writes xml:space wherever a space has to survive.
function textOf(element: XmlElement): string {
  const preserved = attribute(element, "", "space") === "preserve";
  return preserved ? element.text : element.text.replace(INSIGNIFICANT, "");
}

function extentOf(inline: XmlElement): RunPiece {
  const extent = firstNamed(inline, WP_NS, "extent");
  const size = (name: string): number => {
    const raw = extent === null ? undefined : attribute(extent, "", name);
    const value = raw === undefined ? Number.NaN : Number(raw);
    return Number.isFinite(value) ? value : 0;
  };
  return { kind: "drawing", widthEmu: size("cx"), heightEmu: size("cy") };
}

function collectPieces(node: XmlElement, into: RunPiece[]): void {
  for (const child of node.children) {
    if (isDetachedContent(child)) continue;

    if (child.namespace === W_NS && child.name === "t") {
      const text = textOf(child);
      if (text !== "") into.push({ kind: "text", text });
      continue;
    }
    if (child.namespace === W_NS && child.name === "tab") {
      into.push({ kind: "tab" });
      continue;
    }
    if (child.namespace === W_NS && child.name === "br") {
      into.push({ kind: "break" });
      continue;
    }
    if (child.namespace === WP_NS && child.name === "inline") {
      into.push(extentOf(child));
      continue;
    }
    // A floating anchor is out of flow, so nothing under it belongs to this line.
    if (child.namespace === WP_NS && child.name === "anchor") continue;

    collectPieces(child, into);
  }
}

export function readRuns(paragraph: Paragraph, styles: StyleTable): readonly TextRun[] {
  return resolveRuns(paragraph, styles).map((marked) => {
    const pieces: RunPiece[] = [];
    collectPieces(marked.run, pieces);
    return { mark: marked.mark, pieces };
  });
}
