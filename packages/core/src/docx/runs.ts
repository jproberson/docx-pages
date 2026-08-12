import { isDetachedContent, type Paragraph } from "./blocks.js";
import { readDrawingTurn } from "./drawing.js";
import { WP_NS } from "./inlines.js";
import { W_NS } from "./section.js";
import { resolveRuns, type ParagraphMark, type StyleTable } from "./styles.js";
import { inlinePictureOf } from "./vml.js";
import { attribute, firstNamed, type XmlElement } from "./xml.js";

export type RunPiece =
  | { readonly kind: "text"; readonly text: string }
  | { readonly kind: "tab" }
  // A break ends the line it stands on, and one of type "page" starts the line
  // under it on a page of its own. One of type "column" sends the rest of its
  // paragraph to the top of the next column, whatever room is left below it.
  | { readonly kind: "break"; readonly endsPage: boolean; readonly endsColumn: boolean }
  | {
      readonly kind: "drawing";
      readonly widthEmu: number;
      readonly heightEmu: number;
      // The extent is the drawing the right way up, so how far round it was turned
      // is part of how much of the line it takes.
      readonly turnDegrees: number;
    };

export type TextRun = {
  readonly mark: ParagraphMark;
  readonly pieces: readonly RunPiece[];
};

// The whitespace xml leaves insignificant, which is the only kind the edges of a
// w:t lose. A no-break space is a character the text is made of rather than
// whitespace around it, so it stays whether or not the run asks for it.
const INSIGNIFICANT = /^[ \t\r\n]+|[ \t\r\n]+$/g;

// Whitespace at the edges of a w:t is insignificant unless something asks for it,
// which is why Word writes xml:space wherever a space has to survive. It need not be
// the w:t that asks: the attribute stands for everything under the element stating
// it, and a document in the corpus states it once on its own root.
const textOf = (element: XmlElement): string =>
  element.preservesSpace ? element.text : element.text.replace(INSIGNIFICANT, "");

function extentOf(inline: XmlElement): RunPiece {
  const extent = firstNamed(inline, WP_NS, "extent");
  const size = (name: string): number => {
    const raw = extent === null ? undefined : attribute(extent, "", name);
    const value = raw === undefined ? Number.NaN : Number(raw);
    return Number.isFinite(value) ? value : 0;
  };
  return {
    kind: "drawing",
    widthEmu: size("cx"),
    heightEmu: size("cy"),
    turnDegrees: readDrawingTurn(inline),
  };
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
      into.push({
        kind: "break",
        endsPage: attribute(child, W_NS, "type") === "page",
        endsColumn: attribute(child, W_NS, "type") === "column",
      });
      continue;
    }
    if (child.namespace === WP_NS && child.name === "inline") {
      into.push(extentOf(child));
      continue;
    }
    // A floating anchor is out of flow, so nothing under it belongs to this line.
    if (child.namespace === WP_NS && child.name === "anchor") continue;
    if (child.namespace === W_NS && child.name === "pict") {
      const picture = inlinePictureOf(child);
      if (picture !== null) {
        into.push({
          kind: "drawing",
          widthEmu: picture.widthEmu,
          heightEmu: picture.heightEmu,
          turnDegrees: 0,
        });
      }
      continue;
    }

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
