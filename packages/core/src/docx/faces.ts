import { readAnchors } from "./anchors.js";
import { readBlocks, type Block } from "./blocks.js";
import { numberParagraphs } from "./list-numbers.js";
import { MAIN_DOCUMENT_PART, type DocxPackage } from "./package.js";
import { defaultFooterPart, defaultHeaderPart } from "./relationships.js";
import {
  resolveNumberMark,
  resolveParagraphMark,
  resolveRunMarks,
  readStyleTable,
  type InTable,
  type ParagraphMark,
  type StyleTable,
} from "./styles.js";

// A face the document asks to be laid out in, and the sizes it asks for it at.
// The name is null where the cascade resolved no font at all, which is a document
// that cannot be laid out until its styles name one.
export type UsedFace = {
  readonly name: string | null;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly sizesPt: readonly number[];
};

const keyOf = (mark: ParagraphMark): string =>
  [
    mark.font.kind === "named" ? mark.font.name : "",
    mark.bold ? "b" : "",
    mark.italic ? "i" : "",
  ].join("|");

class Faces {
  private readonly found = new Map<string, { face: UsedFace; sizes: Set<number> }>();

  add(mark: ParagraphMark): void {
    const key = keyOf(mark);
    const seen = this.found.get(key);
    if (seen !== undefined) {
      seen.sizes.add(mark.fontSizePt);
      return;
    }
    this.found.set(key, {
      face: {
        name: mark.font.kind === "named" ? mark.font.name : null,
        bold: mark.bold,
        italic: mark.italic,
        sizesPt: [],
      },
      sizes: new Set([mark.fontSizePt]),
    });
  }

  all(): readonly UsedFace[] {
    return [...this.found.values()]
      .map(({ face, sizes }) => ({ ...face, sizesPt: [...sizes].sort((a, b) => a - b) }))
      .sort((one, other) => (one.name ?? "").localeCompare(other.name ?? ""));
  }
}

// A text box holds its own paragraphs, which are laid out in their own faces and
// so have to be walked as well as the story around them.
//
// **A paragraph inside a table is resolved with its table's style**, which is what
// the layout resolves it with: a table style naming a face reaches the paragraph
// mark of every paragraph in the table, and a face this never named is a face the
// caller never loaded. The runs are asked without it, again as the layout asks them.
function collect(blocks: readonly Block[], styles: StyleTable, into: Faces): void {
  const numbered = numberParagraphs(blocks, styles);

  const visit = (of: readonly Block[], inTable: InTable | null): void => {
    for (const block of of) {
      if (block.kind === "table") {
        const under: InTable = { styleId: block.styleId, at: null };
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks, under);
        continue;
      }

      const paragraph = block.paragraph;
      into.add(resolveParagraphMark(paragraph, styles, inTable));
      for (const mark of resolveRunMarks(paragraph, styles)) into.add(mark);

      const number =
        numbered.kind === "numbered" ? numbered.numbers.get(paragraph.index) : undefined;
      if (number !== undefined) into.add(resolveNumberMark(paragraph, styles, number.level));

      for (const anchor of readAnchors(paragraph)) {
        if (anchor.content.kind === "text-box") collect(anchor.content.body.blocks, styles, into);
      }
    }
  };
  visit(blocks, null);
}

// Every face the document needs before it can be laid out, over the stories that
// are actually drawn: the body, the default header and the default footer.
export function facesUsed(pkg: DocxPackage): readonly UsedFace[] {
  const styles = readStyleTable(pkg);
  const faces = new Faces();
  const parts = [MAIN_DOCUMENT_PART, defaultHeaderPart(pkg), defaultFooterPart(pkg)];

  for (const part of parts) {
    if (part === null) continue;
    collect(readBlocks(pkg, part), styles, faces);
  }

  return faces.all();
}
