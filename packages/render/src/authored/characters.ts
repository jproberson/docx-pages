import {
  type MetricsResolver,
  type Paragraph,
  type ParagraphBox,
  type StyleTable,
} from "@docx-pages/core";
import { measureText, readRuns } from "@docx-pages/core/internal";

// Where every character of a paragraph sits along its line, which is the only way
// to say where a tab landed: the line's own start says nothing about what happened
// part way along it.
//
// Word counts a tab as a character of the paragraph and this project's lines do
// not: a tab leaves no segment behind, only a gap before the next one. So the
// paragraph's characters are walked in the order the file spells them, and the
// laid-out segments are consumed alongside: a character the segments spell takes
// its place from them, and one they do not takes the place the last of them
// reached.

export type CharacterPlacement = {
  // Counted from one, as Word counts them.
  readonly index: number;
  readonly leftPt: number;
};

type Spelled = { readonly text: string; readonly spellsText: boolean };

// Every character the paragraph holds, in order, each saying whether it is text
// the layout would have measured or something that only takes a place.
function spelledCharacters(paragraph: Paragraph, styles: StyleTable): readonly Spelled[] {
  const spelled: Spelled[] = [];
  for (const run of readRuns(paragraph, styles)) {
    for (const piece of run.pieces) {
      if (piece.kind !== "text") {
        spelled.push({ text: "", spellsText: false });
        continue;
      }
      for (const text of piece.text) spelled.push({ text, spellsText: true });
    }
  }
  return spelled;
}

type Drawn = { readonly text: string; readonly leftPt: number; readonly endPt: number };

// Each character of a segment, from the point along the line it starts to the one
// the next character begins at.
function segmentCharacters(
  line: ParagraphBox["lines"][number],
  originPt: number,
  metricsFor: MetricsResolver,
): readonly Drawn[] {
  return line.line.segments.flatMap((segment) => {
    if (segment.kind !== "text") return [];
    const startPt = line.leftPt + segment.offsetPt - originPt;

    let widthPt = 0;
    return Array.from(segment.text).map((text) => {
      const leftPt = startPt + widthPt;
      const measured = measureText(text, segment.mark, metricsFor);
      widthPt += measured.kind === "measured" ? measured.widthPt : 0;
      return { text, leftPt, endPt: startPt + widthPt };
    });
  });
}

export function characterPlacements(
  paragraph: Paragraph,
  box: ParagraphBox,
  styles: StyleTable,
  metricsFor: MetricsResolver,
  originPt: number,
): readonly CharacterPlacement[] {
  const drawn = box.lines.flatMap((line) => segmentCharacters(line, originPt, metricsFor));

  const placements: CharacterPlacement[] = [];
  let at = 0;

  const spelled = spelledCharacters(paragraph, styles);
  spelled.forEach((spelled, index) => {
    const next = drawn[at];
    if (spelled.spellsText && next !== undefined && next.text === spelled.text) {
      placements.push({ index: index + 1, leftPt: next.leftPt });
      at += 1;
      return;
    }
    // A tab, a break, or text the line never drew: it stands where the character
    // before it left off, which for a tab is the point the gap opens from.
    placements.push({ index: index + 1, leftPt: drawn[at - 1]?.endPt ?? 0 });
  });

  // Word counts the mark ending the paragraph as a character of it, and places it
  // where the text ran out.
  placements.push({
    index: spelled.length + 1,
    leftPt: drawn[drawn.length - 1]?.endPt ?? 0,
  });

  return placements;
}
