import { describe, expect, it } from "vitest";

import {
  layOutDocument,
  lookupFontMetrics,
  openDocx,
  type LaidOutDocument,
  type LineSegment,
  type SuppliedFace,
} from "@onepager/core";
import { buildDocx } from "@onepager/core/testing";

import { suppliedFaces } from "../testing/cases.js";

// No reference document justifies a paragraph, so this one is authored here and
// measured against Word itself: the positions below are what Word reported for
// each of these paragraphs, in whole points from the left edge of the text.
// Every word of every line is pinned, which is what makes the rules readable off
// a failure: a share going to the wrong place moves one word and leaves the rest.
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const FILLER =
  "chlorophyll quadrature windbreak granulation microphone repository " +
  "thunderclap pendulum wavelength cartography hemisphere kaleidoscope";

const TAIL = "granulation microphone repository";

const NO_BREAK_SPACE = "\u00a0";

const run = (text: string, halfPoints?: number): string => {
  const size =
    halfPoints === undefined ? "" : `<w:rPr><w:sz w:val="${String(halfPoints)}"/></w:rPr>`;
  return `<w:r>${size}<w:t xml:space="preserve">${text}</w:t></w:r>`;
};

const BREAK = "<w:r><w:br/></w:r>";
const TAB = "<w:r><w:tab/></w:r>";

type Pinned = {
  readonly title: string;
  readonly content: string;
  readonly indent?: string;
  // Where Word started each word of each line, in points from the text's left edge.
  readonly wordsPt: readonly string[];
};

const CASES: readonly Pinned[] = [
  {
    title: "fills every line but the last, sharing what is left over among the spaces",
    content: run(FILLER),
    wordsPt: ["0 126 253 375", "0 133 252 385", "0 99 202 303"],
  },
  {
    title: "gives a space the same share whatever size the run it belongs to is set in",
    content:
      run("chlorophyll quadrature ") +
      run("windbreak granulation microphone ", 20) +
      run("repository thunderclap pendulum wavelength"),
    wordsPt: ["0 105 210 265 323 384", "0 103 191"],
  },
  {
    title: "fills the line a manual break ended, which is not the paragraph's last",
    content:
      run("chlorophyll quadrature windbreak") +
      BREAK +
      run("granulation microphone repository thunderclap pendulum"),
    wordsPt: ["0 191 382", "0 125 254 369", "0"],
  },
  {
    title: "leaves a space the line ended on out of the sharing, break or no break",
    content: run("chlorophyll quadrature windbreak ") + BREAK + run(TAIL),
    wordsPt: ["0 191 382", "0 98 201"],
  },
  {
    title: "fills a hanging indent's lines to their own edges",
    content: run(FILLER),
    indent: `<w:ind w:left="1440" w:hanging="1440"/>`,
    wordsPt: ["0 126 253 375", "72 181 276 385", "72 220 371", "72"],
  },
  {
    title: "hands a double space twice what a single one gets",
    content: run("chlorophyll quadrature  windbreak granulation microphone repository"),
    wordsPt: ["0 117 262 375", "0 103"],
  },
  {
    title: "leaves a tab holding the stop it reached, taking no share",
    content:
      run("chlorophyll") + TAB + run("quadrature windbreak granulation microphone repository"),
    wordsPt: ["0 108 244 375", "0 103"],
  },
  {
    title: "leaves a no-break space out of the sharing, as part of the word around it",
    content: run(
      `chlorophyll${NO_BREAK_SPACE}quadrature windbreak granulation microphone repository`,
    ),
    wordsPt: ["0 238 375", "0 103"],
  },
];

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="40"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults></w:styles>`;

// Letter, an inch of margin all round, so the text has 468pt to fill.
const SECTION =
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`;

const documentOf = (each: Pinned): Uint8Array =>
  buildDocx({
    "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>
<w:p><w:pPr>${each.indent ?? ""}<w:jc w:val="both"/></w:pPr>${each.content}</w:p>
${SECTION}</w:body></w:document>`,
    "word/styles.xml": STYLES,
  });

const TEXT_LEFT_PT = 72;

// Word reports where each run of characters with nothing between it starts. A
// segment of nothing but spaces closes the word before it, and so does a gap the
// drawing leaves, which is what a tab is.
function wordsOf(line: readonly LineSegment[], leftPt: number): string {
  const starts: number[] = [];
  let reached: number | null = null;

  for (const segment of line) {
    if (segment.kind !== "text") continue;
    const startPt = leftPt + segment.offsetPt;
    if (/^[^\S\u00a0]+$/.test(segment.text)) {
      reached = null;
      continue;
    }
    if (reached === null || startPt > reached + 1e-6) starts.push(startPt - TEXT_LEFT_PT);
    reached = startPt + segment.widthPt;
  }

  return starts.map((each) => String(Math.round(each))).join(" ");
}

const linesOf = (layout: LaidOutDocument): readonly string[] =>
  layout.pages.flatMap((page) =>
    page.body.flatMap((box) => box.lines.map((line) => wordsOf(line.line.segments, line.leftPt))),
  );

const CALIBRI: SuppliedFace | undefined = suppliedFaces().find(
  (face) => face.name === "Calibri" && !face.bold && !face.italic,
);

describe.skipIf(CALIBRI === undefined)("a justified paragraph against Word", () => {
  const faces = CALIBRI === undefined ? [] : [CALIBRI];

  for (const each of CASES) {
    it(each.title, () => {
      const layout = layOutDocument(openDocx(documentOf(each)), (request) =>
        lookupFontMetrics(request, faces),
      );
      if (layout.kind !== "laid-out") throw new Error(JSON.stringify(layout.blocker));

      expect(linesOf(layout)).toStrictEqual(each.wordsPt);
    });
  }
});
