import { describe, expect, it } from "vitest";

import {
  layOutDocument,
  lineHeightPt,
  openDocx,
  substitutingMetrics,
  WORD_FALLBACK_FACES,
  WORD_SERIF_FALLBACK_FACE,
} from "@docx-pages/core";
import { buildDocx, wordDocument } from "@docx-pages/core/testing";

import { fallbackFaces } from "./fallback.js";
import { installedFaces } from "./installed.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const cutsOf = (name: string) => fallbackFaces().filter((face) => face.name === name);

const cut = (name: string, bold: boolean, italic: boolean) =>
  cutsOf(name).find((face) => face.bold === bold && face.italic === italic);

const SECTION =
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720"/></w:sectPr>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="${WORD_SERIF_FALLBACK_FACE}" w:hAnsi="${WORD_SERIF_FALLBACK_FACE}"/>
<w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`;

// The pack ahead of the disk scan, which is the order a sweep gathers its faces in
// and the whole of what this pins: the scan reads Word's own directory first, and
// Word's Times New Roman states no line gap at all. See `corpusFaces`.
const asASweepGathers = () => [...fallbackFaces(), ...installedFaces()];

// Two paragraphs of each cut, so a line's height is the distance from one repeat to
// the next rather than a difference of two rounded answers.
function stepPt(style: string): number {
  const paragraph =
    `<w:p><w:pPr><w:rPr>${style}</w:rPr></w:pPr>` +
    `<w:r><w:rPr>${style}</w:rPr><w:t>Repeat</w:t></w:r></w:p>`;
  const bytes = buildDocx({
    "word/document.xml": wordDocument(paragraph + paragraph + SECTION),
    "word/styles.xml": STYLES,
  });
  const laid = layOutDocument(
    openDocx(bytes),
    substitutingMetrics(asASweepGathers(), WORD_FALLBACK_FACES),
  );
  if (laid.kind !== "laid-out") throw new Error(JSON.stringify(laid.blocker));

  const tops = (laid.pages[0]?.body ?? []).flatMap((block) =>
    "lines" in block ? (block.lines[0] === undefined ? [] : [block.lines[0].topPt]) : [],
  );
  const [first, second] = tops;
  if (first === undefined || second === undefined) throw new Error("expected two repeats");
  return second - first;
}

describe("the fallback pack", () => {
  it("offers each cut once and no cut twice", () => {
    const keys = fallbackFaces().map(
      (face) => `${face.name}|${face.bold ? "b" : ""}|${face.italic ? "i" : ""}`,
    );
    expect(new Set(keys).size).toBe(keys.length);
  });
});

// Which faces this machine holds decides what there is to say here, and a machine
// without Times New Roman has nothing: see `fallbackFacePath`.
describe.skipIf(cutsOf(WORD_SERIF_FALLBACK_FACE).length === 0)(
  "Times New Roman, which is two files a cut",
  () => {
    it("supplies the bold, the italic and the bold italic beside the regular", () => {
      expect(cutsOf(WORD_SERIF_FALLBACK_FACE)).toHaveLength(4);
    });

    // The system's copy states a line gap of 87 units and Word's own states none.
    // Reading the bold out of Word's copy puts every bold line 0.4248pt short at
    // 10pt, which is what drifts a reference list down the page a line at a time.
    it("states one line height for all four, which is the system's copy's", () => {
      const heights = [
        [false, false],
        [true, false],
        [false, true],
        [true, true],
      ].map(([bold, italic]) => {
        const face = cut(WORD_SERIF_FALLBACK_FACE, bold === true, italic === true);
        return face === undefined ? null : lineHeightPt(face.metrics, 10);
      });

      expect(new Set(heights).size).toBe(1);
    });

    // Word's own copy is still read for the hyphen the system's has not got at all,
    // which is the reason two files are named for one cut.
    it("measures U+2010 in every cut, which only Word's copy carries", () => {
      const advances = cutsOf(WORD_SERIF_FALLBACK_FACE).map((face) =>
        face.advances.kind === "advances" ? face.advances.advanceFor(0x2010) : null,
      );

      expect(advances.filter((each) => each !== null)).toHaveLength(4);
    });

    // The pack is gathered ahead of the disk scan, so a face it offers shadows the
    // same file found there and anything it leaves off is left off for the whole
    // sweep. Dropping the pairs measured a run that asks to kern without them: a
    // line of `AV To Ta Wa Yo AWAY` in Times New Roman 12pt is 8.78pt narrower
    // with them than without.
    it("states the pairs the file does, which the disk scan has always carried", () => {
      const pairs = cutsOf(WORD_SERIF_FALLBACK_FACE).map((face) =>
        face.kerning?.kind === "kerning" ? face.kerning.kerningBetween(0x41, 0x56) : null,
      );

      expect(pairs.filter((each) => each !== null && each < 0)).toHaveLength(4);
    });

    it("lays a bold line out as tall as a regular one", () => {
      expect(stepPt("<w:b/>")).toBeCloseTo(stepPt(""), 4);
      expect(stepPt("<w:i/>")).toBeCloseTo(stepPt(""), 4);
    });
  },
);
