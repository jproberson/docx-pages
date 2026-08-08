import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import {
  layOutDocument,
  lookupFontMetrics,
  openDocx,
  readFontFile,
  type LaidOutDocument,
  type MetricsResolver,
  type SuppliedFace,
} from "@docx-pages/core";
import { buildDocx, wordDocument, WORDPROCESSING_NS } from "@docx-pages/core/testing";
import { writePdf, type PdfFont } from "@docx-pages/pdf";

import { readTextPlacements } from "./text.js";

// What the writer produces, read back through the very reader this project holds
// Word's own pdf to. Layout says where a line goes; the file is written; the file
// is read again; the two are held to each other.
//
// This is not a claim about Word. The claim about Word is the layout's, and it is
// made in `packages/render/src/authored` and in the reference suites. This says the
// narrower thing the writer is answerable for: **that a page written out holds the
// text where layout put it**, so that the two backends cannot drift apart without
// something failing.
//
// A real font file is what makes it possible to check at all. The metric-only
// fixtures in `@docx-pages/core/testing` carry no `glyf` table and so cannot be
// embedded in a pdf and read out of one; the pack ships faces that can.

const packFace = async (file: string): Promise<Uint8Array> =>
  new Uint8Array(
    await readFile(fileURLToPath(new URL(`../../../fonts/fonts/${file}`, import.meta.url))),
  );

// One face under the name a document names, which is how a caller supplies a
// stand-in as well: the bytes are what is drawn, and the name is what is asked for.
const FACE_NAME = "Carlito";

let fonts: readonly PdfFont[] = [];
let supplied: readonly SuppliedFace[] = [];
let metricsFor: MetricsResolver;

beforeAll(async () => {
  const regular = await packFace("Carlito-Regular.ttf");
  const bold = await packFace("Carlito-Bold.ttf");

  fonts = [
    { name: FACE_NAME, bytes: regular },
    { name: FACE_NAME, bold: true, bytes: bold },
  ];
  supplied = fonts.map((font) => {
    const read = readFontFile(font.bytes, font.name);
    return {
      name: font.name,
      bold: font.bold ?? false,
      italic: font.italic ?? false,
      metrics: read.metrics,
      advances: read.advances,
      sansSerif: read.sansSerif,
    };
  });
  metricsFor = (request) => lookupFontMetrics(request, supplied);
});

// Letter, at the margins Word starts a document on.
const LETTER = { widthTwips: 12240, heightTwips: 15840 };

const section = (page = LETTER): string =>
  `<w:sectPr><w:pgSz w:w="${String(page.widthTwips)}" w:h="${String(page.heightTwips)}"/>` +
  `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"` +
  ` w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

const run = (text: string, bold = false): string =>
  `<w:r><w:rPr><w:rFonts w:ascii="${FACE_NAME}" w:hAnsi="${FACE_NAME}"/>` +
  `<w:sz w:val="24"/>${bold ? "<w:b/>" : ""}</w:rPr>` +
  `<w:t xml:space="preserve">${text}</w:t></w:r>`;

const paragraph = (...runs: readonly string[]): string => `<w:p>${runs.join("")}</w:p>`;

// A paragraph mark names no face of its own, and the resolver here stands nothing
// in, so the document has to state a floor for one as a real document does.
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${WORDPROCESSING_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
  <w:rFonts w:ascii="${FACE_NAME}" w:hAnsi="${FACE_NAME}"/><w:sz w:val="24"/>
</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`;

function laidOut(body: string, page = LETTER): LaidOutDocument {
  const bytes = buildDocx({
    "word/document.xml": wordDocument(body + section(page)),
    "word/styles.xml": STYLES,
  });
  const layout = layOutDocument(openDocx(bytes), metricsFor);
  if (layout.kind !== "laid-out") throw new Error(`blocked: ${layout.blocker.kind}`);
  return layout;
}

const written = (layout: LaidOutDocument): Uint8Array =>
  writePdf(layout, { fonts, imageBytes: () => undefined, metricsFor });

// Where layout put every line of the body, so the file can be held to it.
const laidOutLines = (layout: LaidOutDocument) =>
  layout.pages.flatMap((page) =>
    page.body.flatMap((box) =>
      box.lines.map((line) => ({
        pageIndex: page.index,
        leftPt: line.leftPt,
        baselinePt: line.baselinePt,
      })),
    ),
  );

const round = (value: number): number => Math.round(value * 100) / 100;

describe("a page written out", () => {
  it("is a pdf the reader opens, holding the text the document holds", async () => {
    const placements = await readTextPlacements(written(laidOut(paragraph(run("Hello")))));

    expect(placements.map((each) => each.text).join("")).toBe("Hello");
  });

  // The whole point of the writer. A run is written at the place layout put it, so
  // reading the file back has to give the layout's own numbers.
  it("draws every line where layout put it", async () => {
    const layout = laidOut(paragraph(run("first line")) + paragraph(run("second line")));
    const placements = await readTextPlacements(written(layout));

    expect(
      placements.map((each) => ({
        pageIndex: each.pageIndex,
        leftPt: round(each.leftPt),
        baselinePt: round(each.baselinePt),
      })),
    ).toStrictEqual(
      laidOutLines(layout).map((line) => ({
        pageIndex: line.pageIndex,
        leftPt: round(line.leftPt),
        baselinePt: round(line.baselinePt),
      })),
    );
  });

  // A tab opens a gap along the line that the runs after it never account for, so
  // a line laid end to end closes it and everything after the tab slides left.
  // Each run is written at its own offset instead.
  //
  // A tab is what makes that visible at all: the reader joins runs that carry on
  // from one another, so a line with no gap in it comes back as one item however
  // many runs went in. Two items here is the gap surviving the round trip.
  it("starts a run after a tab where layout put it, not where the text ran to", async () => {
    const layout = laidOut(paragraph(run("one"), `<w:r><w:tab/></w:r>`, run("two")));
    const placements = await readTextPlacements(written(layout));

    const line = layout.pages[0]?.body[0]?.lines[0];
    const offsets = (line?.line.segments ?? []).flatMap((segment) =>
      segment.kind === "text" ? [round((line?.leftPt ?? 0) + segment.offsetPt)] : [],
    );

    expect(placements.length).toBeGreaterThan(1);
    expect(round(placements[placements.length - 1]?.leftPt ?? 0)).toBe(offsets[offsets.length - 1]);
  });

  // The reader reads a baseline back against the page's own height, so a page
  // written at the wrong size puts every line somewhere else. A page that is not
  // the usual one is what says the height came from the section rather than from a
  // constant.
  it("gives a page the size its own section asks for", async () => {
    const a4 = { widthTwips: 11906, heightTwips: 16838 };
    const layout = laidOut(paragraph(run("only")), a4);
    const placements = await readTextPlacements(written(layout));

    expect(round(placements[0]?.baselinePt ?? 0)).toBe(
      round(layout.pages[0]?.body[0]?.lines[0]?.baselinePt ?? 0),
    );
  });

  it("breaks onto as many pages as layout broke onto", async () => {
    const body = Array.from({ length: 60 }, (_, at) => paragraph(run(`line ${String(at)}`))).join(
      "",
    );
    const layout = laidOut(body);
    const placements = await readTextPlacements(written(layout));

    expect(layout.pages.length).toBeGreaterThan(1);
    expect(new Set(placements.map((each) => each.pageIndex)).size).toBe(layout.pages.length);
  });

  // Without a ToUnicode map a page written under Identity-H holds no text at all
  // as far as a reader is concerned: the glyph numbers are the face's own.
  it("can still be read as the characters the document was written in", async () => {
    const placements = await readTextPlacements(
      written(laidOut(paragraph(run("Quick brown fox")))),
    );

    expect(placements.map((each) => each.text).join("")).toBe("Quick brown fox");
  });
});

describe("a face nothing supplies", () => {
  it("refuses the document rather than drawing it in another", () => {
    const layout = laidOut(paragraph(run("supplied")));

    expect(() => writePdf(layout, { fonts: [], imageBytes: () => undefined, metricsFor })).toThrow(
      /nothing supplies it/,
    );
  });
});
