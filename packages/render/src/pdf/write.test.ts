import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import { strFromU8, zlibSync } from "fflate";
import { beforeAll, describe, expect, it } from "vitest";

import {
  bestEffortMetrics,
  layOutDocument,
  lookupFontMetrics,
  openDocx,
  paintOfParagraph,
  pdfOfDocx,
  readFaceShapes,
  readFontFile,
  writePdf,
  type LaidOutDocument,
  type MetricsResolver,
  type PdfFont,
  type SuppliedFace,
} from "@docx-pages/core";
import { buildDocx, wordDocument, WORDPROCESSING_NS } from "@docx-pages/core/testing";

import { readFillPlacements } from "./fills.js";
import { readImagePlacements } from "./placements.js";
import { readTextPlacements } from "./text.js";

// The namespaces a drawing is written through, which `wordDocument` does not
// declare for us.
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

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

  // **A run the file scaled is written stretched, not spaced out.** A pdf has `Tz`
  // for exactly this, and the reader reads the drawn width back: the scaled run has
  // to come back as wide as the scale says, glyph for glyph.
  it("writes a scaled run as wide as the scale says", async () => {
    const scaled = (percent: number): string =>
      `<w:r><w:rPr><w:rFonts w:ascii="${FACE_NAME}" w:hAnsi="${FACE_NAME}"/>` +
      `<w:sz w:val="24"/><w:w w:val="${String(percent)}"/></w:rPr>` +
      `<w:t xml:space="preserve">aaaaaaaa</w:t></w:r>`;

    const widthOf = async (body: string): Promise<number> => {
      const placements = await readTextPlacements(written(laidOut(body)));
      return placements.reduce((sum, each) => sum + each.widthPt, 0);
    };

    const plain = await widthOf(paragraph(run("aaaaaaaa")));
    expect(await widthOf(paragraph(scaled(150)))).toBeCloseTo(plain * 1.5, 1);
    expect(await widthOf(paragraph(scaled(50)))).toBeCloseTo(plain * 0.5, 1);
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

// A 4 by 4 grayscale jpeg, which is a real one rather than a shape a reader would
// have to be lenient about: the bytes are passed straight into the file and the
// reader decodes them as it would any other picture.
const TINY_JPEG = Uint8Array.from(
  Buffer.from(
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAFA3PEY8MlBGQUZaVVBfeMiCeG5uePWvuZHI////////////////////////" +
      "////////////////////////////wAALCAAEAAQBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAAAP/EABQQAQAAAAAAAAAA" +
      "AAAAAAAAAAD/2gAIAQEAAD8AP//Z",
    "base64",
  ),
);

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

describe("what a page paints behind its text", () => {
  // A paragraph's own fill is a rectangle, and a rectangle is where the flip is
  // easiest to get wrong: applied to its top edge rather than its bottom, it lands
  // its own height away from where it belongs.
  it("fills a paragraph where the geometry says, the right way up", async () => {
    const shaded =
      `<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="C00000"/></w:pPr>` +
      `${run("shaded")}</w:p>`;
    const layout = laidOut(shaded);
    const fills = await readFillPlacements(written(layout));

    const box = layout.pages[0]?.body[0];
    const paint = box?.paint;
    if (box === undefined || paint === null || paint === undefined) {
      throw new Error("the paragraph states a fill and the layout carries none");
    }

    const wanted = paintOfParagraph(paint, box.lines[0]?.topPt ?? 0, box.contentBottomPt).fills[0];
    if (wanted === undefined) throw new Error("the geometry paints no fill");

    expect(
      fills.map((fill) => ({
        leftPt: round(fill.leftPt),
        topPt: round(fill.topPt),
        widthPt: round(fill.widthPt),
        heightPt: round(fill.heightPt),
      })),
    ).toContainEqual({
      leftPt: round(wanted.leftPt),
      topPt: round(wanted.topPt),
      widthPt: round(wanted.widthPt),
      heightPt: round(wanted.heightPt),
    });
  });
});

// A turn is stated on the transform in sixty-thousandths of a degree, clockwise,
// which is how every drawing in a real document states one.
const drawing = (widthEmu: number, heightEmu: number, turnDegrees = 0): string =>
  `<w:p><w:r><w:drawing><wp:inline xmlns:wp="${WP_NS}">` +
  `<wp:extent cx="${String(widthEmu)}" cy="${String(heightEmu)}"/>` +
  `<wp:docPr id="1" name="Picture 1"/><a:graphic xmlns:a="${A_NS}">` +
  `<a:graphicData uri="${PIC_NS}"><pic:pic xmlns:pic="${PIC_NS}">` +
  `<pic:blipFill><a:blip xmlns:r="${R_NS}" r:embed="rId9"/></pic:blipFill>` +
  `<pic:spPr><a:xfrm rot="${String(turnDegrees * 60000)}">` +
  `<a:ext cx="${String(widthEmu)}" cy="${String(heightEmu)}"/></a:xfrm>` +
  `</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

const withPicture = (
  widthEmu: number,
  heightEmu: number,
  picture: Uint8Array = TINY_JPEG,
  name = "jpeg",
  turnDegrees = 0,
) => {
  const bytes = buildDocx({
    "word/document.xml": wordDocument(drawing(widthEmu, heightEmu, turnDegrees) + section()),
    "word/styles.xml": STYLES,
    [`word/media/image1.${name}`]: picture,
    "word/_rels/document.xml.rels":
      `<?xml version="1.0"?><Relationships ` +
      `xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId9" Target="media/image1.${name}" Type="${R_NS}/image"/>` +
      `</Relationships>`,
  });
  const pkg = openDocx(bytes);
  const layout = layOutDocument(pkg, metricsFor);
  if (layout.kind !== "laid-out") throw new Error(`blocked: ${layout.blocker.kind}`);
  return { pkg, layout };
};

describe("a picture", () => {
  // A jpeg goes into the file as it stands, so what the reader finds is the very
  // bytes the document held, drawn into the rectangle layout placed it in.
  it("is drawn in the rectangle layout placed it in", async () => {
    const inches = 914400;
    const { pkg, layout } = withPicture(inches * 2, inches);
    const bytes = writePdf(layout, {
      fonts,
      imageBytes: (part) => pkg.parts.get(part),
      metricsFor,
    });

    const placements = await readImagePlacements(bytes);
    const placed = layout.pages[0]?.inlines[0];

    expect(placements).toHaveLength(1);
    expect({
      leftPt: round(placements[0]?.rect.leftPt ?? 0),
      topPt: round(placements[0]?.rect.topPt ?? 0),
      widthPt: round(placements[0]?.rect.widthPt ?? 0),
      heightPt: round(placements[0]?.rect.heightPt ?? 0),
    }).toStrictEqual({
      leftPt: round(placed?.leftPt ?? 0),
      topPt: round(placed?.topPt ?? 0),
      widthPt: round(placed?.widthPt ?? 0),
      heightPt: round(placed?.heightPt ?? 0),
    });
  });

  // Layout answers a turned drawing with the box it stood in before it was turned
  // and how far round it went, and turning it is the writer's to do. A quarter turn
  // is the case worth pinning, since it is the one whose answer can be read off the
  // bounds alone: a box turned a quarter about its own middle covers the width and
  // the height the other way round, over the very same middle.
  it("is drawn turned as far round as layout says it was turned", async () => {
    const inches = 914400;
    const { pkg, layout } = withPicture(inches * 2, inches, TINY_JPEG, "jpeg", 90);
    const bytes = writePdf(layout, {
      fonts,
      imageBytes: (part) => pkg.parts.get(part),
      metricsFor,
    });

    const drawn = (await readImagePlacements(bytes))[0]?.rect;
    const placed = layout.pages[0]?.inlines[0];
    if (drawn === undefined || placed === undefined) throw new Error("nothing was drawn");

    expect(placed.turnDegrees).toBe(90);
    expect(round(drawn.widthPt)).toBe(round(placed.heightPt));
    expect(round(drawn.heightPt)).toBe(round(placed.widthPt));
    expect(round(drawn.leftPt + drawn.widthPt / 2)).toBe(round(placed.leftPt + placed.widthPt / 2));
    expect(round(drawn.topPt + drawn.heightPt / 2)).toBe(round(placed.topPt + placed.heightPt / 2));
  });

  // Bytes that are not there leave the frame empty rather than refusing the page,
  // which is what the viewer does with a picture it cannot resolve.
  it("is left undrawn where nothing answers for its bytes", async () => {
    const { layout } = withPicture(914400, 914400);
    const bytes = writePdf(layout, { fonts, imageBytes: () => undefined, metricsFor });

    expect(await readImagePlacements(bytes)).toStrictEqual([]);
  });
});

// Word draws an underline as a filled rectangle, where the drawn face's own `post`
// table says to put it rather than at a place of its own. Measured on 2026-08-07
// off Word's pdf of a reference document: three runs at 13.92pt, every underline
// 0.1207 em below the baseline and 0.0690 em thick, which are that face's stated
// ratios and no constant of Word's.
describe("an underlined run", () => {
  it("draws its line where the face says to put it", async () => {
    const underlined = `<w:p><w:r><w:rPr><w:rFonts w:ascii="${FACE_NAME}" w:hAnsi="${FACE_NAME}"/>
      <w:sz w:val="24"/><w:u w:val="single"/></w:rPr><w:t>linked</w:t></w:r></w:p>`;
    const layout = laidOut(underlined);
    const bytes = written(layout);

    const line = layout.pages[0]?.body[0]?.lines[0];
    const segment = line?.line.segments[0];
    if (line === undefined || segment?.kind !== "text") {
      throw new Error("the paragraph draws no underlined run");
    }

    const face = readFontFile(fonts[0]?.bytes ?? new Uint8Array());
    const underline = face.underline;
    if (underline === null) throw new Error("the pack face states no post table");

    const em = segment.mark.fontSizePt / face.metrics.unitsPerEm;
    const fills = await readFillPlacements(bytes);

    expect(
      fills.map((fill) => ({
        leftPt: round(fill.leftPt),
        topPt: round(fill.topPt),
        widthPt: round(fill.widthPt),
        heightPt: round(fill.heightPt),
      })),
    ).toStrictEqual([
      {
        leftPt: round(line.leftPt),
        topPt: round(line.baselinePt + underline.position * em),
        widthPt: round(segment.widthPt),
        heightPt: round(underline.thickness * em),
      },
    ]);
  });

  it("leaves a run that asks for no underline unlined", async () => {
    expect(await readFillPlacements(written(laidOut(paragraph(run("plain")))))).toStrictEqual([]);
  });
});

// The whole of the usual path in one call, as `DocxDocument` is for the viewer.
describe("pdfOfDocx", () => {
  const docx = (body: string): Uint8Array =>
    buildDocx({
      "word/document.xml": wordDocument(body + section()),
      "word/styles.xml": STYLES,
    });

  it("opens, lays out and writes, and puts the text where laying out alone would", async () => {
    const body = paragraph(run("through the wrapper"));
    const wrapped = await readTextPlacements(pdfOfDocx(docx(body), { fonts }));
    const byHand = await readTextPlacements(written(laidOut(body)));

    expect(
      wrapped.map((each) => [each.text, round(each.leftPt), round(each.baselinePt)]),
    ).toStrictEqual(byHand.map((each) => [each.text, round(each.leftPt), round(each.baselinePt)]));
  });

  it("takes an ArrayBuffer as readily as the bytes themselves", async () => {
    const bytes = docx(paragraph(run("either way")));
    const buffer = bytes.slice().buffer;

    expect(
      (await readTextPlacements(pdfOfDocx(buffer, { fonts }))).map((each) => each.text),
    ).toStrictEqual(
      (await readTextPlacements(pdfOfDocx(bytes, { fonts }))).map((each) => each.text),
    );
  });

  // Nothing is stood in for here: the resolver is the bare one, so a document
  // naming a face that was not supplied is refused while it is being laid out
  // rather than drawn in something else.
  it("refuses a document naming a face nothing supplies", () => {
    expect(() => pdfOfDocx(docx(paragraph(run("unsupplied"))), { fonts: [] })).toThrow(
      /could not be laid out/,
    );
  });

  it("writes what it is told about the file, and nothing it was not", () => {
    const bytes = pdfOfDocx(docx(paragraph(run("titled"))), {
      fonts,
      metadata: { title: "A written page", producer: "docx-pages" },
    });
    const text = strFromU8(bytes, true);

    expect(text).toContain("/Title (A written page)");
    expect(text).toContain("/Producer (docx-pages)");
    // Nothing here reads a clock, so the same document written twice is the same
    // bytes rather than two files differing in when they were made.
    expect(text).not.toContain("/CreationDate");
    expect([...pdfOfDocx(docx(paragraph(run("twice"))), { fonts })]).toStrictEqual([
      ...pdfOfDocx(docx(paragraph(run("twice"))), { fonts }),
    ]);
  });
});

// A png in a `.docx` is nearly always one that carries alpha, and nearly always
// eight bits deep and drawn in one pass; the corpus sweep is what says so. Both
// paths below are worth holding all the same.
describe("a png picture", () => {
  const crcTable = Uint32Array.from({ length: 256 }, (_, byte) => {
    let value = byte;
    for (let bit = 0; bit < 8; bit += 1) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    return value >>> 0;
  });

  const crc32 = (bytes: Uint8Array): number => {
    let value = 0xffffffff;
    for (const byte of bytes) value = (crcTable[(value ^ byte) & 0xff] ?? 0) ^ (value >>> 8);
    return (value ^ 0xffffffff) >>> 0;
  };

  function chunk(name: string, data: Uint8Array): Uint8Array {
    const out = new Uint8Array(data.byteLength + 12);
    const view = new DataView(out.buffer);
    view.setUint32(0, data.byteLength);
    for (const [at, character] of Array.from(name).entries()) out[4 + at] = character.charCodeAt(0);
    out.set(data, 8);
    view.setUint32(out.byteLength - 4, crc32(out.subarray(4, out.byteLength - 4)));
    return out;
  }

  // One pixel, so the fixture states a whole picture rather than a fragment of one.
  function png(colourType: number, samples: readonly number[], interlaced = false): Uint8Array {
    const header = new Uint8Array(13);
    const view = new DataView(header.buffer);
    view.setUint32(0, 1);
    view.setUint32(4, 1);
    header[8] = 8;
    header[9] = colourType;
    header[12] = interlaced ? 1 : 0;

    const parts = [
      Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      chunk("IHDR", header),
      chunk("IDAT", zlibSync(Uint8Array.from([0, ...samples]))),
      chunk("IEND", new Uint8Array(0)),
    ];
    const out = new Uint8Array(parts.reduce((sum, part) => sum + part.byteLength, 0));
    let at = 0;
    for (const part of parts) {
      out.set(part, at);
      at += part.byteLength;
    }
    return out;
  }

  const wroteWith = (picture: Uint8Array): Uint8Array => {
    const { pkg, layout } = withPicture(914400, 914400, picture, "png");
    return writePdf(layout, { fonts, imageBytes: (part) => pkg.parts.get(part), metricsFor });
  };

  it("is drawn in the rectangle layout placed it in", async () => {
    const { pkg, layout } = withPicture(914400, 457200, png(6, [10, 20, 30, 255]), "png");
    const bytes = writePdf(layout, {
      fonts,
      imageBytes: (part) => pkg.parts.get(part),
      metricsFor,
    });

    const placements = await readImagePlacements(bytes);
    const placed = layout.pages[0]?.inlines[0];

    expect(placements.length).toBeGreaterThan(0);
    expect({
      leftPt: round(placements[0]?.rect.leftPt ?? 0),
      topPt: round(placements[0]?.rect.topPt ?? 0),
      widthPt: round(placements[0]?.rect.widthPt ?? 0),
    }).toStrictEqual({
      leftPt: round(placed?.leftPt ?? 0),
      topPt: round(placed?.topPt ?? 0),
      widthPt: round(placed?.widthPt ?? 0),
    });
  });

  // **The whole point of the path that carries no alpha.** A pdf deflates and
  // predicts its pixels exactly as a png does, so the picture in the document is
  // the picture in the file, byte for byte, never decoded and never compressed a
  // second time.
  it("carries a picture with no alpha across untouched", () => {
    const source = png(2, [10, 20, 30]);
    const written = wroteWith(source);
    const idat = readPngIdat(source);

    expect(strFromU8(written, true)).toContain("/Predictor 15");
    expect(indexOfBytes(written, idat)).toBeGreaterThan(-1);
  });

  // Alpha is the one thing that forces the pixels open, since a png keeps it in
  // with the colour and a pdf keeps it in an image of its own.
  it("gives a picture that carries alpha a soft mask of its own", () => {
    const text = strFromU8(wroteWith(png(6, [10, 20, 30, 128])), true);

    expect(text).toContain("/SMask ");
    expect(text).toContain("/ColorSpace /DeviceGray");
  });

  it("draws an indexed picture out of its own palette", () => {
    const paletted = withPalette(png(3, [1]));
    const text = strFromU8(wroteWith(paletted), true);

    expect(text).toContain("/Indexed /DeviceRGB 1");
  });

  // A png may hold its rows in seven passes instead of one. Left undrawn rather
  // than drawn as the smear that reading it straight would give.
  it("leaves an interlaced picture undrawn rather than drawing it wrongly", async () => {
    expect(await readImagePlacements(wroteWith(png(6, [10, 20, 30, 255], true)))).toStrictEqual([]);
  });
});

// The deflated pixels of a png, which is what the no-alpha path hands to the pdf.
function readPngIdat(bytes: Uint8Array): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let at = 8;
  while (at + 12 <= bytes.byteLength) {
    const length = view.getUint32(at);
    const name = String.fromCharCode(...bytes.subarray(at + 4, at + 8));
    if (name === "IDAT") return bytes.subarray(at + 8, at + 8 + length);
    at += length + 12;
  }
  throw new Error("the fixture holds no pixels");
}

// A palette put into a png that states none, so an indexed fixture has one.
function withPalette(bytes: Uint8Array): Uint8Array {
  const table = Uint8Array.from([255, 0, 0, 0, 0, 255]);
  const out = new Uint8Array(bytes.byteLength + table.byteLength + 12);
  const at = 8 + 25;
  out.set(bytes.subarray(0, at), 0);

  const entry = new Uint8Array(table.byteLength + 12);
  const view = new DataView(entry.buffer);
  view.setUint32(0, table.byteLength);
  for (const [index, character] of Array.from("PLTE").entries())
    entry[4 + index] = character.charCodeAt(0);
  entry.set(table, 8);
  out.set(entry, at);
  out.set(bytes.subarray(at), at + entry.byteLength);
  return out;
}

function indexOfBytes(haystack: Uint8Array, needle: Uint8Array): number {
  outer: for (let at = 0; at + needle.byteLength <= haystack.byteLength; at += 1) {
    for (let index = 0; index < needle.byteLength; index += 1) {
      if (haystack[at + index] !== needle[index]) continue outer;
    }
    return at;
  }
  return -1;
}

// The seam the viewer's export hangs on, which is the one thing about it that can
// be wrong everywhere at once.
//
// A best-effort layout measures a face it was never given in the widths of one it
// was, and the viewer paints it by offering the browser the stand-in's bytes under
// the name the document asked for. A file has to carry the same face under the same
// name, and whether that is the asked-for name or the stand-in's own cannot be
// reasoned about from either side: the writer refuses a face nothing supplies, so
// getting it the wrong way round refuses every document that ever stood a face in.
describe("a face the layout stood in for", () => {
  const CAMBRIA = "Cambria";

  const namedRun = (face: string, text: string): string =>
    `<w:r><w:rPr><w:rFonts w:ascii="${face}" w:hAnsi="${face}"/><w:sz w:val="24"/></w:rPr>` +
    `<w:t xml:space="preserve">${text}</w:t></w:r>`;

  // Nothing is supplied for exactness, so every face falls to the defaults, which
  // is the state the viewer is in when it has only its font pack behind it.
  const stoodIn = (): { readonly layout: LaidOutDocument; readonly asked: readonly string[] } => {
    const bytes = buildDocx({
      "word/document.xml": wordDocument(paragraph(namedRun(CAMBRIA, "stood in for")) + section()),
      "word/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${WORDPROCESSING_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
  <w:rFonts w:ascii="${CAMBRIA}" w:hAnsi="${CAMBRIA}"/><w:sz w:val="24"/>
</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    });
    const pkg = openDocx(bytes);
    const faces = bestEffortMetrics([], defaults(), readFaceShapes(pkg));
    const layout = layOutDocument(pkg, faces);
    if (layout.kind !== "laid-out") throw new Error(`blocked: ${layout.blocker.kind}`);

    return { layout, asked: faces.substitutions().map((each) => each.requested.name) };
  };

  const defaults = () => ({
    faces: supplied,
    twins: {},
    sansSerif: FACE_NAME,
    serif: FACE_NAME,
    monospace: FACE_NAME,
    lastResort: FACE_NAME,
  });

  it("stands a face in at all, so the rest of this is worth asking", () => {
    expect(stoodIn().asked).toContain(CAMBRIA);
  });

  // Carried under the name the document asked for, drawn out of the bytes that
  // stood in, which is exactly what `facesPaintedWith` builds for the viewer.
  it("is written where the stand-in's bytes are carried under the asked-for name", async () => {
    const { layout } = stoodIn();
    const bytes = writePdf(layout, {
      fonts: fonts.map((font) => ({ ...font, name: CAMBRIA })),
      imageBytes: () => undefined,
      metricsFor: (request) => lookupFontMetrics(request, supplied),
    });

    const drawn = await readTextPlacements(bytes);
    expect(drawn.map((each) => each.text).join("")).toContain("stood in for");
  });

  // The other way round, which is the mistake worth having a test for: handed the
  // stand-in's own name, the writer finds nothing answering for what the layout
  // measured and refuses the document rather than writing it in the wrong face.
  it("refuses the document where the bytes are carried under the stand-in's own name", () => {
    const { layout } = stoodIn();

    expect(() =>
      writePdf(layout, {
        fonts,
        imageBytes: () => undefined,
        metricsFor: (request) => lookupFontMetrics(request, supplied),
      }),
    ).toThrow(/nothing supplies it/);
  });
});
