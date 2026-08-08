import { strToU8 } from "fflate";

import {
  readFontFile,
  readGlyphIndex,
  DocxPagesError,
  type CodeToGlyph,
  type FontMetrics,
  type ParagraphMark,
} from "@docx-pages/core";

import {
  pdfArray,
  pdfDictionary,
  pdfName,
  pdfNumber,
  pdfStream,
  pdfString,
  type PdfObjects,
  type PdfReference,
  type PdfValue,
} from "./objects.js";
import type { PdfFont } from "./document.js";

// The faces a document is drawn in, written into the file beside the drawing.
//
// A pdf carries its own fonts: nothing about the machine a file is opened on
// decides what a page looks like, which is the whole reason to write one. So every
// face the document draws in goes in whole, under Identity-H, which addresses a
// font by glyph rather than by character.
//
// Writing by glyph is why this file exists at all rather than being three lines in
// `text.ts`. A glyph number means nothing outside the face it came from, so what
// each face was asked to draw has to be gathered as the pages are written and the
// widths and the reverse map built from it afterwards.

const AT = "pdf/document.writePdf";

// Glyph space in a pdf is a thousandth of the em, whatever the face's own units
// per em are.
const GLYPH_UNITS = 1000;

const scaled = (value: number, metrics: FontMetrics): number =>
  Math.round((value * GLYPH_UNITS) / metrics.unitsPerEm);

const normalise = (name: string): string => name.trim().toLowerCase();

type Wanted = {
  readonly name: string;
  readonly bold: boolean;
  readonly italic: boolean;
};

export const faceOf = (mark: ParagraphMark): Wanted => ({
  name: mark.font.kind === "named" ? mark.font.name : "",
  bold: mark.bold,
  italic: mark.italic,
});

const keyOf = (want: Wanted): string =>
  `${normalise(want.name)}|${want.bold ? "b" : ""}${want.italic ? "i" : ""}`;

const supplies = (font: PdfFont, want: Wanted): boolean =>
  normalise(font.name) === normalise(want.name) &&
  (font.bold ?? false) === want.bold &&
  (font.italic ?? false) === want.italic;

/**
 * One face as this file uses it: the name a content stream calls it by, and what
 * to write to draw a string in it.
 */
export type PdfFace = {
  readonly resource: string;
  // The glyphs a string is drawn as, two bytes to each, which is what an
  // Identity-H encoding takes. Every one is recorded, so that the widths and the
  // reverse map written afterwards cover exactly what was drawn.
  readonly glyphsFor: (text: string) => Uint8Array;
};

export type PdfFonts = {
  // The face a run is drawn in. **Refuses the document** where nothing supplies
  // it: there is no stand-in here, and drawing a page in the wrong face is the one
  // thing this project will not do quietly.
  readonly faceFor: (mark: ParagraphMark) => PdfFace;
  // The `/Font` resource dictionary, once every page has been written and it is
  // settled what each face was asked to draw.
  readonly resources: (objects: PdfObjects) => PdfValue;
};

type Used = {
  readonly resource: string;
  readonly font: PdfFont;
  readonly metrics: FontMetrics;
  readonly advanceFor: (codePoint: number) => number | null;
  readonly glyphFor: CodeToGlyph;
  // What the face was asked to draw, by glyph. The character is kept beside it for
  // the reverse map, which is what lets the text be selected and searched.
  readonly drawn: Map<number, number>;
};

function openFace(font: PdfFont, resource: string): Used {
  const read = readFontFile(font.bytes, font.name);
  if (read.advances.kind !== "advances") {
    throw new DocxPagesError({
      code: "font-unmeasurable",
      message: `the supplied ${font.name} states no usable widths: ${read.advances.reason}`,
      at: AT,
      context: { fontName: font.name, reason: read.advances.reason },
    });
  }

  return {
    resource,
    font,
    metrics: read.metrics,
    advanceFor: read.advances.advanceFor,
    glyphFor: readGlyphIndex(font.bytes, font.name),
    drawn: new Map(),
  };
}

export function pdfFonts(fonts: readonly PdfFont[]): PdfFonts {
  const used = new Map<string, Used>();

  const faceFor = (mark: ParagraphMark): PdfFace => {
    const want = faceOf(mark);
    const key = keyOf(want);

    let face = used.get(key);
    if (face === undefined) {
      const supplied = fonts.find((font) => supplies(font, want));
      if (supplied === undefined) {
        throw new DocxPagesError({
          code: "font-not-supplied",
          message: `the document draws in ${want.name || "an unnamed face"} and nothing supplies it; a pdf carries the faces it draws in`,
          at: AT,
          context: {
            fontName: want.name,
            bold: want.bold,
            italic: want.italic,
            supplied: fonts.map((font) => font.name).sort(),
          },
        });
      }
      face = openFace(supplied, `F${String(used.size)}`);
      used.set(key, face);
    }

    const opened = face;
    return {
      resource: opened.resource,
      glyphsFor: (text) => glyphsOf(opened, text),
    };
  };

  return {
    faceFor,
    resources: (objects) =>
      pdfDictionary(
        Object.fromEntries(
          [...used.values()].map((face) => [face.resource, writeFace(objects, face)]),
        ),
      ),
  };
}

// A string as the glyphs the face draws it with, recording each so that the widths
// written afterwards are the widths this text was measured at.
//
// Iterating the string rather than its code units is what keeps a character
// outside the basic plane one character: it is two code units and one glyph.
function glyphsOf(face: Used, text: string): Uint8Array {
  const characters = Array.from(text);
  const out = new Uint8Array(characters.length * 2);
  const view = new DataView(out.buffer);

  characters.forEach((character, at) => {
    const codePoint = character.codePointAt(0) ?? 0;
    const glyph = face.glyphFor(codePoint);
    face.drawn.set(glyph, codePoint);
    view.setUint16(at * 2, glyph);
  });

  return out;
}

// A pdf holds one font object per face and a second describing it, which is the
// shape Identity-H asks for: a Type0 font over a CIDFontType2 descendant, the
// descendant carrying the file and the widths.
function writeFace(objects: PdfObjects, face: Used): PdfReference {
  const name = face.font.name;
  const file = objects.add(
    pdfStream({ Length1: pdfNumber(face.font.bytes.byteLength) }, face.font.bytes),
  );

  const descriptor = objects.add(
    pdfDictionary({
      Type: pdfName("FontDescriptor"),
      FontName: pdfName(name),
      Flags: pdfNumber(flagsOf(face)),
      FontBBox: pdfArray(boundsOf(face.metrics).map(pdfNumber)),
      // Nothing here reads the `post` table, so the angle an italic face states
      // is not known. A reader draws the outlines in the embedded file, which
      // carry the slant themselves; this is what it would fall back on and never
      // does.
      ItalicAngle: pdfNumber(0),
      Ascent: pdfNumber(scaled(face.metrics.ascender, face.metrics)),
      Descent: pdfNumber(scaled(face.metrics.descender, face.metrics)),
      CapHeight: pdfNumber(scaled(face.metrics.ascender, face.metrics)),
      // Required, and read by nothing where the face itself is embedded.
      StemV: pdfNumber(80),
      FontFile2: file,
    }),
  );

  const descendant = objects.add(
    pdfDictionary({
      Type: pdfName("Font"),
      Subtype: pdfName("CIDFontType2"),
      BaseFont: pdfName(name),
      CIDSystemInfo: pdfDictionary({
        Registry: pdfString("Adobe"),
        Ordering: pdfString("Identity"),
        Supplement: pdfNumber(0),
      }),
      FontDescriptor: descriptor,
      // What a glyph the widths do not name is drawn at. Every glyph drawn is
      // named, so this answers for none of them.
      DW: pdfNumber(GLYPH_UNITS),
      W: widthsOf(face),
      // A character identifier is the glyph number itself, which is what makes the
      // whole face reachable without a mapping table of its own.
      CIDToGIDMap: pdfName("Identity"),
    }),
  );

  return objects.add(
    pdfDictionary({
      Type: pdfName("Font"),
      Subtype: pdfName("Type0"),
      BaseFont: pdfName(name),
      Encoding: pdfName("Identity-H"),
      DescendantFonts: pdfArray([descendant]),
      ToUnicode: toUnicodeOf(objects, face),
    }),
  );
}

// Symbolic, which is what a font addressed by glyph is: its characters are the
// face's own rather than a standard encoding's. Italic is stated beside it because
// a reader offering to substitute the face has nothing else to go on.
const SYMBOLIC = 4;
const ITALIC = 64;

const flagsOf = (face: Used): number => SYMBOLIC | ((face.font.italic ?? false) ? ITALIC : 0);

// The box every glyph in the face fits inside. Nothing here reads the `head` table
// that states it, so this is taken from the vertical metrics and made wide enough
// not to cut anything off: a reader draws the outlines in the embedded file and
// uses this only to decide what to repaint.
const boundsOf = (metrics: FontMetrics): readonly number[] => [
  -GLYPH_UNITS,
  scaled(metrics.descender, metrics) - GLYPH_UNITS,
  GLYPH_UNITS * 2,
  scaled(metrics.ascender, metrics) + GLYPH_UNITS,
];

// The width of every glyph the face was asked to draw, written as runs of
// consecutive ones: `first [w w w]`. A glyph the document never drew is not named,
// which is the one place the whole-face embed is not paid for twice.
function widthsOf(face: Used): PdfValue {
  const glyphs = [...face.drawn.keys()].sort((one, other) => one - other);

  const runs: PdfValue[] = [];
  let run: number[] = [];
  let startedAt = 0;

  const close = (): void => {
    if (run.length === 0) return;
    runs.push(pdfNumber(startedAt), pdfArray(run.map(pdfNumber)));
    run = [];
  };

  for (const glyph of glyphs) {
    if (run.length > 0 && glyph !== startedAt + run.length) close();
    if (run.length === 0) startedAt = glyph;
    run.push(widthOf(face, glyph));
  }
  close();

  return pdfArray(runs);
}

// What the glyph advances, in glyph space. Asked at the character it was drawn
// for, which is how the advance table answers, so the width written is the very
// one the line was measured with.
function widthOf(face: Used, glyph: number): number {
  const codePoint = face.drawn.get(glyph);
  const advance = codePoint === undefined ? null : face.advanceFor(codePoint);
  return advance === null ? 0 : scaled(advance, face.metrics);
}

// How many `bfchar` lines a block may hold, which the CMap syntax fixes at 100.
const BLOCK = 100;

const CMAP_HEAD = `/CIDInit /ProcSet findresource begin
12 dict begin
begincmap
/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def
/CMapName /Adobe-Identity-UCS def
/CMapType 2 def
1 begincodespacerange
<0000> <ffff>
endcodespacerange
`;

const CMAP_TAIL = `endcmap
CMapName currentdict /CMap defineresource pop
end
end
`;

const hex = (value: number, digits: number): string =>
  value.toString(16).padStart(digits, "0").slice(-digits);

// The map is written in UTF-16, so a character outside the basic plane takes the
// two code units it is made of. Walked by index rather than iterated, since
// iterating a string gives back whole characters and it is the units that are
// wanted here.
function utf16Of(codePoint: number): string {
  const text = String.fromCodePoint(codePoint);
  let out = "";
  for (let at = 0; at < text.length; at += 1) out += hex(text.charCodeAt(at), 4);
  return out;
}

/**
 * Which character each glyph drawn stands for, so that the text in the page can
 * still be selected, copied and searched.
 *
 * Without this a page drawn under Identity-H holds no text at all as far as a
 * reader is concerned: the glyph numbers are the face's own and mean nothing.
 */
function toUnicodeOf(objects: PdfObjects, face: Used): PdfReference {
  const entries = [...face.drawn.entries()]
    .sort(([one], [other]) => one - other)
    .map(([glyph, codePoint]) => `<${hex(glyph, 4)}> <${utf16Of(codePoint)}>`);

  let body = CMAP_HEAD;
  for (let at = 0; at < entries.length; at += BLOCK) {
    const block = entries.slice(at, at + BLOCK);
    body += `${String(block.length)} beginbfchar\n${block.join("\n")}\nendbfchar\n`;
  }
  body += CMAP_TAIL;

  return objects.add(pdfStream({}, strToU8(body, true)));
}
