import type { CropInsets } from "../docx/drawing.js";
import { METAFILE_EXTENSION, OLD_METAFILE_EXTENSION, pictureExtension } from "../docx/pictures.js";
import { METAFILE_PEN_OFFSET } from "../layout/drawables.js";
import { pngFromMetafile } from "../metafile/wmf.js";
import type { MetricsResolver } from "../layout/stack.js";
import {
  type MetafilePicture,
  type MetafileRect,
  type MetafileShape,
  readMetafilePicture,
} from "../metafile/picture.js";

import { bottomOf, type PdfPage } from "./coordinates.js";
import type { Content } from "./content.js";
import type { PdfFonts } from "./fonts.js";
import {
  pdfArray,
  pdfDictionary,
  pdfHexString,
  pdfName,
  pdfNumber,
  pdfStream,
  type PdfEntries,
  type PdfObjects,
  type PdfReference,
  type PdfValue,
} from "./objects.js";
import type { ObjectDrawable } from "./objects-paint.js";
import { hasAlpha, readPng, samplesOf, splitAlpha, type PngImage } from "./png.js";

// The two kinds of picture a `.docx` holds that this writes, which are not the
// same kind of thing at all.
//
// A jpeg is already the compression a pdf would have applied to it, so it goes in
// as it stands, under the filter it is already in. Nothing is decoded and nothing
// is re-encoded, and the file grows by exactly the size of the picture.
//
// A metafile is not a picture: it is a recording of the drawing that made one, and
// `readMetafilePicture` has already played it. So it is written as the lines and
// the text it records, which stay sharp at whatever size the frame gives them,
// rather than being rasterised into something that would not.

export type PdfImages = {
  // Draws the picture where the object stands, and answers whether it drew
  // anything: bytes that are not there, and a jpeg in a colour space this cannot
  // name, leave the frame empty rather than refusing the page.
  readonly draw: (
    out: Content,
    page: PdfPage,
    at: ObjectDrawable,
    part: string,
    crop: CropInsets,
  ) => boolean;
  // The `/XObject` resources the pages turned out to need, or nothing where no
  // page drew a bitmap at all.
  readonly resources: () => PdfValue | null;
};

export type ImageOptions = {
  readonly imageBytes: (part: string) => Uint8Array | undefined;
  readonly metricsFor: MetricsResolver;
  readonly fonts: PdfFonts;
  readonly objects: PdfObjects;
  // Told once about each picture that reached the page and was drawn nowhere on
  // it. See `undrawn` on `WritePdfOptions` for why the writer has to say.
  readonly onUndrawn?: (part: string) => void;
};

type Bitmap = { readonly resource: string; readonly object: PdfReference };

export function pdfImages(options: ImageOptions): PdfImages {
  // A picture used twice is written once, so a document putting the same logo on
  // every page carries it once however many pages there are.
  const bitmaps = new Map<string, Bitmap | null>();
  const metafiles = new Map<string, MetafilePicture | null>();

  const bitmapFor = (part: string): Bitmap | null => {
    const known = bitmaps.get(part);
    if (known !== undefined) return known;

    const object = writeBitmap(options, part);
    const written = object === null ? null : { resource: `Im${String(bitmaps.size)}`, object };
    bitmaps.set(part, written);
    return written;
  };

  const metafileFor = (part: string): MetafilePicture | null => {
    const known = metafiles.get(part);
    if (known !== undefined) return known;

    const bytes = options.imageBytes(part);
    const picture = bytes === undefined ? null : readMetafilePicture(bytes, options.metricsFor);
    metafiles.set(part, picture);
    return picture;
  };

  // Said once a part, however many pages drew it, so a logo on every page of a
  // long document is one line and not two hundred.
  const told = new Set<string>();
  const undrawn = (part: string): false => {
    if (options.onUndrawn !== undefined && !told.has(part)) {
      told.add(part);
      options.onUndrawn(part);
    }
    return false;
  };

  return {
    draw: (out, page, at, part, crop) => {
      if (pictureExtension(part) === METAFILE_EXTENSION) {
        const picture = metafileFor(part);
        if (picture === null) return undrawn(part);
        playMetafile(out, page, at, picture, crop, options.fonts);
        return true;
      }

      const bitmap = bitmapFor(part);
      if (bitmap === null) return undrawn(part);
      drawBitmap(out, page, at, bitmap.resource, crop);
      return true;
    },
    resources: () => {
      const entries = [...bitmaps.values()].flatMap((bitmap) =>
        bitmap === null ? [] : [[bitmap.resource, bitmap.object] as const],
      );
      return entries.length === 0 ? null : pdfDictionary(Object.fromEntries(entries));
    },
  };
}

// A jpeg's own header says how large it is and how many channels it has, which is
// all a pdf needs to be told about bytes it is not going to decode.
type Jpeg = {
  readonly widthPixels: number;
  readonly heightPixels: number;
  readonly components: number;
};

// The frame markers, which are the ones carrying the size. C4, C8 and CC are a
// Huffman table, an extension and an arithmetic-coding table, and share the range
// without sharing the shape.
const isFrameMarker = (marker: number): boolean =>
  marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;

// Markers that carry no length after them: padding, and the standalone ones.
const isStandalone = (marker: number): boolean =>
  marker === 0xff || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd9);

function readJpeg(bytes: Uint8Array): Jpeg | null {
  if (bytes[0] !== 0xff || bytes[1] !== 0xd8) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  let at = 2;
  while (at + 4 <= bytes.byteLength) {
    if (bytes[at] !== 0xff) return null;

    const marker = bytes[at + 1] ?? 0;
    if (isStandalone(marker)) {
      at += 2;
      continue;
    }

    if (isFrameMarker(marker)) {
      if (at + 10 > bytes.byteLength) return null;
      return {
        heightPixels: view.getUint16(at + 5),
        widthPixels: view.getUint16(at + 7),
        components: bytes[at + 9] ?? 0,
      };
    }

    const length = view.getUint16(at + 2);
    if (length < 2) return null;
    at += 2 + length;
  }
  return null;
}

// What a pdf calls a picture of that many channels. A four-channel jpeg is CMYK,
// which Word writes inverted often enough that drawing one the wrong way round is
// worse than not drawing it: it is left undrawn, and said so in the README.
const COLOR_SPACES: Readonly<Record<number, string>> = {
  1: "DeviceGray",
  3: "DeviceRGB",
};

function writeBitmap(options: ImageOptions, part: string): PdfReference | null {
  const bytes = options.imageBytes(part);
  if (bytes === undefined) return null;

  // The older metafile records a bitmap rather than a drawing, so what is written
  // is the png that bitmap makes, through the same reader every other png goes
  // through.
  if (pictureExtension(part) === OLD_METAFILE_EXTENSION) {
    const png = pngFromMetafile(bytes);
    return png === null ? null : writePng(options, png);
  }

  return writeJpeg(options, bytes) ?? writePng(options, bytes);
}

function writeJpeg(options: ImageOptions, bytes: Uint8Array): PdfReference | null {
  const jpeg = readJpeg(bytes);
  const space = jpeg === null ? undefined : COLOR_SPACES[jpeg.components];
  if (jpeg === null || space === undefined) return null;

  return options.objects.add(
    pdfStream(
      {
        Type: pdfName("XObject"),
        Subtype: pdfName("Image"),
        Width: pdfNumber(jpeg.widthPixels),
        Height: pdfNumber(jpeg.heightPixels),
        ColorSpace: pdfName(space),
        BitsPerComponent: pdfNumber(8),
        // Already the compression a pdf would have applied. Passed through rather
        // than decoded and deflated, which would grow the file and lose nothing
        // back.
        Filter: pdfName("DCTDecode"),
      },
      bytes,
      false,
    ),
  );
}

// What a pdf calls a png's pixels. An indexed one names its palette instead, which
// is written beside it below.
const PNG_COLOR_SPACES: Readonly<Record<number, string>> = {
  0: "DeviceGray",
  2: "DeviceRGB",
  4: "DeviceGray",
  6: "DeviceRGB",
};

// A png filters its rows exactly as a pdf's predictor 15 does, so the two agree on
// what the deflated bytes mean and the pixels never have to be opened.
const PNG_PREDICTOR = 15;

const BITS = 8;

/**
 * A png written into the file.
 *
 * Two paths, and which one a picture takes is decided by whether it carries alpha.
 *
 * **Without alpha the bytes go across untouched.** A pdf deflates and predicts its
 * pixels the same way a png does, so the IDAT stream is already a pdf image
 * stream: it is written as it stands, with the predictor named, and nothing is
 * inflated, decoded or compressed again. A picture written this way is the very
 * picture the document held.
 *
 * **With alpha the pixels have to be opened**, because a png keeps what shows
 * through a picture in with the colour and a pdf keeps it in an image of its own.
 * They are inflated, unfiltered, split, and deflated again as two streams.
 */
function writePng(options: ImageOptions, bytes: Uint8Array): PdfReference | null {
  const png = readPng(bytes);
  if (png === null) return null;

  // The corpus sweep finds no png of another depth at all, and interlaced ones
  // vanishingly rare. An interlaced png holds its rows in seven passes that would
  // have to be woven back together, and is left undrawn rather than drawn as the
  // smear that reading it straight would give. The README names both.
  if (png.interlaced || png.bitDepth !== BITS) return null;

  const width = pdfNumber(png.widthPixels);
  const height = pdfNumber(png.heightPixels);
  const shared: PdfEntries = {
    Type: pdfName("XObject"),
    Subtype: pdfName("Image"),
    Width: width,
    Height: height,
    BitsPerComponent: pdfNumber(BITS),
  };

  if (!hasAlpha(png.colourType)) {
    const space = colourSpaceOf(png);
    if (space === null) return null;
    return options.objects.add(
      pdfStream(
        {
          ...shared,
          ColorSpace: space,
          Mask: maskOf(png),
          Filter: pdfName("FlateDecode"),
          DecodeParms: pdfDictionary({
            Predictor: pdfNumber(PNG_PREDICTOR),
            Colors: pdfNumber(samplesOf(png.colourType)),
            Columns: width,
            BitsPerComponent: pdfNumber(BITS),
          }),
        },
        png.deflated,
        false,
      ),
    );
  }

  const split = splitAlpha(png);
  const space = PNG_COLOR_SPACES[png.colourType];
  if (split === null || space === undefined) return null;

  // What shows through the picture, as a greyscale image of its own the same size.
  const soft = options.objects.add(
    pdfStream({ ...shared, ColorSpace: pdfName("DeviceGray") }, split.alpha),
  );

  return options.objects.add(
    pdfStream({ ...shared, ColorSpace: pdfName(space), SMask: soft }, split.colour),
  );
}

function colourSpaceOf(png: PngImage): PdfValue | null {
  if (png.colourType !== 3) {
    const named = PNG_COLOR_SPACES[png.colourType];
    return named === undefined ? null : pdfName(named);
  }

  // An indexed png draws out of its own palette, which the pdf carries beside it
  // rather than the picture being expanded into full colour.
  const palette = png.palette;
  if (palette === null || palette.byteLength === 0) return null;
  return pdfArray([
    pdfName("Indexed"),
    pdfName("DeviceRGB"),
    pdfNumber(Math.floor(palette.byteLength / 3) - 1),
    pdfHexString(palette),
  ]);
}

/**
 * The palette entries a png says are not drawn at all.
 *
 * `tRNS` states an alpha for each entry, and only an indexed png in these
 * documents ever carries one. Where every alpha it states is all or nothing, the
 * entries that are nothing become a colour-key mask, which costs nothing: the
 * picture still goes across untouched.
 *
 * A palette that is **partly** transparent would need the pixels opened to make a
 * soft mask of, and the pass-through given up with them. Nothing here does that,
 * so such a png is drawn opaque rather than not at all: the shape is right and only
 * what shows through it is wrong. The README names it.
 */
function maskOf(png: PngImage): PdfValue | undefined {
  const stated = png.transparency;
  if (png.colourType !== 3 || stated === null) return undefined;

  const ranges: PdfValue[] = [];
  for (const [index, alpha] of stated.entries()) {
    if (alpha === 0xff) continue;
    if (alpha !== 0) return undefined;
    ranges.push(pdfNumber(index), pdfNumber(index));
  }
  return ranges.length === 0 ? undefined : pdfArray(ranges);
}

// An image draws into the unit square, so the transform that places it is its
// whole geometry.
//
// A source rectangle hides a fraction of each edge, so the whole bitmap is larger
// than the placed rectangle by exactly that much and is drawn behind it, cut to
// the frame. The viewer does the same with an oversized `img` inside an
// `overflow: hidden` box.
//
// A document is free to state a rectangle that hides the whole picture, and there
// is no bitmap large enough to show none of itself: the frame is left empty, as a
// window of no width leaves a metafile's below.
function drawBitmap(
  out: Content,
  page: PdfPage,
  at: ObjectDrawable,
  resource: string,
  crop: CropInsets,
): void {
  const acrossShown = 1 - crop.left - crop.right;
  const downShown = 1 - crop.top - crop.bottom;
  if (acrossShown <= 0 || downShown <= 0) return;

  const widthPt = at.widthPt / acrossShown;
  const heightPt = at.heightPt / downShown;
  const leftPt = at.leftPt - crop.left * widthPt;
  const topPt = at.topPt - crop.top * heightPt;

  out.save();
  out.rectangle(at.leftPt, bottomOf(page, at.topPt, at.heightPt), at.widthPt, at.heightPt);
  out.clip();
  out.transform([widthPt, 0, 0, heightPt, leftPt, bottomOf(page, topPt, heightPt)]);
  out.drawObject(resource);
  out.restore();
}

// Where a metafile's pen stands, which `drawables.ts` states for both backends.
const PEN_OFFSET = METAFILE_PEN_OFFSET;

/**
 * A played metafile, written into the frame the document gave it.
 *
 * The frame decides the scale on its own, whatever the recording's own proportions
 * were, and a source rectangle is a narrower window onto the same coordinates
 * rather than a larger picture behind a smaller frame. That window is what the
 * transform below maps onto the object's own rectangle.
 *
 * Metafile units count down the page, as a device context does, so the transform
 * flips as well as scales. Text is the one thing that has to undo the flip for
 * itself: a glyph drawn under it would come out mirrored.
 */
function playMetafile(
  out: Content,
  page: PdfPage,
  at: ObjectDrawable,
  picture: MetafilePicture,
  crop: CropInsets,
  fonts: PdfFonts,
): void {
  const { widthUnits, heightUnits } = picture;
  const windowWidth = widthUnits * (1 - crop.left - crop.right);
  const windowHeight = heightUnits * (1 - crop.top - crop.bottom);
  if (windowWidth <= 0 || windowHeight <= 0) return;

  const acrossPt = at.widthPt / windowWidth;
  const downPt = at.heightPt / windowHeight;
  const windowLeft = crop.left * widthUnits;
  const windowTop = crop.top * heightUnits;

  out.save();
  out.rectangle(at.leftPt, bottomOf(page, at.topPt, at.heightPt), at.widthPt, at.heightPt);
  out.clip();
  out.transform([
    acrossPt,
    0,
    0,
    -downPt,
    at.leftPt - windowLeft * acrossPt,
    page.heightPt - at.topPt + windowTop * downPt,
  ]);

  for (const shape of picture.shapes) {
    out.save();
    if (shape.clipTo !== null) clipToRect(out, shape.clipTo);
    metafileShape(out, shape, fonts);
    out.restore();
  }

  out.restore();
}

// In the metafile's own units, which the transform above has already been given,
// so the rectangle is stated the way the recording states it and reaches the page
// through the one flip.
function clipToRect(out: Content, rect: MetafileRect): void {
  out.rectangle(
    rect.leftUnits,
    rect.topUnits + rect.heightUnits,
    rect.widthUnits,
    -rect.heightUnits,
  );
  out.clip();
}

function metafileShape(out: Content, shape: MetafileShape, fonts: PdfFonts): void {
  switch (shape.kind) {
    case "fill":
      out.fillColor(shape.color);
      out.rectangle(
        shape.rect.leftUnits,
        shape.rect.topUnits + shape.rect.heightUnits,
        shape.rect.widthUnits,
        -shape.rect.heightUnits,
      );
      out.fill();
      return;
    case "stroke":
      out.strokeColor(shape.color);
      out.lineWidth(shape.widthUnits);
      out.dash(null);
      out.line(
        shape.fromXUnits + PEN_OFFSET,
        shape.fromYUnits + PEN_OFFSET,
        shape.toXUnits + PEN_OFFSET,
        shape.toYUnits + PEN_OFFSET,
      );
      out.stroke();
      return;
    case "text": {
      const face = fonts.faceFor(shape.face);
      const characters = Array.from(shape.text);

      out.fillColor(shape.color);
      out.beginText();
      out.font(face.resource, shape.emUnits);
      out.characterSpacing(0);

      // A metafile states where each character starts, so each is written at its
      // own place rather than being left to the face's advances. One entry alone
      // is the recording leaving the spacing to the face, and then the whole
      // string goes down as one.
      if (shape.xUnits.length < characters.length) {
        out.textMatrix([1, 0, 0, -1, shape.xUnits[0] ?? 0, shape.baselineUnits]);
        out.showGlyphs(face.glyphsFor(shape.text));
      } else {
        characters.forEach((character, index) => {
          out.textMatrix([1, 0, 0, -1, shape.xUnits[index] ?? 0, shape.baselineUnits]);
          out.showGlyphs(face.glyphsFor(character));
        });
      }
      out.endText();
      return;
    }
  }
}
