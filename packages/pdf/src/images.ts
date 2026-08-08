import {
  pictureExtension,
  readMetafilePicture,
  METAFILE_EXTENSION,
  type CropInsets,
  type MetafilePicture,
  type MetafileRect,
  type MetafileShape,
  type MetricsResolver,
} from "@docx-pages/core";

import { bottomOf, type PdfPage } from "./coordinates.js";
import type { Content } from "./content.js";
import type { PdfFonts } from "./fonts.js";
import {
  pdfDictionary,
  pdfName,
  pdfNumber,
  pdfStream,
  type PdfObjects,
  type PdfReference,
  type PdfValue,
} from "./objects.js";
import type { ObjectDrawable } from "./objects-paint.js";

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

  return {
    draw: (out, page, at, part, crop) => {
      if (pictureExtension(part) === METAFILE_EXTENSION) {
        const picture = metafileFor(part);
        if (picture === null) return false;
        playMetafile(out, page, at, picture, crop, options.fonts);
        return true;
      }

      const bitmap = bitmapFor(part);
      if (bitmap === null) return false;
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

// An image draws into the unit square, so the transform that places it is its
// whole geometry.
//
// A source rectangle hides a fraction of each edge, so the whole bitmap is larger
// than the placed rectangle by exactly that much and is drawn behind it, cut to
// the frame. The viewer does the same with an oversized `img` inside an
// `overflow: hidden` box.
function drawBitmap(
  out: Content,
  page: PdfPage,
  at: ObjectDrawable,
  resource: string,
  crop: CropInsets,
): void {
  const widthPt = at.widthPt / Math.max(1 - crop.left - crop.right, Number.EPSILON);
  const heightPt = at.heightPt / Math.max(1 - crop.top - crop.bottom, Number.EPSILON);
  const leftPt = at.leftPt - crop.left * widthPt;
  const topPt = at.topPt - crop.top * heightPt;

  out.save();
  out.rectangle(at.leftPt, bottomOf(page, at.topPt, at.heightPt), at.widthPt, at.heightPt);
  out.clip();
  out.transform([widthPt, 0, 0, heightPt, leftPt, bottomOf(page, topPt, heightPt)]);
  out.drawObject(resource);
  out.restore();
}

// A metafile's pen hangs its line off the cell whose corner the line runs through,
// where a stroke is centred on the line it follows, so the line moves half a unit
// down and right to cover the same cells.
const PEN_OFFSET = 0.5;

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
