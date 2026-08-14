import type { MetricsResolver } from "../layout/lines.js";
import { readMetafilePicture } from "../metafile/picture.js";
import { readMetafileBitmap } from "../metafile/wmf.js";

// The picture formats this project can put on a page, which is what a browser has
// a decoder for plus the one this project decodes itself.
//
// A metafile holds a recording of the drawing rather than a picture of it, so it
// is played into shapes and drawn beside the rest of the page. Everything else is
// handed over as it stands, under the type named here.
export const PICTURE_MEDIA_TYPES: ReadonlyMap<string, string> = new Map([
  ["png", "image/png"],
  ["jpg", "image/jpeg"],
  ["jpeg", "image/jpeg"],
  ["gif", "image/gif"],
  ["bmp", "image/bmp"],
  ["tif", "image/tiff"],
  ["tiff", "image/tiff"],
  ["svg", "image/svg+xml"],
  ["webp", "image/webp"],
]);

export const METAFILE_EXTENSION = "emf";

// The older metafile, which is read as the bitmap it blits rather than played.
export const OLD_METAFILE_EXTENSION = "wmf";

export function pictureExtension(part: string): string {
  const dot = part.lastIndexOf(".");
  return dot === -1 ? "" : part.slice(dot + 1).toLowerCase();
}

// Whether a part holding a picture is one anything here can draw. Neither metafile
// answers by its name: a WMF is drawn where the bitmap it blits can be read out of
// it, and an EMF where the records it holds play, so the bytes are asked for where
// they are to hand.
//
// **Playing an EMF needs the faces**, since one selecting a face nothing can supply
// metrics for is refused whole at its first run of text, and the ones met in the
// wild write text. So a caller with no resolver cannot ask, and the picture is taken
// on trust rather than named on a guess: put a resolver that finds nothing to the
// corpus and most of the metafiles it names are metafiles that play. What is left
// named with the machine's own faces is one that refuses on its own records, which
// is a metafile blitting a bitmap.
export function drawablePicture(
  part: string,
  bytes?: Uint8Array,
  metricsFor?: MetricsResolver,
): boolean {
  const extension = pictureExtension(part);
  if (extension === OLD_METAFILE_EXTENSION) {
    return bytes !== undefined && readMetafileBitmap(bytes) !== null;
  }
  if (extension === METAFILE_EXTENSION) {
    if (metricsFor === undefined) return true;
    return bytes !== undefined && readMetafilePicture(bytes, metricsFor) !== null;
  }
  return PICTURE_MEDIA_TYPES.has(extension);
}
