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

// Whether a part holding a picture is one anything here can draw. A WMF is drawn
// only where the bitmap it blits can be read out of it, which is what its own bytes
// say and not what its name does, so one is asked for where they are to hand.
export function drawablePicture(part: string, bytes?: Uint8Array): boolean {
  const extension = pictureExtension(part);
  if (extension === OLD_METAFILE_EXTENSION) {
    return bytes !== undefined && readMetafileBitmap(bytes) !== null;
  }
  return PICTURE_MEDIA_TYPES.has(extension) || extension === METAFILE_EXTENSION;
}
