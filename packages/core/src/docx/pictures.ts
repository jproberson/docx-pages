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

export function pictureExtension(part: string): string {
  const dot = part.lastIndexOf(".");
  return dot === -1 ? "" : part.slice(dot + 1).toLowerCase();
}

// Whether a part holding a picture is one anything here can draw. Word writes a
// WMF where it writes an EMF, and nothing reads that one, so a document holding
// one draws a mark in its place.
export function drawablePicture(part: string): boolean {
  const extension = pictureExtension(part);
  return PICTURE_MEDIA_TYPES.has(extension) || extension === METAFILE_EXTENSION;
}
