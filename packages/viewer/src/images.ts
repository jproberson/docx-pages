import {
  readMetafilePicture,
  type DocxPackage,
  type MetafilePicture,
  type MetricsResolver,
} from "@onepager/core";

const MEDIA_TYPES: ReadonlyMap<string, string> = new Map([
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

// What a browser has no decoder for and this reads itself instead. A metafile is
// a recording of the drawing rather than a picture of it, so it is played into
// shapes and drawn beside the rest of the page.
const METAFILE = "emf";

const extensionOf = (part: string): string => {
  const dot = part.lastIndexOf(".");
  return dot === -1 ? "" : part.slice(dot + 1).toLowerCase();
};

// btoa takes a binary string, and spreading a whole image at once overflows the
// argument limit, so the bytes go across in chunks.
const CHUNK = 0x8000;

function toBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let at = 0; at < bytes.length; at += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(at, at + CHUNK));
  }
  return btoa(binary);
}

export function imageDataUrl(pkg: DocxPackage, part: string): string | undefined {
  const bytes = pkg.parts.get(part);
  const mediaType = MEDIA_TYPES.get(extensionOf(part));
  if (bytes === undefined || mediaType === undefined) return undefined;
  return `data:${mediaType};base64,${toBase64(bytes)}`;
}

export type DrawableImage =
  | { readonly kind: "bitmap"; readonly url: string }
  | { readonly kind: "metafile"; readonly picture: MetafilePicture };

export type ImageResolver = (part: string) => DrawableImage | undefined;

function drawableImage(
  pkg: DocxPackage,
  part: string,
  metricsFor: MetricsResolver,
): DrawableImage | undefined {
  if (extensionOf(part) === METAFILE) {
    const bytes = pkg.parts.get(part);
    const picture = bytes === undefined ? null : readMetafilePicture(bytes, metricsFor);
    return picture === null ? undefined : { kind: "metafile", picture };
  }

  const url = imageDataUrl(pkg, part);
  return url === undefined ? undefined : { kind: "bitmap", url };
}

// Encoding the same picture once per render would be wasteful, and playing the
// same metafile once per render more so; a document's images do not change once it
// is open.
export function imageResolver(pkg: DocxPackage, metricsFor: MetricsResolver): ImageResolver {
  const cache = new Map<string, DrawableImage | undefined>();
  return (part) => {
    if (!cache.has(part)) cache.set(part, drawableImage(pkg, part, metricsFor));
    return cache.get(part);
  };
}
