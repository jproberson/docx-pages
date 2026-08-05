import {
  pictureExtension,
  readMetafilePicture,
  METAFILE_EXTENSION,
  PICTURE_MEDIA_TYPES,
  type DocxPackage,
  type MetafilePicture,
  type MetricsResolver,
} from "@docx-pages/core";

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
  const mediaType = PICTURE_MEDIA_TYPES.get(pictureExtension(part));
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
  if (pictureExtension(part) === METAFILE_EXTENSION) {
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
