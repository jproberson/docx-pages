import type { DocxPackage } from "@onepager/core";

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

const mediaTypeOf = (part: string): string | undefined => {
  const dot = part.lastIndexOf(".");
  return dot === -1 ? undefined : MEDIA_TYPES.get(part.slice(dot + 1).toLowerCase());
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
  const mediaType = mediaTypeOf(part);
  if (bytes === undefined || mediaType === undefined) return undefined;
  return `data:${mediaType};base64,${toBase64(bytes)}`;
}

export type ImageResolver = (part: string) => string | undefined;

// Encoding the same picture once per render would be wasteful; a document's
// images do not change once it is open.
export function imageResolver(pkg: DocxPackage): ImageResolver {
  const cache = new Map<string, string | undefined>();
  return (part) => {
    if (!cache.has(part)) cache.set(part, imageDataUrl(pkg, part));
    return cache.get(part);
  };
}
