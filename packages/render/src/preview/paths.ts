import { join, normalize, sep } from "node:path";

// A root without its separator is a prefix of its own siblings, so a check for one
// lets `samples-backup` out of `samples`. The url parser collapses a plain `..`
// before we see it, but an encoded one survives to here.
export const pathOf = (url: string, root: string, start: string): string => {
  const asked = decodeURIComponent(new URL(url, "http://localhost").pathname);
  const path = normalize(join(root, asked === "/" ? start : asked));
  return path === root || path.startsWith(root + sep) ? path : root;
};
