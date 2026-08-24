import { readSuppliedFace, type FaceDefaults, type SuppliedFace } from "@docx-pages/core";

// Faces that may be shipped where the ones documents actually name may not.
// Each family here was drawn to match the advance widths of a face Word ships,
// glyph for glyph, so a line laid out on the twin breaks where the named face
// would have broken it. Checked with this project's own reader on 2026-08-06:
// identical advances over printable ASCII and the common punctuation beyond it,
// and identical vertical metrics, for every pair below. (One nuance: Liberation
// Serif carries the 87-unit line gap Word's layout gives Times New Roman, which
// the times.ttf some machines hold does not state.)
export type PackFace = {
  readonly name: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly file: string;
  // The name documents write, whose widths this face matches. A face carried
  // for its looks rather than its widths twins nothing.
  readonly twinOf?: string;
  // What the face is, where its own PANOSE bytes misstate it: Open Sans's say
  // it is not a sans, and the borrowing rule would believe them.
  readonly sansSerif?: boolean;
};

const family = (
  name: string,
  files: readonly [string, string, string, string],
  rest: Partial<PackFace> = {},
): readonly PackFace[] => [
  { name, bold: false, italic: false, file: files[0], ...rest },
  { name, bold: true, italic: false, file: files[1], ...rest },
  { name, bold: false, italic: true, file: files[2], ...rest },
  { name, bold: true, italic: true, file: files[3], ...rest },
];

export const PACK_FACES: readonly PackFace[] = [
  ...family(
    "Carlito",
    ["Carlito-Regular.ttf", "Carlito-Bold.ttf", "Carlito-Italic.ttf", "Carlito-BoldItalic.ttf"],
    { twinOf: "Calibri" },
  ),
  ...family(
    "Caladea",
    ["Caladea-Regular.ttf", "Caladea-Bold.ttf", "Caladea-Italic.ttf", "Caladea-BoldItalic.ttf"],
    { twinOf: "Cambria" },
  ),
  ...family(
    "Liberation Sans",
    [
      "LiberationSans-Regular.ttf",
      "LiberationSans-Bold.ttf",
      "LiberationSans-Italic.ttf",
      "LiberationSans-BoldItalic.ttf",
    ],
    { twinOf: "Arial" },
  ),
  ...family(
    "Liberation Serif",
    [
      "LiberationSerif-Regular.ttf",
      "LiberationSerif-Bold.ttf",
      "LiberationSerif-Italic.ttf",
      "LiberationSerif-BoldItalic.ttf",
    ],
    { twinOf: "Times New Roman" },
  ),
  ...family(
    "Liberation Mono",
    [
      "LiberationMono-Regular.ttf",
      "LiberationMono-Bold.ttf",
      "LiberationMono-Italic.ttf",
      "LiberationMono-BoldItalic.ttf",
    ],
    { twinOf: "Courier New" },
  ),
  ...family(
    "Open Sans",
    ["OpenSans-Regular.ttf", "OpenSans-Bold.ttf", "OpenSans-Italic.ttf", "OpenSans-BoldItalic.ttf"],
    { sansSerif: true },
  ),
];

// Which pack face answers for a name, keyed as the resolver keys its lookups.
// Helvetica is here beside Arial because Arial was drawn to Helvetica's widths,
// so the same twin serves both names.
export const METRIC_TWINS: Readonly<Record<string, string>> = {
  calibri: "Carlito",
  cambria: "Caladea",
  arial: "Liberation Sans",
  helvetica: "Liberation Sans",
  "times new roman": "Liberation Serif",
  "courier new": "Liberation Mono",
  courier: "Liberation Mono",
};

// A name with no twin gets no right widths from anyone, so the shape defaults
// are chosen for how the page reads rather than for any measurement. Open Sans
// answers for an unknown sans because it reads like the faces such documents
// were set in, where an Arial clone reads like a different decade; the serif
// and monospace defaults are the twins, which at least stand near the faces
// Word substitutes through. The last resort is Caladea because Word's own last
// resort for a name it cannot place is Cambria, whatever the name suggested.
export const SANS_SERIF_DEFAULT = "Open Sans";
export const SERIF_DEFAULT = "Liberation Serif";
export const MONOSPACE_DEFAULT = "Liberation Mono";
export const LAST_RESORT_DEFAULT = "Caladea";

export const fontUrl = (face: PackFace): URL => new URL(`../fonts/${face.file}`, import.meta.url);

export type ReadBytes = (url: URL) => Promise<Uint8Array>;

// The reader a browser needs; a runtime whose fetch cannot reach the pack's own
// files, node being the one, goes through `./node` instead.
const overFetch: ReadBytes = async (url) => {
  // A dev server that prebundles its dependencies moves this module into a deps
  // cache, away from the files it resolves beside itself. Vite's is the one met so
  // far, and the way out is configuration, so say so.
  if (url.href.includes("/.vite/")) {
    throw new Error(
      `the font pack was prebundled away from its own files (${url.href}); add optimizeDeps: { exclude: ["@docx-pages/viewer", "@docx-pages/fonts"] } to vite.config`,
    );
  }
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`the font at ${url.href} came back ${String(response.status)}`);
  }
  return new Uint8Array(await response.arrayBuffer());
};

// One face's own file, under the name the pack carries it as.
export type PackBytes = {
  readonly name: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly bytes: Uint8Array;
};

export type FontPack = {
  readonly defaults: FaceDefaults;
  readonly bytes: readonly PackBytes[];
};

/**
 * The pack read once: the defaults `bestEffortMetrics` asks for, so that each face
 * knows its widths, its shape and its missing-glyph advance, and beside them the
 * very bytes those were read out of.
 *
 * The bytes are here because a page is not only measured. A browser has to be
 * offered the face it was measured in, and a pdf carries the faces it draws in, so
 * whoever draws the page needs the file and not only what was read out of it.
 */
export async function readPack(read: ReadBytes = overFetch): Promise<FontPack> {
  const pack = await Promise.all(
    PACK_FACES.map(async (each) => {
      const bytes = await read(fontUrl(each));
      const face: SuppliedFace = readSuppliedFace(
        bytes,
        { name: each.name, bold: each.bold, italic: each.italic },
        each.sansSerif === undefined ? {} : { sansSerif: each.sansSerif },
      );
      return { face, bytes: { name: each.name, bold: each.bold, italic: each.italic, bytes } };
    }),
  );

  return {
    defaults: {
      faces: pack.map((each) => each.face),
      twins: METRIC_TWINS,
      sansSerif: SANS_SERIF_DEFAULT,
      serif: SERIF_DEFAULT,
      monospace: MONOSPACE_DEFAULT,
      lastResort: LAST_RESORT_DEFAULT,
    },
    bytes: pack.map((each) => each.bytes),
  };
}

// What a caller that only lays out asks for, which is the pack without its files.
export const defaultFaces = async (read: ReadBytes = overFetch): Promise<FaceDefaults> =>
  (await readPack(read)).defaults;
