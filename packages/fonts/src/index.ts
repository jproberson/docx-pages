import { readFontFile, type FaceDefaults, type SuppliedFace } from "@docx-pages/core";

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
  // The name documents write, whose widths this face matches.
  readonly twinOf: string;
};

const family = (
  name: string,
  twinOf: string,
  files: readonly [string, string, string, string],
): readonly PackFace[] => [
  { name, bold: false, italic: false, file: files[0], twinOf },
  { name, bold: true, italic: false, file: files[1], twinOf },
  { name, bold: false, italic: true, file: files[2], twinOf },
  { name, bold: true, italic: true, file: files[3], twinOf },
];

export const PACK_FACES: readonly PackFace[] = [
  ...family("Carlito", "Calibri", [
    "Carlito-Regular.ttf",
    "Carlito-Bold.ttf",
    "Carlito-Italic.ttf",
    "Carlito-BoldItalic.ttf",
  ]),
  ...family("Caladea", "Cambria", [
    "Caladea-Regular.ttf",
    "Caladea-Bold.ttf",
    "Caladea-Italic.ttf",
    "Caladea-BoldItalic.ttf",
  ]),
  ...family("Liberation Sans", "Arial", [
    "LiberationSans-Regular.ttf",
    "LiberationSans-Bold.ttf",
    "LiberationSans-Italic.ttf",
    "LiberationSans-BoldItalic.ttf",
  ]),
  ...family("Liberation Serif", "Times New Roman", [
    "LiberationSerif-Regular.ttf",
    "LiberationSerif-Bold.ttf",
    "LiberationSerif-Italic.ttf",
    "LiberationSerif-BoldItalic.ttf",
  ]),
  ...family("Liberation Mono", "Courier New", [
    "LiberationMono-Regular.ttf",
    "LiberationMono-Bold.ttf",
    "LiberationMono-Italic.ttf",
    "LiberationMono-BoldItalic.ttf",
  ]),
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

// The shape defaults follow the faces Word itself reaches for: Arial for a sans,
// Times New Roman for a serif (see the fallback faces core names), each answered
// by its twin. The last resort is Caladea because Word's own last resort for a
// name it cannot place is Cambria, whatever the name suggested.
export const SANS_SERIF_DEFAULT = "Liberation Sans";
export const SERIF_DEFAULT = "Liberation Serif";
export const MONOSPACE_DEFAULT = "Liberation Mono";
export const LAST_RESORT_DEFAULT = "Caladea";

export const fontUrl = (face: PackFace): URL => new URL(`../fonts/${face.file}`, import.meta.url);

export type ReadBytes = (url: URL) => Promise<Uint8Array>;

// The reader a browser needs; a runtime whose fetch cannot reach the pack's own
// files, node being the one, goes through `./node` instead.
const overFetch: ReadBytes = async (url) => {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`the font at ${url.href} came back ${String(response.status)}`);
  }
  return new Uint8Array(await response.arrayBuffer());
};

// The defaults `bestEffortMetrics` asks for, read out of the pack's own files so
// that each face knows its widths, its shape and its missing-glyph advance.
export async function defaultFaces(read: ReadBytes = overFetch): Promise<FaceDefaults> {
  const faces = await Promise.all(
    PACK_FACES.map(async (each): Promise<SuppliedFace> => {
      const found = readFontFile(await read(fontUrl(each)));
      return {
        name: each.name,
        bold: each.bold,
        italic: each.italic,
        metrics: found.metrics,
        advances: found.advances,
        sansSerif: found.sansSerif,
      };
    }),
  );

  return {
    faces,
    twins: METRIC_TWINS,
    sansSerif: SANS_SERIF_DEFAULT,
    serif: SERIF_DEFAULT,
    monospace: MONOSPACE_DEFAULT,
    lastResort: LAST_RESORT_DEFAULT,
  };
}
