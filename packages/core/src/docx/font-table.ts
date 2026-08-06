import { partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import { attribute, childrenNamed } from "./xml.js";

export const FONT_TABLE_PART = "word/fontTable.xml";

// The shape of a face, which is all a stand-in can be asked to match once the
// face itself is not there to measure: a serif default for a serif name keeps a
// page recognisable where the right widths are already lost.
export type FaceShape = "sans-serif" | "serif" | "monospace";

// The classification bytes Word writes for every face a document names, kept for
// exactly this: choosing a substitute of the right shape on a machine without the
// face. The same PANOSE bytes are read out of a font file's own OS/2 table in
// `font-file.ts`; here they come from the document, which is the only place left
// to look when there is no file to read them from.
const LATIN_TEXT_FAMILY = 2;
const FIRST_SANS_SERIF_STYLE = 11;
const LAST_SANS_SERIF_STYLE = 15;
const MONOSPACED_PROPORTION = 9;

// The PANOSE value is ten bytes written as twenty hex digits. Word pads a face it
// knows nothing about with zeroes, which classify nothing.
function panoseBytes(value: string | undefined): readonly number[] | null {
  if (value === undefined || value.length < 20) return null;
  const bytes: number[] = [];
  for (let at = 0; at < 20; at += 2) {
    const byte = Number.parseInt(value.slice(at, at + 2), 16);
    if (Number.isNaN(byte)) return null;
    bytes.push(byte);
  }
  return bytes;
}

// Style bytes 0 and 1 are "any" and "no fit", which classify nothing; 2 to 10
// are the serif cuts and 11 to 15 the sans ones.
function shapeOfPanose(bytes: readonly number[]): FaceShape | null {
  if (bytes[0] !== LATIN_TEXT_FAMILY) return null;
  if (bytes[3] === MONOSPACED_PROPORTION) return "monospace";
  const serifStyle = bytes[1] ?? 0;
  if (serifStyle >= FIRST_SANS_SERIF_STYLE && serifStyle <= LAST_SANS_SERIF_STYLE)
    return "sans-serif";
  return serifStyle >= 2 ? "serif" : null;
}

// The coarser signal, kept for a face whose PANOSE says nothing: `modern` is the
// fixed-pitch family Courier New declares, `swiss` the sans one and `roman` the
// serif one. `script`, `decorative` and `auto` say nothing about shape.
function shapeOfFamily(family: string | undefined): FaceShape | null {
  if (family === "swiss") return "sans-serif";
  if (family === "roman") return "serif";
  if (family === "modern") return "monospace";
  return null;
}

const normalise = (name: string): string => name.trim().toLowerCase();

/**
 * What shape each face the document names draws its letters in, read from the
 * font table Word writes beside the document. The table exists for machines
 * without the faces: Word substitutes off these bytes, and so does the
 * best-effort resolution here. Keyed by the lowercased name, as
 * `lookupFontMetrics` keys its own lookups.
 *
 * A document without the part, which nothing obliges a producer to write,
 * classifies nothing and comes back empty.
 */
export function readFaceShapes(pkg: DocxPackage): ReadonlyMap<string, FaceShape> {
  if (!pkg.parts.has(FONT_TABLE_PART)) return new Map();

  const shapes = new Map<string, FaceShape>();
  for (const font of childrenNamed(partXml(pkg, FONT_TABLE_PART), W_NS, "font")) {
    const name = attribute(font, W_NS, "name");
    if (name === undefined) continue;

    const panose = childrenNamed(font, W_NS, "panose1")
      .map((each) => panoseBytes(attribute(each, W_NS, "val")))
      .find((bytes) => bytes !== null);
    const family = childrenNamed(font, W_NS, "family")
      .map((each) => attribute(each, W_NS, "val"))
      .find((value) => value !== undefined);

    const shape = (panose ? shapeOfPanose(panose) : null) ?? shapeOfFamily(family);
    if (shape !== null) shapes.set(normalise(name), shape);
  }
  return shapes;
}
