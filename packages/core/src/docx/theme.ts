import { partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import { A_NS } from "./styles.js";
import { attribute, childrenNamed, firstNamed, type XmlElement } from "./xml.js";

export const THEME_PART = "word/theme/theme1.xml";
const SETTINGS_PART = "word/settings.xml";

// The slots a theme actually holds. Word's own bg1, tx1, bg2 and tx2 are not
// among them: each of those is a name for whichever of these the document's
// colour mapping points it at.
const THEME_SLOTS = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

// The mapping is written with Word's own names on both sides: w:bg1="light1"
// against a theme that calls the same slot lt1.
const MAPPED_NAMES: Readonly<Record<string, string>> = {
  bg1: "bg1",
  tx1: "t1",
  bg2: "bg2",
  tx2: "t2",
};

const MAPPING_TARGETS: Readonly<Record<string, string>> = {
  light1: "lt1",
  dark1: "dk1",
  light2: "lt2",
  dark2: "dk2",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
};

// Word's presets run to well over a hundred names. These documents ask for two,
// and a name that is not here resolves to nothing rather than to a wrong colour.
const PRESET_COLORS: Readonly<Record<string, string>> = {
  white: "FFFFFF",
  black: "000000",
};

export type Theme = {
  readonly scheme: ReadonlyMap<string, string>;
  readonly mapping: ReadonlyMap<string, string>;
};

export const NO_THEME: Theme = { scheme: new Map(), mapping: new Map() };

// A colour as the file writes it: a literal, or a slot the theme fills in, under
// the luminance transform that turns one white into every grey in these documents.
export type ColorReference = {
  readonly base:
    | { readonly kind: "literal"; readonly hex: string }
    | { readonly kind: "scheme"; readonly slot: string };
  readonly luminanceScale: number;
  readonly luminanceOffset: number;
};

const PERCENT_UNITS = 100000;

function fraction(parent: XmlElement, name: string, fallback: number): number {
  const element = firstNamed(parent, A_NS, name);
  if (element === null) return fallback;
  const raw = attribute(element, "", "val");
  const value = raw === undefined ? Number.NaN : Number(raw);
  return Number.isFinite(value) ? value / PERCENT_UNITS : fallback;
}

const HEX = /^[0-9A-Fa-f]{6}$/;

function baseOf(color: XmlElement): ColorReference["base"] | null {
  const value = attribute(color, "", "val");
  switch (color.name) {
    case "srgbClr":
      return value !== undefined && HEX.test(value) ? { kind: "literal", hex: value } : null;
    case "schemeClr":
      return value === undefined ? null : { kind: "scheme", slot: value };
    // A system colour carries the value the producer last saw it resolve to,
    // which is the only one readable away from the machine that wrote it.
    case "sysClr": {
      const last = attribute(color, "", "lastClr");
      return last !== undefined && HEX.test(last) ? { kind: "literal", hex: last } : null;
    }
    case "prstClr": {
      const preset = value === undefined ? undefined : PRESET_COLORS[value];
      return preset === undefined ? null : { kind: "literal", hex: preset };
    }
    default:
      return null;
  }
}

/**
 * The colour a fill or an outline is given, which is the first colour element it
 * holds. An element that names no readable colour reads as none at all.
 *
 * **A colour stated fully transparent names none.** `a:alpha` is a share of
 * opacity, and a fill or an outline stated at nought is one Word puts no ink down
 * for: a corpus document draws two full-width rectangles whose line is
 * `<a:srgbClr val="000000"><a:alpha val="0"/></a:srgbClr>`, and Word's own pdf
 * draws neither of them where this drew a black hairline across the head of the
 * page and another near its foot.
 *
 * **Nothing here draws a colour half way, and the corpus says it need not.** Read
 * on 2026-08-18 over the flow parts of all 718: 13 of them state an alpha of
 * nought, 29 times inside an `a:ln` and 45 in a fill. Of every other alpha those
 * parts state, 124 are the full 100000 and **two** are anything else at all. So
 * nought is answered for and every other alpha is drawn opaque, exactly as before.
 */
export function readColorReference(container: XmlElement): ColorReference | null {
  for (const child of container.children) {
    if (child.namespace !== A_NS) continue;
    const base = baseOf(child);
    if (base === null) continue;
    if (fraction(child, "alpha", 1) === 0) return null;
    return {
      base,
      luminanceScale: fraction(child, "lumMod", 1),
      luminanceOffset: fraction(child, "lumOff", 0),
    };
  }
  return null;
}

const clamp = (value: number): number => Math.min(1, Math.max(0, value));

type Hsl = { readonly hue: number; readonly saturation: number; readonly luminance: number };

function toHsl(hex: string): Hsl {
  const channels = [0, 2, 4].map((at) => Number.parseInt(hex.slice(at, at + 2), 16) / 255);
  const [red = 0, green = 0, blue = 0] = channels;
  const high = Math.max(red, green, blue);
  const low = Math.min(red, green, blue);
  const luminance = (high + low) / 2;
  if (high === low) return { hue: 0, saturation: 0, luminance };

  const span = high - low;
  const saturation = luminance > 0.5 ? span / (2 - high - low) : span / (high + low);
  const hue =
    high === red
      ? (green - blue) / span + (green < blue ? 6 : 0)
      : high === green
        ? (blue - red) / span + 2
        : (red - green) / span + 4;
  return { hue: hue / 6, saturation, luminance };
}

function channel(from: number, to: number, at: number): number {
  const turn = at < 0 ? at + 1 : at > 1 ? at - 1 : at;
  if (turn < 1 / 6) return from + (to - from) * 6 * turn;
  if (turn < 1 / 2) return to;
  if (turn < 2 / 3) return from + (to - from) * (2 / 3 - turn) * 6;
  return from;
}

function toHex({ hue, saturation, luminance }: Hsl): string {
  const to =
    luminance < 0.5
      ? luminance * (1 + saturation)
      : luminance + saturation - luminance * saturation;
  const from = 2 * luminance - to;
  const parts =
    saturation === 0
      ? [luminance, luminance, luminance]
      : [channel(from, to, hue + 1 / 3), channel(from, to, hue), channel(from, to, hue - 1 / 3)];
  return parts
    .map((part) =>
      Math.round(clamp(part) * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase();
}

// lumMod scales a colour's luminance and lumOff shifts it, which is how one white
// theme slot becomes both of the greys these documents are drawn with.
export function themeColor(theme: Theme, reference: ColorReference): string | null {
  const { base } = reference;
  const hex =
    base.kind === "literal"
      ? base.hex
      : (theme.scheme.get(theme.mapping.get(base.slot) ?? base.slot) ?? null);
  if (hex === null) return null;

  const hsl = toHsl(hex);
  const luminance = clamp(hsl.luminance * reference.luminanceScale + reference.luminanceOffset);
  return `#${toHex({ ...hsl, luminance })}`;
}

function readScheme(pkg: DocxPackage): ReadonlyMap<string, string> {
  const scheme = new Map<string, string>();
  if (!pkg.parts.has(THEME_PART)) return scheme;

  const elements = firstNamed(partXml(pkg, THEME_PART), A_NS, "themeElements");
  const colors = elements === null ? null : firstNamed(elements, A_NS, "clrScheme");
  if (colors === null) return scheme;

  for (const slot of THEME_SLOTS) {
    const holder = firstNamed(colors, A_NS, slot);
    const reference = holder === null ? null : readColorReference(holder);
    if (reference !== null && reference.base.kind === "literal") {
      scheme.set(slot, reference.base.hex);
    }
  }
  return scheme;
}

function readMapping(pkg: DocxPackage): ReadonlyMap<string, string> {
  const mapping = new Map<string, string>();
  if (!pkg.parts.has(SETTINGS_PART)) return mapping;

  const settings = childrenNamed(partXml(pkg, SETTINGS_PART), W_NS, "clrSchemeMapping");
  const [element] = settings;
  if (element === undefined) return mapping;

  for (const [slot, name] of Object.entries(MAPPED_NAMES)) {
    const target = attribute(element, W_NS, name);
    if (target === undefined) continue;
    mapping.set(slot, MAPPING_TARGETS[target] ?? target);
  }
  return mapping;
}

export const readTheme = (pkg: DocxPackage): Theme => ({
  scheme: readScheme(pkg),
  mapping: readMapping(pkg),
});
