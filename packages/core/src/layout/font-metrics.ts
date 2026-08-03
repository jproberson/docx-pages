export type FontMetrics = {
  readonly unitsPerEm: number;
  readonly ascender: number;
  readonly descender: number;
  readonly lineGap: number;
};

export type MetricsLookup =
  | {
      readonly kind: "found";
      readonly source: "builtin" | "supplied";
      readonly metrics: FontMetrics;
    }
  | { readonly kind: "missing"; readonly fontName: string };

export const lineHeightPt = (metrics: FontMetrics, fontSizePt: number): number =>
  (fontSizePt * (metrics.ascender - metrics.descender + metrics.lineGap)) / metrics.unitsPerEm;

export const ascentPt = (metrics: FontMetrics, fontSizePt: number): number =>
  (fontSizePt * metrics.ascender) / metrics.unitsPerEm;

const BUILTIN: ReadonlyMap<string, FontMetrics> = new Map([
  ["arial", { unitsPerEm: 2048, ascender: 1854, descender: -434, lineGap: 67 }],
  ["calibri", { unitsPerEm: 2048, ascender: 1950, descender: -550, lineGap: 0 }],
  ["times new roman", { unitsPerEm: 2048, ascender: 1825, descender: -443, lineGap: 87 }],
  ["courier new", { unitsPerEm: 2048, ascender: 1705, descender: -615, lineGap: 0 }],
  ["georgia", { unitsPerEm: 2048, ascender: 1878, descender: -449, lineGap: 0 }],
  ["verdana", { unitsPerEm: 2048, ascender: 2059, descender: -430, lineGap: 0 }],
  ["trebuchet ms", { unitsPerEm: 2048, ascender: 1923, descender: -455, lineGap: 0 }],
  ["tahoma", { unitsPerEm: 2048, ascender: 2049, descender: -423, lineGap: 0 }],
  ["comic sans ms", { unitsPerEm: 2048, ascender: 2257, descender: -597, lineGap: 0 }],
  ["impact", { unitsPerEm: 2048, ascender: 2066, descender: -432, lineGap: 0 }],
]);

const normalise = (fontName: string): string => fontName.trim().toLowerCase();

export function lookupFontMetrics(
  fontName: string,
  supplied?: ReadonlyMap<string, FontMetrics>,
): MetricsLookup {
  const key = normalise(fontName);

  for (const [name, metrics] of supplied ?? []) {
    if (normalise(name) === key) return { kind: "found", source: "supplied", metrics };
  }

  const builtin = BUILTIN.get(key);
  if (builtin !== undefined) return { kind: "found", source: "builtin", metrics: builtin };

  return { kind: "missing", fontName };
}
