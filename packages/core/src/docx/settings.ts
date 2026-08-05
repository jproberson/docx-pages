import { partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import { attribute, childrenNamed, firstNamed } from "./xml.js";

export const SETTINGS_PART = "word/settings.xml";

// What the document says about itself rather than about any of its content.
export type DocumentSettings = {
  // How far apart the stops a paragraph falls back on stand, which Word writes as
  // 720 twips and a document that says nothing keeps.
  readonly defaultTabStopTwips: number;
  // Which Word the document was written to be laid out by, and null for one that
  // never says: a file out of another word processor usually does not.
  readonly compatibilityMode: number | null;
};

export const DEFAULT_SETTINGS: DocumentSettings = {
  defaultTabStopTwips: 720,
  compatibilityMode: null,
};

// Word rounds a floating object's position to the whole twip in a document that
// declares no compatibility mode, and leaves it where the flow put it in one
// declaring 15. That rounding is worth this much: an object anchored at the top of
// its paragraph stands at the foot of the paragraph before it, so rounding down
// puts the object's wrap band over that paragraph's last line, which is then
// blocked and falls past the object like any other blocked line. Measured over 34
// documents in five faces at sizes from 10 to 24pt: whether the fraction of a twip
// is under a half called every one of them, and nothing else did.
export const roundsAnchorsToTwips = (settings: DocumentSettings): boolean =>
  settings.compatibilityMode === null || settings.compatibilityMode < 15;

const COMPATIBILITY_MODE = "compatibilityMode";

function numbered(raw: string | undefined): number | null {
  if (raw === undefined) return null;
  const value = Number(raw);
  return Number.isFinite(value) ? value : null;
}

export function readDocumentSettings(pkg: DocxPackage): DocumentSettings {
  if (!pkg.parts.has(SETTINGS_PART)) return DEFAULT_SETTINGS;

  const root = partXml(pkg, SETTINGS_PART);
  const stop = firstNamed(root, W_NS, "defaultTabStop");
  const compat = firstNamed(root, W_NS, "compat");
  const setting = compat === null ? [] : childrenNamed(compat, W_NS, "compatSetting");

  return {
    defaultTabStopTwips:
      numbered(stop === null ? undefined : attribute(stop, W_NS, "val")) ??
      DEFAULT_SETTINGS.defaultTabStopTwips,
    compatibilityMode: numbered(
      setting
        .filter((each) => attribute(each, W_NS, "name") === COMPATIBILITY_MODE)
        .map((each) => attribute(each, W_NS, "val"))
        .at(-1),
    ),
  };
}
