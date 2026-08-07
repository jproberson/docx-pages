import { partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import { attribute, childrenNamed, firstNamed } from "./xml.js";

export const SETTINGS_PART = "word/settings.xml";

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

// Whether a document is one of the old ones, which the two rules below are the
// whole of besides the rounding above.
const legacy = roundsAnchorsToTwips;

// An object asking for the largest side with the same room either side of it puts
// the text on the left in a document declaring 15 and on the right in one
// declaring nothing. Measured with the same body written both ways: an object
// exactly centred in the column, at two widths and stated both as an offset and as
// an alignment, and every one of the four flipped sides with the setting.
export const takesTheRightOnEqualSides = legacy;

// A table's indent is measured to the leading edge of the text in its first cell
// in a document declaring no compatibility mode, and to the table's own edge in
// one declaring 15. So an old document's table hangs its first column's margin and
// border outside the indent, and outside the page margin altogether where it asks
// for no indent at all.
//
// Measured with the same body written both ways, at indents under, over and equal
// to the cell margin they stand against, and again with no margin at all: in the
// old form the text landed on the indent to the twip every time and the table's
// edge moved to put it there, whatever margin the cell asked for. Where two rows
// ask for different margins the table still has one edge, and it is the first
// row's cell that decides it.
export const measuresTheIndentToTheText = legacy;

// A document declaring no compatibility mode does not keep text off the right of
// an object wrapped on its left: the line takes the run of free space beside the
// object as though the wrap named both sides. Measured with an object flush to the
// left margin, and again with one wide enough that the free run past it starts
// beyond the middle of the column, which both put the line beside it. The same
// document declaring 15 drops the line past the object instead.
//
// A wrap naming the right is honoured either way round: an object flush to the
// right margin drops the line past itself in both.
export const honoursAWrapOnTheLeft = (settings: DocumentSettings): boolean => !legacy(settings);

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
