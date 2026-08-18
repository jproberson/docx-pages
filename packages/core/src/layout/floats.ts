import type { AnchorOrigin, AnchorPosition, FloatingAnchor } from "../docx/anchors.js";
import type {
  CropInsets,
  DrawingContent,
  DrawingFlip,
  PathCommand,
  ShapeGeometry,
  ShapePaint,
  TextBoxBody,
} from "../docx/drawing.js";
import type { SectionGeometry } from "../docx/section.js";
import { roundsAnchorsToTwips, DEFAULT_SETTINGS, type DocumentSettings } from "../docx/settings.js";
import { themeColor, type Theme } from "../docx/theme.js";
import type { PlacedTextBox } from "./text-boxes.js";
import { emuToPoints, twipsToPoints } from "./units.js";

// A shape's paint once the theme has answered for its colours.
export type PlacedPaint = {
  readonly fillColor: string | null;
  readonly outline: { readonly color: string; readonly widthPt: number } | null;
  readonly geometry: ShapeGeometry;
  // The outline a `custom` geometry is drawn with, in shares of its own box.
  readonly path: readonly PathCommand[] | null;
};

export const UNPAINTED: PlacedPaint = {
  fillColor: null,
  outline: null,
  geometry: "rectangle",
  path: null,
};

export type PlacedContent =
  | {
      readonly kind: "picture";
      readonly part: string;
      readonly crop: CropInsets;
      readonly paint: PlacedPaint;
    }
  | { readonly kind: "missing-picture"; readonly relationshipId: string }
  | {
      readonly kind: "text-box";
      readonly body: TextBoxBody;
      readonly text: PlacedTextBox | null;
      readonly paint: PlacedPaint;
    }
  | { readonly kind: "shape"; readonly paint: PlacedPaint }
  // A group's children keep the fractions of its box they stand in, so nothing
  // above the drawing has to know a group is one object rather than many.
  | { readonly kind: "group"; readonly children: readonly PlacedGroupChild[] }
  | { readonly kind: "unknown" };

export type PlacedGroupChild = {
  readonly leftFraction: number;
  readonly topFraction: number;
  readonly widthFraction: number;
  readonly heightFraction: number;
  readonly flip: DrawingFlip;
  readonly turnDegrees: number;
  readonly content: PlacedContent;
};

export type PlacedFloat = {
  readonly anchor: FloatingAnchor;
  readonly content: PlacedContent;
  // Where the object stands before it is turned, as an inline drawing's box is.
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly turnDegrees: number;
  readonly flip: DrawingFlip;
};

export type PartResolver = (relationshipId: string) => string | null;

export type FloatSize = {
  readonly widthPt: number;
  readonly heightPt: number;
};

export type PlaceFloatInput = {
  readonly anchor: FloatingAnchor;
  readonly page: SectionGeometry;
  readonly paragraphTopPt: number;
  readonly bodyTopPt: number;
  // Where the body's own text starts, which a margin-relative offset is measured
  // down from even for an object anchored in the header.
  readonly marginTopPt: number;
  readonly resolvePart: PartResolver;
  readonly theme: Theme;
  readonly settings?: DocumentSettings;
  // What the object turned out to be, for one that sizes itself to its content.
  // An aligned object lands on its own size, so this decides where it goes.
  readonly sizePt?: FloatSize;
  // The cell the anchoring paragraph stands in, or null outside a table.
  readonly cell?: CellFrame | null;
};

// How far across the page a table cell reaches. This is the cell itself and not
// the room its text has: the margin holding that text off the wall is no part of
// what an object anchored in the cell is measured from.
export type CellFrame = {
  readonly leftPt: number;
  readonly widthPt: number;
};

type Band = { readonly startPt: number; readonly extentPt: number };

/**
 * How far across the page an object anchored in a cell is measured, which is the
 * cell and not the page whatever origin the object names.
 *
 * **Measured on 2026-08-18 off Word's own pdf, on both of the origins the corpus
 * states.** `c8ca0c3c8292` and `2c1289b95c31` hold four pictures anchored
 * `relativeFrom="page"` in the two cells of a floating table's last row, and Word
 * draws them 340.80, 340.76, 456.16 and 456.18 right of where a page-relative
 * offset puts them: two numbers, one a cell, and 115.40 apart, which is the first
 * cell's own width off the table grid to a hundredth. `7eaa70746b70` holds three
 * groups anchored `relativeFrom="column"` in two cells of two rows, and Word draws
 * all three at their cell's left plus the stated offset, out by a constant 0.23
 * that is this project's own reading of where the table starts.
 *
 * Nothing states `margin` or `character` inside a cell, and nothing states a
 * horizontal `paragraph` or `line` there either, so the rule is written once for
 * every origin rather than split where no document could say which way.
 */
function horizontalBand(page: SectionGeometry, from: AnchorOrigin, cell: CellFrame | null): Band {
  if (cell !== null) return { startPt: cell.leftPt, extentPt: cell.widthPt };
  const left = twipsToPoints(page.margin.leftTwips);
  const right = twipsToPoints(page.margin.rightTwips);
  const width = twipsToPoints(page.widthTwips);
  switch (from) {
    case "page":
      return { startPt: 0, extentPt: width };
    case "margin":
    case "column":
    case "character":
      return { startPt: left, extentPt: width - left - right };
    case "paragraph":
    case "line":
      return { startPt: left, extentPt: width - left - right };
  }
}

// Measured against Word, which put a header picture 55.77pt above the margin at
// 13.3pt down the page: the top margin an offset is measured from is where the
// body's text begins, which a header deep enough to overrun the margin pushes
// down with it.
function verticalBand(input: PlaceFloatInput, from: AnchorOrigin): Band {
  const { page, paragraphTopPt, bodyTopPt, marginTopPt } = input;
  const bottom = twipsToPoints(page.margin.bottomTwips);
  const height = twipsToPoints(page.heightTwips);
  switch (from) {
    case "page":
      return { startPt: 0, extentPt: height };
    case "margin":
      return { startPt: marginTopPt, extentPt: height - marginTopPt - bottom };
    case "paragraph":
    case "line":
    case "character":
      return { startPt: paragraphTopPt, extentPt: height - paragraphTopPt - bottom };
    case "column":
      return { startPt: bodyTopPt, extentPt: height - bodyTopPt - bottom };
  }
}

function resolve(position: AnchorPosition, band: Band, sizePt: number): number {
  if (position.kind === "offset") return band.startPt + emuToPoints(position.offsetEmu);
  switch (position.align) {
    case "right":
    case "bottom":
      return band.startPt + band.extentPt - sizePt;
    case "center":
      return band.startPt + (band.extentPt - sizePt) / 2;
    default:
      return band.startPt;
  }
}

const resolvePaint = (paint: ShapePaint, theme: Theme): PlacedPaint => {
  const outlineColor =
    paint.outline?.color === undefined || paint.outline.color === null
      ? null
      : themeColor(theme, paint.outline.color);
  return {
    fillColor: paint.fill === null ? null : themeColor(theme, paint.fill),
    outline:
      paint.outline === null || outlineColor === null
        ? null
        : { color: outlineColor, widthPt: paint.outline.widthPt },
    geometry: paint.geometry,
    path: paint.path,
  };
};

// A text box's own text is laid out once its frame has a place on the page, so it
// arrives here unresolved and is filled in afterwards.
export function resolveContent(
  content: DrawingContent,
  resolvePart: PartResolver,
  theme: Theme,
): PlacedContent {
  if (content.kind === "unknown") return content;
  if (content.kind === "group") {
    return {
      kind: "group",
      children: content.children.map((child) => ({
        ...child,
        content: resolveContent(child.content, resolvePart, theme),
      })),
    };
  }

  const paint = resolvePaint(content.paint, theme);
  if (content.kind === "shape") return { kind: "shape", paint };
  if (content.kind === "text-box") {
    return { kind: "text-box", body: content.body, text: null, paint };
  }

  const part = resolvePart(content.relationshipId);
  return part === null
    ? { kind: "missing-picture", relationshipId: content.relationshipId }
    : { kind: "picture", part, crop: content.crop, paint };
}

// A document declaring no compatibility mode has its objects put on the twip grid,
// which is the whole of what tells its wrapping apart from a modern one: Word draws
// the same picture at 77.30 in one and 77.28 in the other, and where the rounding
// went down the object now stands over the paragraph above it.
const onTheTwip = (positionPt: number, settings: DocumentSettings): number =>
  roundsAnchorsToTwips(settings)
    ? Math.round(positionPt * TWIPS_PER_POINT) / TWIPS_PER_POINT
    : positionPt;

const TWIPS_PER_POINT = 20;

export function placeFloat(input: PlaceFloatInput): PlacedFloat {
  const { anchor } = input;
  const settings = input.settings ?? DEFAULT_SETTINGS;
  const widthPt = input.sizePt?.widthPt ?? emuToPoints(anchor.widthEmu);
  const heightPt = input.sizePt?.heightPt ?? emuToPoints(anchor.heightEmu);
  // **Only the horizontal is moved.** 91 of the 92 anchors the corpus states inside
  // a cell measure their vertical from the paragraph or the line, which stands in
  // the cell already; the one that does not is a text box aligned to the foot of the
  // page in a footer, and what the cell does to that is unmeasured.
  const cell = anchor.inTheCell ? (input.cell ?? null) : null;

  return {
    anchor,
    flip: anchor.flip,
    content: resolveContent(anchor.content, input.resolvePart, input.theme),
    leftPt: resolve(
      anchor.horizontal,
      horizontalBand(input.page, anchor.horizontal.from, cell),
      widthPt,
    ),
    topPt: onTheTwip(
      resolve(anchor.vertical, verticalBand(input, anchor.vertical.from), heightPt),
      settings,
    ),
    widthPt,
    heightPt,
    turnDegrees: anchor.turnDegrees,
  };
}
