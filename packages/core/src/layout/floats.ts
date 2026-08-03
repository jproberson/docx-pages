import type { AnchorOrigin, AnchorPosition, FloatingAnchor } from "../docx/anchors.js";
import type { SectionGeometry } from "../docx/section.js";
import { twipsToPoints } from "./document.js";
import { EMU_PER_POINT } from "./stack.js";

export type PlacedFloat = {
  readonly anchor: FloatingAnchor;
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

export type PlaceFloatInput = {
  readonly anchor: FloatingAnchor;
  readonly page: SectionGeometry;
  readonly paragraphTopPt: number;
  readonly bodyTopPt: number;
};

const emuToPoints = (emu: number): number => emu / EMU_PER_POINT;

type Band = { readonly startPt: number; readonly extentPt: number };

function horizontalBand(page: SectionGeometry, from: AnchorOrigin): Band {
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

function verticalBand(input: PlaceFloatInput, from: AnchorOrigin): Band {
  const { page, paragraphTopPt, bodyTopPt } = input;
  const top = twipsToPoints(page.margin.topTwips);
  const bottom = twipsToPoints(page.margin.bottomTwips);
  const height = twipsToPoints(page.heightTwips);
  switch (from) {
    case "page":
      return { startPt: 0, extentPt: height };
    case "margin":
      return { startPt: top, extentPt: height - top - bottom };
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

export function placeFloat(input: PlaceFloatInput): PlacedFloat {
  const { anchor } = input;
  const widthPt = emuToPoints(anchor.widthEmu);
  const heightPt = emuToPoints(anchor.heightEmu);

  return {
    anchor,
    leftPt: resolve(anchor.horizontal, horizontalBand(input.page, anchor.horizontal.from), widthPt),
    topPt: resolve(anchor.vertical, verticalBand(input, anchor.vertical.from), heightPt),
    widthPt,
    heightPt,
  };
}
