import type { TextBoxAnchor, TextBoxBody } from "../docx/drawing.js";
import type { StyleTable } from "../docx/styles.js";
import type { DocumentSettings } from "../docx/settings.js";
import {
  measureStack,
  shiftBoxes,
  shiftCells,
  type LayoutBlocker,
  type MetricsResolver,
  type ParagraphBox,
  type PlacedCell,
} from "./stack.js";
import { emuToPoints } from "./units.js";

export type TextBoxRect = {
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

export type PlacedTextBox = {
  readonly boxes: readonly ParagraphBox[];
  // The cells of any table the box holds, which are drawn behind its text like
  // any other table's.
  readonly cells: readonly PlacedCell[];
  readonly contentHeightPt: number;
  // How far the widest line reaches past the box's own left inset, which is what
  // a box that fits itself to its text is as wide as.
  readonly contentWidthPt: number;
};

export type LayOutTextBoxInput = {
  readonly body: TextBoxBody;
  readonly rect: TextBoxRect;
  readonly styles: StyleTable;
  readonly metricsFor: MetricsResolver;
  readonly settings?: DocumentSettings;
  readonly part: string;
};

export type TextBoxLayout =
  | { readonly kind: "laid-out"; readonly text: PlacedTextBox }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

export function layOutTextBox(input: LayOutTextBoxInput): TextBoxLayout {
  const { body, rect } = input;
  const leftPt = rect.leftPt + emuToPoints(body.insets.leftEmu);
  const topPt = rect.topPt + emuToPoints(body.insets.topEmu);
  const widthPt =
    rect.widthPt - emuToPoints(body.insets.leftEmu) - emuToPoints(body.insets.rightEmu);
  const availablePt =
    rect.heightPt - emuToPoints(body.insets.topEmu) - emuToPoints(body.insets.bottomEmu);

  const measured = measureStack({
    blocks: body.blocks,
    styles: input.styles,
    metricsFor: input.metricsFor,
    ...(input.settings === undefined ? {} : { settings: input.settings }),
    part: input.part,
    originPt: topPt,
    leftPt,
    widthPt,
    wraps: body.wraps,
  });

  if (measured.kind === "blocked") return { kind: "blocked", blocker: measured.blocker };

  const offsetPt = seatingOffset(body.anchor, availablePt, measured.heightPt);
  return {
    kind: "laid-out",
    text: {
      boxes: shiftBoxes(measured.boxes, offsetPt),
      cells: shiftCells(measured.cells, offsetPt),
      contentHeightPt: measured.heightPt,
      contentWidthPt: widestParagraphPt(measured.boxes),
    },
  };
}

const widestParagraphPt = (boxes: readonly ParagraphBox[]): number =>
  boxes.reduce((widest, box) => Math.max(widest, box.contentWidthPt), 0);

// Word lets text overflow a box it does not fit, so a negative slack still seats
// the text where the anchor asks rather than clamping it back inside.
function seatingOffset(anchor: TextBoxAnchor, availablePt: number, contentPt: number): number {
  if (anchor === "center") return (availablePt - contentPt) / 2;
  if (anchor === "bottom") return availablePt - contentPt;
  return 0;
}
