import type { TextBoxAnchor, TextBoxBody } from "../docx/drawing.js";
import type { StyleTable } from "../docx/styles.js";
import {
  measureStack,
  type LayoutBlocker,
  type MetricsResolver,
  type ParagraphBox,
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
  readonly contentHeightPt: number;
};

export type LayOutTextBoxInput = {
  readonly body: TextBoxBody;
  readonly rect: TextBoxRect;
  readonly styles: StyleTable;
  readonly metricsFor: MetricsResolver;
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
      boxes: offsetPt === 0 ? measured.boxes : measured.boxes.map((box) => shift(box, offsetPt)),
      contentHeightPt: measured.heightPt,
    },
  };
}

// Word lets text overflow a box it does not fit, so a negative slack still seats
// the text where the anchor asks rather than clamping it back inside.
function seatingOffset(anchor: TextBoxAnchor, availablePt: number, contentPt: number): number {
  if (anchor === "center") return (availablePt - contentPt) / 2;
  if (anchor === "bottom") return availablePt - contentPt;
  return 0;
}

const shift = (box: ParagraphBox, byPt: number): ParagraphBox => ({
  ...box,
  topPt: box.topPt + byPt,
  lines: box.lines.map((line) => ({
    ...line,
    topPt: line.topPt + byPt,
    baselinePt: line.baselinePt + byPt,
  })),
});
