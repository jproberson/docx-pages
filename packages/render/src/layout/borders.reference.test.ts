import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

import {
  layOutDocument,
  lookupFontMetrics,
  paintOfCell,
  paintOfParagraph,
  type LaidOutDocument,
  type PaintedFill,
  type PaintedLine,
  type SuppliedFace,
} from "@docx-pages/core";

import { authoredCases } from "../authored/cases.js";
import { authoredFace } from "../authored/faces.js";
import { readFillPlacements, type FillPlacement } from "../pdf/fills.js";
import { readReferenceDocument } from "../testing/documents.js";
import { referenceCases, suppliedFaces, type ReferenceCase } from "../testing/cases.js";

// Every line round a table and every colour behind text, against the same thing
// in Word's own pdf. Word paints both as filled rectangles, so its pdf reports a
// border as a rectangle as long as the edge and as thick as the line: a line of
// ours is agreed when Word filled the same colour along the same edge.
//
// This is the whole of the oracle for the feature. Nothing about a border reaches
// the numbers the other suites compare, since a border is painted rather than laid
// out, and the layout only hears of one through the room it takes.

type Compared = {
  readonly each: ReferenceCase;
  readonly facesFor: () => readonly SuppliedFace[];
};

const FACE = authoredFace();

const CASES: readonly Compared[] = [
  ...referenceCases().map((each) => ({ each, facesFor: suppliedFaces })),
  ...(FACE === null ? [] : authoredCases().map((each) => ({ each, facesFor: () => [FACE] }))),
].filter(({ each }) => each.renderedPath !== null);

// Word's own pdf lands everything on a grid of about a quarter point, so nothing
// here is read finer than that.
const TOLERANCE_PT = 0.3;

// A line whose colour the file left to whatever draws it comes out black.
const colorOf = (color: string | null): string => (color ?? "#000000").toLowerCase();

type Drawn = {
  readonly pageIndex: number;
  readonly lines: readonly PaintedLine[];
  readonly fills: readonly PaintedFill[];
};

function paintedIn(layout: LaidOutDocument): readonly Drawn[] {
  return layout.pages.map((page) => {
    const cells = [...layout.headerCells, ...page.cells, ...layout.footerCells];
    const boxes = [...layout.header, ...page.body, ...layout.footer];
    const painted = [
      ...cells.map((cell) => paintOfCell(cell)),
      ...boxes.flatMap((box) =>
        box.paint === null
          ? []
          : [
              paintOfParagraph(
                box.paint,
                box.lines[0]?.topPt ?? box.markTopPt,
                box.contentBottomPt,
              ),
            ],
      ),
    ];
    return {
      pageIndex: page.index,
      lines: painted.flatMap((each) => each.lines),
      fills: painted.flatMap((each) => each.fills),
    };
  });
}

// Word fills a border in pieces: a corner square at each end and the run between
// them, and a dashed one piece by piece. Any piece of the right colour lying
// along the line, as thick as the line and no thicker, is the line drawn.
function drewLine(line: PaintedLine, drawn: readonly FillPlacement[]): boolean {
  return drawn.some((fill) => {
    const across = line.vertical ? fill.widthPt : fill.heightPt;
    const along = line.vertical ? fill.heightPt : fill.widthPt;
    const at = line.vertical ? fill.leftPt + fill.widthPt / 2 : fill.topPt + fill.heightPt / 2;
    const from = line.vertical ? fill.topPt : fill.leftPt;
    return (
      colorOf(line.color) === fill.color.toLowerCase() &&
      Math.abs(at - line.atPt) <= TOLERANCE_PT &&
      Math.abs(across - line.widthPt) <= TOLERANCE_PT &&
      from + along >= line.fromPt - TOLERANCE_PT &&
      from <= line.toPt + TOLERANCE_PT
    );
  });
}

// A fill is agreed where Word laid the same colour across the same width somewhere
// inside the same room. Word paints a paragraph's fill line by line and a cell's
// twice over, once to its walls and once to its text, so what it drew is a piece
// of what is drawn here rather than the same rectangle.
function drewFill(fill: PaintedFill, drawn: readonly FillPlacement[]): boolean {
  return drawn.some(
    (each) =>
      each.color.toLowerCase() === fill.color.toLowerCase() &&
      Math.abs(each.leftPt - fill.leftPt) <= TOLERANCE_PT &&
      Math.abs(each.leftPt + each.widthPt - (fill.leftPt + fill.widthPt)) <= TOLERANCE_PT &&
      each.topPt >= fill.topPt - TOLERANCE_PT &&
      each.topPt + each.heightPt <= fill.topPt + fill.heightPt + TOLERANCE_PT,
  );
}

// How much of each document Word drew where this project draws it. A case left
// out of this has to agree about every line and every colour; one named here has
// a rule measured and not answered behind it, and the number cannot quietly grow.
type Agreed = {
  readonly lines: number;
  readonly fills: number;
};

const AGREED: Readonly<Record<string, Agreed>> = {
  // Every line but the wave borders, four sides of three tables, which Word draws
  // as a stroked zigzag rather than as a filled rectangle: nothing of one reaches
  // this comparison at all.
  "authored-borders": { lines: 124, fills: 7 },
  // Every horizontal line of the one table, and none of the vertical ones: the
  // table stands about 2.7pt right of where Word puts it, which is the same gap
  // its text has always had ([[a-document-from-outside-the-family]]). The fills
  // are white on white and miss by the same distance.
  h: { lines: 48, fills: 0 },
};

type Agreement = {
  readonly lines: number;
  readonly linesDrawn: number;
  readonly fills: number;
  readonly fillsDrawn: number;
};

async function compare(compared: Compared): Promise<Agreement> {
  const { each } = compared;
  const layout = layOutDocument(readReferenceDocument(each), (request) =>
    lookupFontMetrics(request, compared.facesFor()),
  );
  if (layout.kind !== "laid-out") throw new Error(layout.blocker.kind);

  const drawn = await readFillPlacements(new Uint8Array(readFileSync(each.renderedPath ?? "")));

  let lines = 0;
  let linesDrawn = 0;
  let fills = 0;
  let fillsDrawn = 0;

  for (const page of paintedIn(layout)) {
    const onPage = drawn.filter((fill) => fill.pageIndex === page.pageIndex);
    for (const line of page.lines) {
      lines += 1;
      if (drewLine(line, onPage)) linesDrawn += 1;
    }
    for (const fill of page.fills) {
      fills += 1;
      if (drewFill(fill, onPage)) fillsDrawn += 1;
    }
  }

  return { lines, linesDrawn, fills, fillsDrawn };
}

const agreements = new Map<string, Promise<Agreement>>();

function agreementOf(compared: Compared): Promise<Agreement> {
  const found = agreements.get(compared.each.id);
  if (found !== undefined) return found;

  const made = compare(compared);
  agreements.set(compared.each.id, made);
  return made;
}

const COMPARISON_TIMEOUT_MS = 60_000;

describe.skipIf(CASES.length === 0)(
  "borders and fills against Word",
  { timeout: COMPARISON_TIMEOUT_MS },
  () => {
    for (const compared of CASES) {
      describe(compared.each.id, () => {
        it("draws every line where Word drew one", async () => {
          const { lines, linesDrawn } = await agreementOf(compared);
          const agreed = AGREED[compared.each.id]?.lines ?? lines;
          expect(`${String(linesDrawn)} of ${String(lines)}`).toBe(
            `${String(agreed)} of ${String(lines)}`,
          );
        });

        it("lays every colour where Word laid one", async () => {
          const { fills, fillsDrawn } = await agreementOf(compared);
          const agreed = AGREED[compared.each.id]?.fills ?? fills;
          expect(`${String(fillsDrawn)} of ${String(fills)}`).toBe(
            `${String(agreed)} of ${String(fills)}`,
          );
        });
      });
    }
  },
);
