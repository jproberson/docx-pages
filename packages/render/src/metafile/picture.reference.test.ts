import { describe, expect, it } from "vitest";

import {
  layOutDocument,
  lookupFontMetrics,
  type FaceRequest,
  type MetafilePicture,
  type MetafileShape,
  type MetricsLookup,
} from "@docx-pages/core";
import { imageResolver } from "@docx-pages/viewer";

import { readFillPlacements, type FillPlacement } from "../pdf/fills.js";
import { readTextPlacements, type TextPlacement } from "../pdf/text.js";
import { referenceCases, suppliedFaces, type ReferenceCase } from "../testing/cases.js";
import { readReferenceDocument, readRenderedPages } from "../testing/documents.js";

const CASES = referenceCases().filter(
  (each) => each.renderedPath !== null && each.metafileFills !== null,
);

// Where a metafile was played onto the page: its shapes are measured in the units
// the recording was made in, and the frame the document gave it decides the scale
// on each axis on its own.
type Placement = {
  readonly picture: MetafilePicture;
  readonly leftPt: number;
  readonly topPt: number;
  readonly acrossPt: number;
  readonly downPt: number;
};

function placedMetafiles(each: ReferenceCase): readonly Placement[] {
  const pkg = readReferenceDocument(each);
  const supplied = suppliedFaces();
  const metricsFor = (request: FaceRequest): MetricsLookup => lookupFontMetrics(request, supplied);

  const layout = layOutDocument(pkg, metricsFor);
  if (layout.kind !== "laid-out") throw new Error(`blocked: ${layout.blocker.kind}`);
  const imageUrl = imageResolver(pkg, metricsFor);

  const drawn = [
    ...layout.headerFloats,
    ...layout.footerFloats,
    ...layout.headerInlines,
    ...layout.footerInlines,
    ...layout.pages.flatMap((page) => [...page.floats, ...page.inlines]),
  ];

  return drawn.flatMap((placed) => {
    if (placed.content.kind !== "picture") return [];
    const image = imageUrl(placed.content.part);
    if (image?.kind !== "metafile") return [];
    return [
      {
        picture: image.picture,
        leftPt: placed.leftPt,
        topPt: placed.topPt,
        acrossPt: placed.widthPt / image.picture.widthUnits,
        downPt: placed.heightPt / image.picture.heightUnits,
      },
    ];
  });
}

const near = (one: number, other: number, tolerancePt: number): boolean =>
  Math.abs(one - other) <= tolerancePt;

// A block of colour is the same block Word drew when it is the same colour and
// covers the same rectangle. Word writes the path before it is cut to the clip, as
// the metafile states it before being cut to its own, so the two are the same
// rectangle to compare.
function fillDrawnBy(
  drawn: readonly FillPlacement[],
  where: Placement,
  shape: Extract<MetafileShape, { kind: "fill" }>,
  tolerancePt: number,
): FillPlacement | undefined {
  const leftPt = where.leftPt + shape.rect.leftUnits * where.acrossPt;
  const topPt = where.topPt + shape.rect.topUnits * where.downPt;
  const widthPt = shape.rect.widthUnits * where.acrossPt;
  const heightPt = shape.rect.heightUnits * where.downPt;

  return drawn.find(
    (item) =>
      item.color === shape.color &&
      near(item.leftPt, leftPt, tolerancePt) &&
      near(item.topPt, topPt, tolerancePt) &&
      near(item.widthPt, widthPt, tolerancePt) &&
      near(item.heightPt, heightPt, tolerancePt),
  );
}

// The text of a run cannot be read back out of Word's own pdf: the face it drew
// this diagram in writes some pairs of letters as one glyph of its own, which
// comes back under whatever character that glyph happens to shadow. Where the run
// starts and what baseline it sits on are the whole of what the playing decides,
// and both of those Word does say.
function runDrawnBy(
  drawn: readonly TextPlacement[],
  where: Placement,
  shape: Extract<MetafileShape, { kind: "text" }>,
  tolerancePt: number,
): TextPlacement | undefined {
  const leftPt = where.leftPt + (shape.xUnits[0] ?? 0) * where.acrossPt;
  const baselinePt = where.topPt + shape.baselineUnits * where.downPt;
  return drawn.find(
    (item) =>
      near(item.leftPt, leftPt, tolerancePt) && near(item.baselinePt, baselinePt, tolerancePt),
  );
}

describe.skipIf(CASES.length === 0)("a metafile picture in a real document", () => {
  for (const each of CASES) {
    describe(each.id, () => {
      // Every block of colour and every run of text the recording draws, against
      // what Word drew from the same recording. A metafile decides both the scale
      // it is played at and where each of its shapes lands, so a rendering that
      // agrees with Word on all of them has played it the way Word plays it.
      it("draws each shape of each metafile where Word drew it", async () => {
        const placements = placedMetafiles(each);
        // Reading a pdf takes the bytes it was handed, so each pass gets its own.
        const fills = await readFillPlacements(readRenderedPages(each));
        const runs = await readTextPlacements(readRenderedPages(each));

        let fillsPlaced = 0;
        let runsPlaced = 0;
        const missing: string[] = [];

        for (const where of placements) {
          for (const shape of where.picture.shapes) {
            if (shape.kind === "fill") {
              if (fillDrawnBy(fills, where, shape, each.tolerancePt) !== undefined) {
                fillsPlaced += 1;
              } else {
                missing.push(
                  `fill ${shape.color} at ${String(shape.rect.leftUnits)},${String(shape.rect.topUnits)}`,
                );
              }
            }
            if (shape.kind === "text") {
              if (runDrawnBy(runs, where, shape, each.tolerancePt) !== undefined) {
                runsPlaced += 1;
              } else {
                missing.push(
                  `run at ${String(shape.xUnits[0] ?? 0)},${String(shape.baselineUnits)}`,
                );
              }
            }
          }
        }

        expect(missing).toEqual([]);
        expect(fillsPlaced).toBe(each.metafileFills);
        expect(runsPlaced).toBe(each.metafileRuns);
      });
    });
  }
});
