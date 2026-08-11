import { describe, expect, it } from "vitest";

import {
  layOutDocument,
  lookupFontMetrics,
  type FaceRequest,
  type LaidOutDocument,
  type MetricsResolver,
  type PlacedFloat,
  type PlacedInline,
} from "@docx-pages/core";

import { authoredCases } from "../authored/cases.js";
import { authoredFace, authoredMetrics } from "../authored/faces.js";
import { readImagePlacements, type PlacedRect } from "../pdf/placements.js";
import { readReferenceDocument, readRenderedPages } from "../testing/documents.js";
import { referenceCases, suppliedFaces, type ReferenceCase } from "../testing/cases.js";

// Where a picture stands, asked of Word's own pdf rather than of a number written
// down about it. Every document Word has rendered answers this, which is what
// separates it from the origins in the manifest: those were measured one document
// at a time and only two of the real ones ever were.

type Compared = {
  readonly each: ReferenceCase;
  readonly metricsFor: () => MetricsResolver;
};

const FACE = authoredFace();

const CASES: readonly Compared[] = [
  ...referenceCases().map((each) => ({
    each,
    metricsFor: () => {
      const faces = suppliedFaces();
      return (request: FaceRequest) => lookupFontMetrics(request, faces);
    },
  })),
  ...(FACE === null ? [] : authoredCases().map((each) => ({ each, metricsFor: authoredMetrics }))),
].filter(({ each }) => each.renderedPath !== null);

type Box = {
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

type OurPicture = Box & { readonly pageIndex: number };

// Word writes the whole of a cropped picture into its pdf and clips it, and a clip
// is not something the reader follows, so what comes back is the picture at its
// full extent. Ours is the part left showing, which is that extent less what each
// edge's crop took off it.
function uncropped(box: Box, crop: { left: number; top: number; right: number; bottom: number }) {
  const widthPt = box.widthPt / (1 - crop.left - crop.right);
  const heightPt = box.heightPt / (1 - crop.top - crop.bottom);
  return {
    leftPt: box.leftPt - widthPt * crop.left,
    topPt: box.topPt - heightPt * crop.top,
    widthPt,
    heightPt,
  };
}

// A picture whose part is missing still takes up its frame, and Word draws
// something there, so it stands beside the ones that resolved.
const pictureAt = (each: PlacedFloat | PlacedInline, pageIndex: number): OurPicture | null => {
  if (each.content.kind === "missing-picture") return { ...each, pageIndex };
  if (each.content.kind !== "picture") return null;
  return { ...uncropped(each, each.content.crop), pageIndex };
};

// The header and the footer are drawn again on every page, so every page of the
// pdf holds their pictures as well as its own.
function ourPictures(layout: LaidOutDocument): readonly OurPicture[] {
  const repeated = [
    ...(layout.pages[0]?.headerFloats ?? []),
    ...(layout.pages[0]?.footerFloats ?? []),
    ...(layout.pages[0]?.headerInlines ?? []),
    ...(layout.pages[0]?.footerInlines ?? []),
  ];
  return layout.pages.flatMap((page) =>
    [...repeated, ...page.floats, ...page.inlines].flatMap((each) => {
      const found = pictureAt(each, page.index);
      return found === null ? [] : [found];
    }),
  );
}

const centre = (box: Box): readonly [number, number] => [
  box.leftPt + box.widthPt / 2,
  box.topPt + box.heightPt / 2,
];

const apart = (one: Box, other: Box): number => {
  const [x, y] = centre(one);
  const [otherX, otherY] = centre(other);
  return Math.hypot(x - otherX, y - otherY);
};

const agrees = (one: Box, other: Box, tolerancePt: number): boolean =>
  Math.abs(one.leftPt - other.leftPt) <= tolerancePt &&
  Math.abs(one.topPt - other.topPt) <= tolerancePt &&
  Math.abs(one.widthPt - other.widthPt) <= tolerancePt &&
  Math.abs(one.heightPt - other.heightPt) <= tolerancePt;

const spell = (box: Box): string =>
  `${box.widthPt.toFixed(1)}x${box.heightPt.toFixed(1)} at ${box.leftPt.toFixed(1)},${box.topPt.toFixed(1)}`;

// A picture of ours and an image of Word's are either a pair or they are not, and
// the three answers are held apart: a pair standing in different places is a fault
// whatever the document, while one side having nothing to pair with says only that
// the two are not the same population, and the document counts how far apart they
// are known to be.
type Pairing = {
  readonly misplaced: readonly string[];
  readonly onlyWord: readonly string[];
  readonly onlyOurs: readonly string[];
};

// Word's pictures are taken one at a time by the nearest of ours left unclaimed on
// the same page, so two pictures of a size do not swap and report themselves as a
// pair of misplacements.
function pair(
  ours: readonly OurPicture[],
  drawn: readonly PlacedRect[],
  tolerancePt: number,
): Pairing {
  const taken = new Set<number>();
  const misplaced: string[] = [];
  const onlyWord: string[] = [];

  for (const item of drawn) {
    let nearest = -1;
    let apartBy = Number.POSITIVE_INFINITY;
    ours.forEach((mine, at) => {
      if (taken.has(at) || mine.pageIndex !== item.pageIndex) return;
      const away = apart(mine, item);
      if (away < apartBy) {
        apartBy = away;
        nearest = at;
      }
    });

    const mine = nearest === -1 ? undefined : ours[nearest];
    if (mine === undefined) {
      onlyWord.push(`page ${String(item.pageIndex)}: Word drew ${spell(item)}`);
      continue;
    }
    taken.add(nearest);
    if (!agrees(mine, item, tolerancePt)) {
      misplaced.push(`page ${String(item.pageIndex)}: Word ${spell(item)}, ours ${spell(mine)}`);
    }
  }

  return {
    misplaced,
    onlyWord,
    onlyOurs: ours
      .filter((_, at) => !taken.has(at))
      .map((mine) => `page ${String(mine.pageIndex)}: we drew ${spell(mine)}`),
  };
}

// Laying a document out and reading a pdf of it back is a second or two of work,
// and both assertions below ask the same question of the same document.
const pairings = new Map<string, Promise<Pairing>>();

function pairingOf(compared: Compared): Promise<Pairing> {
  const found = pairings.get(compared.each.id);
  if (found !== undefined) return found;

  const made = (async () => {
    const result = layOutDocument(readReferenceDocument(compared.each), compared.metricsFor());
    if (result.kind !== "laid-out") throw new Error(`blocked: ${result.blocker.kind}`);
    const drawn = await readImagePlacements(readRenderedPages(compared.each));
    return pair(
      ourPictures(result),
      drawn.map((placement) => placement.rect),
      compared.each.tolerancePt,
    );
  })();
  pairings.set(compared.each.id, made);
  return made;
}

const PAIRING_TIMEOUT_MS = 60_000;

describe.skipIf(CASES.length === 0)(
  "pictures against Word's own output",
  { timeout: PAIRING_TIMEOUT_MS },
  () => {
    for (const compared of CASES) {
      describe(compared.each.id, () => {
        it("stands every picture where Word stood it, at the size Word drew it", async () => {
          expect((await pairingOf(compared)).misplaced).toStrictEqual([]);
        });

        it("pairs every picture with an image but the ones Word drew as vector", async () => {
          const { onlyOurs } = await pairingOf(compared);
          expect(onlyOurs.length).toBe(compared.each.picturesWordDrewWithoutAnImage);
        });

        it("pairs every image with a picture but the ones that are no picture", async () => {
          const { onlyWord } = await pairingOf(compared);
          expect(onlyWord.length).toBe(compared.each.imagesWordDrewOutsideAPicture);
        });
      });
    }
  },
);
