// The rectangle a wrapping object keeps text out of: its own frame grown by the
// distances the anchor asks text to stay off it.
export type WrapBand = {
  readonly leftPt: number;
  readonly rightPt: number;
  readonly topPt: number;
  readonly bottomPt: number;
};

// Where a line ended up: the run of free space it was given, at the height it had
// to fall to for that space to exist.
export type LineSlot = {
  readonly topPt: number;
  readonly leftPt: number;
  readonly rightPt: number;
};

export type FitLineInput = {
  readonly topPt: number;
  readonly heightPt: number;
  readonly leftPt: number;
  readonly rightPt: number;
  // What the line takes up, which decides whether a run of free space holds it.
  readonly widthPt: number;
  readonly bands: readonly WrapBand[];
};

// Widths and heights here are sums of exact ratios, so only the last bits of one
// need absorbing; a band that ends exactly where a line starts does not cross it.
const EPSILON = 1e-9;

// Word will not put a line in a sliver of space, however little the line needs.
// Bisected against Word with a document whose objects are placed to the hundredth
// of a point: the least run of free space it takes is exactly 18pt, a quarter
// inch, and 17.99pt is refused. The same beside a `wrapTight` object as beside a
// `wrapSquare` one, against the column's own edge as between two objects, and at
// any size of text.
export const LEAST_SPAN_PT = 18;

type Span = { readonly leftPt: number; readonly rightPt: number };

// The runs of the frame left over once every band crossing the line is taken out
// of it, in the order text would meet them.
export function freeSpans(
  leftPt: number,
  rightPt: number,
  bands: readonly WrapBand[],
): readonly Span[] {
  const ordered = [...bands].sort((one, other) => one.leftPt - other.leftPt);
  const spans: Span[] = [];
  let openPt = leftPt;

  for (const band of ordered) {
    if (band.rightPt <= openPt + EPSILON) continue;
    if (band.leftPt > openPt + EPSILON) {
      spans.push({ leftPt: openPt, rightPt: Math.min(band.leftPt, rightPt) });
    }
    openPt = band.rightPt;
    if (openPt >= rightPt - EPSILON) return spans;
  }

  spans.push({ leftPt: openPt, rightPt });
  return spans;
}

// An object stands in a line's way while it covers the middle of it: Word lets a
// line whose top half is behind an object stay where it is, and moves one whose
// middle is. Measured against Word by moving an object down a little at a time,
// which leaves the line alone until the step below opens up.
const covers = (band: WrapBand, middlePt: number): boolean =>
  band.topPt < middlePt - EPSILON && band.bottomPt > middlePt + EPSILON;

// A line that cannot sit beside the objects in its way drops by its own height and
// tries again, rather than falling to the bottom edge of whatever blocked it, and
// takes the first run of free space wide enough to hold it. Word gives up the same
// way this does: a line that fits nowhere is left in the frame it started in.
export function fitLine(input: FitLineInput): LineSlot {
  const { bands, heightPt, leftPt, rightPt, widthPt } = input;
  const lowestPt = Math.max(...bands.map((band) => band.bottomPt), input.topPt);
  let topPt = input.topPt;

  while (topPt <= lowestPt + EPSILON) {
    const crossing = bands.filter((band) => covers(band, topPt + heightPt / 2));
    const leastPt = Math.max(widthPt, LEAST_SPAN_PT);
    const span = freeSpans(leftPt, rightPt, crossing).find(
      (each) => each.rightPt - each.leftPt >= leastPt - EPSILON,
    );
    if (span !== undefined) return { topPt, leftPt: span.leftPt, rightPt: span.rightPt };
    if (heightPt <= EPSILON) break;
    topPt += heightPt;
  }

  return { topPt, leftPt, rightPt };
}
