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

// A line falls past every object it cannot sit beside, and takes the first run of
// free space wide enough to hold it. Word gives up the same way this does: a line
// that fits nowhere is left in the frame it started in.
export function fitLine(input: FitLineInput): LineSlot {
  const { bands, heightPt, leftPt, rightPt, widthPt } = input;
  let topPt = input.topPt;

  for (;;) {
    const crossing = bands.filter(
      (band) => band.topPt < topPt + heightPt - EPSILON && band.bottomPt > topPt + EPSILON,
    );
    const span = freeSpans(leftPt, rightPt, crossing).find(
      (each) => each.rightPt - each.leftPt >= widthPt - EPSILON,
    );
    if (span !== undefined) return { topPt, leftPt: span.leftPt, rightPt: span.rightPt };

    const below = crossing
      .map((band) => band.bottomPt)
      .filter((bottomPt) => bottomPt > topPt + EPSILON);
    if (below.length === 0) return { topPt, leftPt, rightPt };
    topPt = Math.min(...below);
  }
}
