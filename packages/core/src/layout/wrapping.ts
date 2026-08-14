// A point on the outline a wrapping object keeps text off, in page points.
export type OutlinePoint = { readonly xPt: number; readonly yPt: number };

// The one side of a band a line may sit on. An object that allows both sides says
// nothing, and one asking for the largest has already been resolved to a side by
// the time it gets here, since which side that is depends on the column the band
// stands in rather than on the line.
export type BandSide = "left" | "right";

// The rectangle a wrapping object keeps text out of: its own frame grown by the
// distances the anchor asks text to stay off it. An object whose wrap draws an
// outline narrower than that rectangle in places carries it as well, and the
// rectangle stays the one every line is blocked by.
export type WrapBand = {
  readonly leftPt: number;
  readonly rightPt: number;
  readonly topPt: number;
  readonly bottomPt: number;
  readonly side?: BandSide;
  readonly outline?: readonly OutlinePoint[];
  // Whether the wrap follows the object's outline rather than its frame, which
  // Word treats differently in two ways that go together: see `crosses` and
  // `belowPt` below.
  readonly outlined?: boolean;
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
// any size of text, and of a paragraph mark that has to be moved as much as of a
// line: an empty paragraph offered 10pt beside an object falls past it instead.
export const LEAST_SPAN_PT = 18;

type Span = { readonly leftPt: number; readonly rightPt: number };

// Whether a band that allows text on one side of itself allows it in this run of
// free space. A run lies on the side of a band it does not overlap, and a band that
// allows both sides bars nothing.
const allows = (band: WrapBand, span: Span): boolean =>
  band.side === undefined ||
  (band.side === "left"
    ? span.rightPt <= band.leftPt + EPSILON
    : span.leftPt >= band.rightPt - EPSILON);

// The runs of the frame left over once every band crossing the line is taken out
// of it, in the order text would meet them. A band allowing text on one side of
// itself takes the runs on the other side out as well, so a line that could have
// sat there falls past the object instead.
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
    if (openPt >= rightPt - EPSILON) break;
  }

  if (openPt < rightPt - EPSILON) spans.push({ leftPt: openPt, rightPt });
  return spans.filter((span) => bands.every((band) => allows(band, span)));
}

// How far an outline reaches across between two heights, which is asked of the
// part of each edge lying between them: an edge is straight, so its own extremes
// over that stretch are at the two ends it is cut to.
function acrossPt(outline: readonly OutlinePoint[], topPt: number, bottomPt: number): Span | null {
  const xs: number[] = [];

  for (const [at, from] of outline.entries()) {
    const to = outline[(at + 1) % outline.length];
    if (to === undefined) continue;

    const lowPt = Math.min(from.yPt, to.yPt);
    const highPt = Math.max(from.yPt, to.yPt);
    if (highPt < topPt - EPSILON || lowPt > bottomPt + EPSILON) continue;
    if (highPt - lowPt <= EPSILON) {
      xs.push(from.xPt, to.xPt);
      continue;
    }

    for (const atPt of [Math.max(lowPt, topPt), Math.min(highPt, bottomPt)]) {
      const along = (atPt - from.yPt) / (to.yPt - from.yPt);
      xs.push(from.xPt + along * (to.xPt - from.xPt));
    }
  }

  return xs.length === 0 ? null : { leftPt: Math.min(...xs), rightPt: Math.max(...xs) };
}

// A line meets only the stretch of an object's outline beside it, so the band it
// has to keep clear of is drawn in from its own edges by however far the outline
// falls short of them there. Word wraps a logo's outline this way: the same
// paragraph's taller first line starts a point to the left of the one under it,
// because the widest part of the shape hangs below it.
function besideLine(band: WrapBand, topPt: number, bottomPt: number): WrapBand {
  const outline = band.outline;
  if (outline === undefined) return band;

  const whole = acrossPt(outline, -Infinity, Infinity);
  const here = acrossPt(outline, topPt, bottomPt);
  if (whole === null || here === null) return band;

  return {
    ...band,
    leftPt: band.leftPt + (here.leftPt - whole.leftPt),
    rightPt: band.rightPt - (whole.rightPt - here.rightPt),
  };
}

// An object wrapped to its frame stands in a line's way as soon as its band reaches
// the line at all, however little of it: sweeping a box's bottom edge down through a
// line moved the line at the first half point, and a document's line falls past a
// box reaching only the last third of it. An outline is answered for the middle of
// the line alone, and leaves a line whose top it reaches into where it is: two
// documents place a line beside an outline that covers 4pt of it.
const crosses = (band: WrapBand, topPt: number, bottomPt: number): boolean =>
  band.outlined === true
    ? band.topPt < (topPt + bottomPt) / 2 - EPSILON &&
      band.bottomPt > (topPt + bottomPt) / 2 + EPSILON
    : band.topPt < bottomPt - EPSILON && band.bottomPt > topPt + EPSILON;

/**
 * **Fourteen authored cases say the outline half of that is wrong, and two real pages
 * say it is right. Nothing found so far tells them apart, so it stands as measured
 * and this is written down instead of acted on.**
 *
 * The fourteen were asked of Word on 2026-08-14, three repeats each, over a line
 * standing 108 to 132 in a frame of 36 to 576 beside a band 120pt wide against the
 * right of it, held 9pt off, so a line it reaches is narrowed rather than moved. A
 * band whose top reaches the last 1, 2, 4 or 8pt of the line, and one whose foot
 * reaches the first 1, 2, 4 or 8pt, wrapped to an outline and wrapped to a frame; a
 * text box and a picture, in front of the text and behind it; a square wrap behind
 * the text. **Word narrowed the line in all fourteen**, and the reading above says it
 * should have left eight of them alone. Two controls standing 40pt clear of the line
 * came back whole, so the cases were read right.
 *
 * Every text box in those cases is 20pt tall and holds a line ruled exactly 24pt, so
 * by the rule since measured and built in `bandFor`, that a tight wrap follows a text
 * that has run out of its box, their bands stood 4pt lower than the cases meant: the
 * four asking after the foot of a band reached 5, 6, 8 and 12pt of the line rather
 * than 1, 2, 4 and 8. The four asking after the top of one are as stated, since a
 * text starts where its box does, and so are the four put to a picture, which holds
 * no text to run out of. **A picture reaching 2pt of a line's head still narrowed
 * it**, so the fitting rule moves the numbers and settles nothing here.
 *
 * Reference `d` is the counter-example, over the same geometry. Its line's own place
 * is 600.52; Word draws it at 736.95, eight steps of its own 17.09pt height down,
 * with the tight band's foot at 739.19 standing 2.04pt into it and the line running
 * from 291 straight through that band's ground. The reading above puts it at 737.23.
 * Reference `f` says the same with no fall in it: its band covers the first 4.9pt of
 * the line and Word draws the line at the frame's own left through the band. Word was
 * asked about `d` twice over, by its pdf and by the paragraph oracle, and the oracle
 * says the paragraphs above the line do not move and the three objects stand exactly
 * where the file states them.
 *
 * What has been put to Word and does not separate the two: the wrap kind, the wrap
 * polygon, `behindDoc`, what the object holds, the distances, an object overhanging
 * the frame, a `wrapNone` object covering the same ground, the pile the objects stand
 * in, the page's margins, its header, its footer, the faces, the settings, the other
 * objects on the page, and the paragraphs above. Cut `d` down to any one of those and
 * Word still answers as `d`; rebuild the same three objects at the same offsets from
 * the same paragraph on the same page and it answers as the fourteen.
 */

// Where a line refused its place looks next: past a frame it falls to the edge that
// blocked it, landing on it exactly, and past an outline it steps down by its own
// height. Both are measured, and nothing lands between the two.
const belowPt = (band: WrapBand, topPt: number, heightPt: number): number =>
  band.outlined === true ? topPt + heightPt : band.bottomPt;

// A line that cannot sit beside the objects in its way drops past the nearest of
// them and tries again, and takes the first run of free space wide enough to hold
// it. Word gives up the same way this does: a line that fits nowhere is left in the
// frame it started in.
export function fitLine(input: FitLineInput): LineSlot {
  const { bands, heightPt, leftPt, rightPt, widthPt } = input;
  const lowestPt = Math.max(...bands.map((band) => band.bottomPt), input.topPt);
  let topPt = input.topPt;

  while (topPt <= lowestPt + EPSILON) {
    const crossing = bands.filter((band) => crosses(band, topPt, topPt + heightPt));
    const leastPt = Math.max(widthPt, LEAST_SPAN_PT);
    const span = freeSpans(
      leftPt,
      rightPt,
      crossing.map((band) => besideLine(band, topPt, topPt + heightPt)),
    ).find((each) => each.rightPt - each.leftPt >= leastPt - EPSILON);
    if (span !== undefined) return { topPt, leftPt: span.leftPt, rightPt: span.rightPt };
    // A line wider than the frame it is being laid into has nothing to fall past,
    // and is left where it started rather than falling out of the story.
    if (crossing.length === 0) break;

    const nextPt = Math.min(...crossing.map((band) => belowPt(band, topPt, heightPt)));
    if (!(nextPt > topPt + EPSILON)) break;
    topPt = nextPt;
  }

  return { topPt, leftPt, rightPt };
}

// Where a paragraph's own mark comes to rest, which is not quite where a line
// does. **A mark standing clear of every object is not moved at all, however
// narrow the run of free space it is standing in**, and only a mark that has to be
// moved is held to `LEAST_SPAN_PT` like a line.
//
// Measured against Word on 2026-08-12, over an empty paragraph beside a box put
// down to the quarter point. Indented 144pt with the box 2.25pt to the right of
// that, the paragraph stays where it is; with the box over the indent instead and
// 10pt of room left beyond it, the same paragraph falls to the box's foot rather
// than take the 10pt, and with 180pt left beyond it, it takes that. A numbered
// paragraph asks for the reach from its number to where the number's suffix moves
// the text on to, so hanging the number 18pt in front of the indent it stays with
// 20.25pt of room before the box and falls with 14.25pt, and hanging it 36pt it
// falls with 22pt.
export function fitMark(input: FitLineInput): LineSlot {
  const bottomPt = input.topPt + input.heightPt;
  const standing = freeSpans(
    input.leftPt,
    input.rightPt,
    input.bands
      .filter((band) => crosses(band, input.topPt, bottomPt))
      .map((band) => besideLine(band, input.topPt, bottomPt)),
  ).find(
    (span) =>
      span.leftPt <= input.leftPt + EPSILON &&
      span.rightPt >= input.leftPt + input.widthPt - EPSILON,
  );

  return standing === undefined
    ? fitLine(input)
    : { topPt: input.topPt, leftPt: standing.leftPt, rightPt: standing.rightPt };
}
