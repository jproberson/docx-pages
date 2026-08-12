import type { RasterImage } from "./png.js";

// How different two drawings of one page look.
//
// The line comparison beside this counts baselines and lefts over lines we laid
// out, and so can see nothing it did not lay out: not paint, not a picture, not a
// border, not a crop, not what covers what, and least of all content we draw
// nowhere at all, which never enters its denominator. A page that scored 35 of its
// 35 lines was wrong five ways. This answers the other question: put the two
// drawings side by side and say how much of the page does not match.
//
// **A pixel count never reaches nought**, and no tuning will make it. Chrome and
// Word hint and antialias differently, so the same glyph in the same place differs
// pixel for pixel every time. So nothing here counts pixels: the page is cut into
// cells a few points across, and a cell is judged by how much ink is in it rather
// than by which pixels hold it. Ink is what the two rasterisers agree about: they
// put the same glyph in the same place out of the same outline, and differ over
// how its edges are shaded. What that costs when nothing at all is wrong was
// measured rather than assumed, and `floor.ts` is what measures it.

// A cell nine points across, which is about the height of the lower case of a
// 12pt face. Text a line out of place lands in different cells; text a hair out
// of place lands in the same ones with the same ink in them.
export const CELL_PX = 12;

export type Tolerances = {
  // How much more of a cell's ink one side may lay down than the other, as a
  // share of the cell being covered in black.
  readonly ink: number;
  // And how far apart what the two drew that cell in may be, out of 255. This is
  // what still answers where a cell is inked on both sides: a photograph cropped
  // differently, or a yellow arrow missing off one, leaves the ink alone and the
  // colour not.
  readonly colour: number;
  // Below this a cell counts as blank, so that a page's margins are not counted
  // as an agreement and the odd stray pixel is not counted as a page.
  readonly blank: number;
};

/**
 * Measured on 2026-08-10 by scoring the eight documents whose pages are already
 * known to be right against three known to be wrong, over cells of 8, 12, 16 and
 * 24 pixels and four tolerances each. What the numbers said:
 *
 * - **Cells of 8 pixels cost five times the floor of cells of 12** and bought
 *   nothing: 3.3% of a right page against 1.6%, for a wrong page reading 8.7%
 *   rather than 7.9%. That is a cell smaller than a glyph, so which pixels an
 *   edge is shaded over starts to matter again.
 * - **Cells of 16 and 24 lower the floor no further and cost the signal**: at 24
 *   a document five faults deep read 3.5% where at 12 it read 5.5%.
 * - **Colour is the noisy half.** Holding it to 24 out of 255 put the floor at
 *   0.9% and its worst at 2.4%; at 48 they are 0.3% and 1.6%, and a wrong page
 *   fell only from 8.6% to 7.9%. Ink already answers for a yellow arrow, since
 *   the channel furthest from white is what counts it.
 */
export const TOLERANCES: Tolerances = { ink: 0.12, colour: 48, blank: 0.004 };

export type PageGrid = {
  readonly columns: number;
  readonly rows: number;
  // Per cell, how much ink is in it and how light what it holds is.
  readonly ink: Float32Array;
  readonly lightness: Float32Array;
};

export function gridOf(image: RasterImage, cellPx: number = CELL_PX): PageGrid {
  const columns = Math.ceil(image.width / cellPx);
  const rows = Math.ceil(image.height / cellPx);
  const ink = new Float32Array(columns * rows);
  const lightness = new Float32Array(columns * rows);
  const counted = new Float32Array(columns * rows);

  for (let y = 0; y < image.height; y += 1) {
    for (let x = 0; x < image.width; x += 1) {
      const at = (y * image.width + x) * 4;
      const alpha = (image.pixels[at + 3] ?? 255) / 255;
      // What is drawn nowhere is the paper, and the paper is white on both sides.
      const red = 255 - alpha * (255 - (image.pixels[at] ?? 255));
      const green = 255 - alpha * (255 - (image.pixels[at + 1] ?? 255));
      const blue = 255 - alpha * (255 - (image.pixels[at + 2] ?? 255));

      const cell = Math.floor(y / cellPx) * columns + Math.floor(x / cellPx);
      counted[cell] = (counted[cell] ?? 0) + 1;
      lightness[cell] = (lightness[cell] ?? 0) + (red + green + blue) / 3;
      // The channel furthest from white, so that yellow is as much ink as black
      // is. By lightness it is hardly ink at all.
      ink[cell] = (ink[cell] ?? 0) + (255 - Math.min(red, green, blue)) / 255;
    }
  }

  for (let cell = 0; cell < columns * rows; cell += 1) {
    const pixels = counted[cell] ?? 1;
    ink[cell] = (ink[cell] ?? 0) / pixels;
    lightness[cell] = (lightness[cell] ?? 0) / pixels;
  }

  return { columns, rows, ink, lightness };
}

export type Difference = {
  // Cells one side or the other drew something in, which is the whole of what
  // there is to be right or wrong about. A blank margin agreed on by both is not
  // an agreement worth counting.
  readonly interesting: number;
  readonly differing: number;
};

const BLANK: PageGrid = {
  columns: 0,
  rows: 0,
  ink: new Float32Array(0),
  lightness: new Float32Array(0),
};

const cellAt = (grid: PageGrid, column: number, row: number): readonly [number, number] =>
  column >= grid.columns || row >= grid.rows
    ? [0, 255]
    : [
        grid.ink[row * grid.columns + column] ?? 0,
        grid.lightness[row * grid.columns + column] ?? 255,
      ];

/**
 * A page one side drew and the other did not is `null` on that side, and every
 * cell the other drew in counts against it: a document making the wrong number of
 * pages is wrong before a pixel is compared.
 *
 * Where the two pages are different sizes they are laid over each other by their
 * top left corners and the rest counts as blank, which is what a page of the
 * wrong geometry deserves.
 */
export function differenceBetween(
  ours: PageGrid | null,
  theirs: PageGrid | null,
  tolerances: Tolerances = TOLERANCES,
): Difference {
  const one = ours ?? BLANK;
  const other = theirs ?? BLANK;

  let interesting = 0;
  let differing = 0;

  for (let row = 0; row < Math.max(one.rows, other.rows); row += 1) {
    for (let column = 0; column < Math.max(one.columns, other.columns); column += 1) {
      const [ourInk, ourLightness] = cellAt(one, column, row);
      const [theirInk, theirLightness] = cellAt(other, column, row);
      if (ourInk < tolerances.blank && theirInk < tolerances.blank) continue;

      interesting += 1;
      if (
        Math.abs(ourInk - theirInk) > tolerances.ink ||
        Math.abs(ourLightness - theirLightness) > tolerances.colour
      ) {
        differing += 1;
      }
    }
  }

  return { interesting, differing };
}

export const shareOf = (difference: Difference): number =>
  difference.interesting === 0 ? 0 : difference.differing / difference.interesting;
