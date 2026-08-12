import type { RasterImage } from "./png.js";

// Two drawings of a page in one image, so that what the score is counting can be
// looked at rather than believed.
//
// Ours is laid down in red and Word's in green, both taken off white. Where the
// two put the same ink the page comes out black; where only ours did it comes out
// red, and where only Word's did, green. So a fault reads as its colour: a line
// of green with a line of red under it is text of ours that is low, a block of
// green alone is content we draw nowhere at all, and a page that is right is a
// page in black and white.

const inkOf = (image: RasterImage | null, x: number, y: number): number => {
  if (image === null || x >= image.width || y >= image.height) return 0;
  const at = (y * image.width + x) * 4;
  const alpha = (image.pixels[at + 3] ?? 255) / 255;
  const red = 255 - alpha * (255 - (image.pixels[at] ?? 255));
  const green = 255 - alpha * (255 - (image.pixels[at + 1] ?? 255));
  const blue = 255 - alpha * (255 - (image.pixels[at + 2] ?? 255));
  return (255 - Math.min(red, green, blue)) / 255;
};

export function overlayOf(ours: RasterImage | null, theirs: RasterImage | null): RasterImage {
  const width = Math.max(ours?.width ?? 0, theirs?.width ?? 0, 1);
  const height = Math.max(ours?.height ?? 0, theirs?.height ?? 0, 1);
  const pixels = new Uint8Array(width * height * 4);

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const mine = inkOf(ours, x, y);
      const theirsHere = inkOf(theirs, x, y);
      const at = (y * width + x) * 4;
      pixels[at] = Math.round(255 * (1 - theirsHere));
      pixels[at + 1] = Math.round(255 * (1 - mine));
      pixels[at + 2] = Math.round(255 * (1 - Math.max(mine, theirsHere)));
      pixels[at + 3] = 255;
    }
  }

  return { width, height, pixels };
}
