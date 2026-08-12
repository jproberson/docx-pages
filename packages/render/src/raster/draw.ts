import { execFile, execFileSync, spawn } from "node:child_process";
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync } from "node:fs";
import { basename, dirname, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { promisify } from "node:util";

import { readPng, type RasterImage } from "./png.js";

// Getting a drawing of a page out of each side.
//
// Word's comes out of its own pdf through poppler. Ours comes out of either a
// browser or, since the writer landed, a pdf of our own through that same poppler,
// which is what `--written` asks for and what `written.ts` does. Both are asked for
// ninety-six pixels to the inch, which is a css pixel to the point and a third, so
// a page of the same size comes back the same size from both.
//
// **Prefer the written path.** A browser was once the only thing that drew what the
// viewer describes, and that is why this file starts one; it stopped being true
// when `writePdf` began walking the same `drawablesOf`. Two rasterisers hint and
// antialias differently and no tuning takes a pixel count to nought, so the browser
// path carries a floor that the written path does not. Measured on 2026-08-11 over
// the eight documents already known to be right: 0.6% against 0.4%, and twelve of
// their twenty pages coming out exactly equal against eighteen.

const CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

const PDF_TO_PNG = "pdftoppm";

export const DOTS_PER_INCH = 96;

const run = promisify(execFile);

export const canDraw = (): boolean => {
  if (!existsSync(CHROME)) return false;
  try {
    execFileSync("which", [PDF_TO_PNG], { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
};

// **Chrome takes the photograph and then does not exit**, measured on 2026-08-10
// at Chrome 151: the file is on disk in about two seconds and the process is
// still there half a minute later, under old headless and new alike. So the file
// is what says the work is done, and the browser is stopped rather than waited
// for. A png ends with an `IEND` chunk, which is how a file half written is told
// from a finished one.
const FINISHED = "IEND";

const WAITING = 60_000;

const finishedPng = (path: string): Uint8Array | null => {
  let bytes: Uint8Array;
  try {
    bytes = new Uint8Array(readFileSync(path));
  } catch {
    return null;
  }
  const tail = String.fromCharCode(...bytes.subarray(Math.max(bytes.length - 8, 0)));
  return tail.includes(FINISHED) ? bytes : null;
};

const sleep = (ms: number): Promise<void> => new Promise((done) => setTimeout(done, ms));

/**
 * A photograph of one html file, the size of the window it is given.
 *
 * A profile of its own, because a headless run that reaches for the one the
 * person at this machine is using will not start at all while their browser is
 * open. Files may reach files, because the stylesheet beside the page names the
 * font files by where they lie rather than copying a thousand of them in.
 */
export async function photograph(
  htmlPath: string,
  widthPx: number,
  heightPx: number,
  intoPath: string,
  profilePath: string,
): Promise<RasterImage> {
  mkdirSync(dirname(resolve(intoPath)), { recursive: true });
  rmSync(resolve(intoPath), { force: true });

  const chrome = spawn(
    CHROME,
    [
      "--headless",
      "--disable-gpu",
      "--hide-scrollbars",
      "--allow-file-access-from-files",
      "--force-device-scale-factor=1",
      "--no-first-run",
      "--no-default-browser-check",
      "--disable-background-networking",
      `--user-data-dir=${resolve(profilePath)}`,
      `--window-size=${String(widthPx)},${String(heightPx)}`,
      // A page holding a hundred pictures of its own is not drawn by the time
      // the window has been made, and a photograph taken then is of half a page.
      "--virtual-time-budget=10000",
      `--screenshot=${resolve(intoPath)}`,
      pathToFileURL(resolve(htmlPath)).href,
    ],
    { detached: true, stdio: "ignore" },
  );

  try {
    for (let waited = 0; waited < WAITING; waited += 50) {
      const drawn = finishedPng(resolve(intoPath));
      if (drawn !== null) return readPng(drawn);
      await sleep(50);
    }
    throw new Error("the browser drew nothing");
  } finally {
    // The whole group: the browser starts helpers of its own and they outlive it.
    try {
      if (chrome.pid !== undefined) process.kill(-chrome.pid, "SIGKILL");
    } catch {
      // Already gone, which is the state being asked for.
    }
  }
}

const numberIn = (name: string): number => Number(/-(\d+)\.png$/.exec(name)?.[1] ?? 0);

/**
 * Every page of a pdf, drawn at the same size as the browser draws ours.
 *
 * How many digits the page number is written in depends on how many pages there
 * are, so the files are found by looking rather than by being named.
 */
export async function drawPdf(
  pdfPath: string,
  intoDirectory: string,
  stem: string,
  discard = false,
): Promise<readonly RasterImage[]> {
  mkdirSync(resolve(intoDirectory), { recursive: true });
  const prefix = resolve(intoDirectory, stem);

  await run(PDF_TO_PNG, ["-r", String(DOTS_PER_INCH), "-png", resolve(pdfPath), prefix], {
    maxBuffer: 64 * 1024 * 1024,
  });

  const drawn = readdirSync(resolve(intoDirectory))
    .filter((each: string) => each.startsWith(`${basename(prefix)}-`) && each.endsWith(".png"))
    .sort((one: string, other: string) => numberIn(one) - numberIn(other));

  return drawn.map((each: string) => {
    const path = resolve(intoDirectory, each);
    const page = readPng(new Uint8Array(readFileSync(path)));
    if (discard) rmSync(path, { force: true });
    return page;
  });
}

// One page out of a photograph of several stacked together.
export function partOf(
  image: RasterImage,
  topPx: number,
  widthPx: number,
  heightPx: number,
): RasterImage {
  const width = Math.min(widthPx, image.width);
  const height = Math.min(heightPx, Math.max(image.height - topPx, 0));
  const pixels = new Uint8Array(width * height * 4);

  for (let row = 0; row < height; row += 1) {
    const from = (topPx + row) * image.width * 4;
    pixels.set(image.pixels.subarray(from, from + width * 4), row * width * 4);
  }

  return { width, height, pixels };
}
