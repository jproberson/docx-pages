import { existsSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

import {
  layOutDocument,
  openDocx,
  substitutingMetrics,
  WORD_FALLBACK_FACES,
} from "@docx-pages/core";
import { imageResolver } from "@docx-pages/viewer";

import { corpusFaces } from "../corpus/faces.js";
import {
  differenceBetween,
  gridOf,
  TOLERANCES,
  type PageGrid,
  type Tolerances,
} from "./difference.js";
import { writePages } from "./document.js";
import { drawPdf, partOf, photograph } from "./draw.js";
import { faceStylesheet } from "./faces.js";
import type { RasterImage } from "./png.js";

// One document beside Word's own drawing of it, page for page.

// How many pages go into one photograph. A browser is given a window as tall as
// the pages stacked in it, and a window as tall as a long document is one no
// browser will give.
const PAGES_AT_ONCE = 8;

export type Workspace = {
  readonly directory: string;
  readonly stylesheet: string;
  readonly profile: string;
  // Whether the drawings are left on disk to be looked at. A sweep of seven
  // hundred documents would otherwise fill the disk with pages nobody reads.
  readonly keep: boolean;
};

// The stylesheet is written once for a whole run rather than once a document:
// it names every face this machine has, and there are twelve hundred of them.
export function workspaceIn(directory: string, keep: boolean): Workspace {
  mkdirSync(resolve(directory), { recursive: true });
  const stylesheet = "fonts.css";
  writeFileSync(resolve(directory, stylesheet), faceStylesheet());
  return {
    directory: resolve(directory),
    stylesheet,
    profile: resolve(directory, "profile"),
    keep,
  };
}

export type Looks = {
  readonly id: string;
  readonly outcome: "compared" | "blocked" | "threw" | "not drawn";
  readonly pagesOurs: number;
  readonly pagesWord: number;
  readonly facesStoodIn: number;
  readonly asks: readonly string[];
  // Cells one side or the other drew in, and how many of those the two do not
  // agree about.
  readonly interesting: number;
  readonly differing: number;
  readonly detail: string;
};

const empty = (id: string, outcome: Looks["outcome"], detail: string): Looks => ({
  id,
  outcome,
  pagesOurs: 0,
  pagesWord: 0,
  facesStoodIn: 0,
  asks: [],
  interesting: 0,
  differing: 0,
  detail,
});

export const shareOfLooks = (looks: Looks): number =>
  looks.interesting === 0 ? 0 : looks.differing / looks.interesting;

export type OurPages = {
  readonly pages: readonly RasterImage[];
  readonly facesStoodIn: number;
  readonly asks: readonly string[];
};

// A page at a time, so that a long document is never all in memory at once: the
// sweep turns each into a grid a thousandth of its size and lets it go.
async function eachOfOurPages(
  bytes: Uint8Array,
  id: string,
  workspace: Workspace,
  take: (page: RasterImage) => void,
): Promise<Omit<OurPages, "pages">> {
  const measuring = substitutingMetrics(corpusFaces(), WORD_FALLBACK_FACES);
  const pkg = openDocx(bytes);
  const laid = layOutDocument(pkg, measuring);
  if (laid.kind !== "laid-out") throw new Error(`blocked: ${laid.blocker.kind}`);

  const imageUrl = imageResolver(pkg, measuring.metricsFor);

  for (let from = 0; from < laid.pages.length; from += PAGES_AT_ONCE) {
    const stem = `${id}.ours-${String(from)}`;
    const htmlPath = resolve(workspace.directory, `${stem}.html`);
    const pngPath = resolve(workspace.directory, `${stem}.png`);

    const written = writePages(
      htmlPath,
      laid,
      imageUrl,
      workspace.stylesheet,
      from,
      from + PAGES_AT_ONCE,
    );
    const photo = await photograph(
      htmlPath,
      written.widthPx,
      written.heightPx,
      pngPath,
      workspace.profile,
    );

    for (const page of written.pages) take(partOf(photo, page.topPx, page.widthPx, page.heightPx));

    if (!workspace.keep) {
      rmSync(htmlPath, { force: true });
      rmSync(pngPath, { force: true });
    }
  }

  return {
    facesStoodIn: measuring.substitutions().length,
    asks: laid.unhonoured.map((each) => each.kind),
  };
}

export async function ourPages(
  bytes: Uint8Array,
  id: string,
  workspace: Workspace,
): Promise<OurPages> {
  const pages: RasterImage[] = [];
  const rest = await eachOfOurPages(bytes, id, workspace, (page) => pages.push(page));
  return { pages, ...rest };
}

export async function wordPages(
  drawnPath: string,
  id: string,
  workspace: Workspace,
): Promise<readonly RasterImage[]> {
  const stem = `${id}.word`;
  const drawn = await drawPdf(drawnPath, workspace.directory, stem, !workspace.keep);
  return drawn;
}

/**
 * How much of this document does not look like Word's drawing of it.
 *
 * **The page count is part of the answer and not a note beside it.** A document
 * making one page too few is wrong about a whole page before a pixel is compared,
 * so a page one side drew and the other did not counts every cell the one side
 * drew in against it.
 */
export async function looksOf(
  bytes: Uint8Array,
  id: string,
  drawnPath: string,
  workspace: Workspace,
  tolerances: Tolerances = TOLERANCES,
): Promise<Looks> {
  if (!existsSync(drawnPath)) return empty(id, "not drawn", "no pdf of Word's");

  const ours: PageGrid[] = [];
  let rest: Omit<OurPages, "pages">;
  try {
    rest = await eachOfOurPages(bytes, id, workspace, (page) => ours.push(gridOf(page)));
  } catch (thrown) {
    const detail = thrown instanceof Error ? thrown.message : String(thrown);
    return empty(id, detail.startsWith("blocked: ") ? "blocked" : "threw", detail);
  }

  let theirs: readonly PageGrid[];
  try {
    theirs = (await wordPages(drawnPath, id, workspace)).map((page) => gridOf(page));
  } catch (thrown) {
    return empty(id, "not drawn", thrown instanceof Error ? thrown.message : String(thrown));
  }

  let interesting = 0;
  let differing = 0;
  for (let at = 0; at < Math.max(ours.length, theirs.length); at += 1) {
    const difference = differenceBetween(ours[at] ?? null, theirs[at] ?? null, tolerances);
    interesting += difference.interesting;
    differing += difference.differing;
  }

  return {
    id,
    outcome: "compared",
    pagesOurs: ours.length,
    pagesWord: theirs.length,
    facesStoodIn: rest.facesStoodIn,
    asks: rest.asks,
    interesting,
    differing,
    detail: ours.length === theirs.length ? "" : "a different number of pages",
  };
}
