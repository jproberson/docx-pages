import { existsSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

import {
  layOutDocument,
  openDocx,
  substitutingMetrics,
  WORD_FALLBACK_FACES,
  readFaceAlternatives,
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

// One page of the two drawings: cells one side or the other drew in, and how many
// of those the two do not agree about.
export type PageLooks = {
  readonly interesting: number;
  readonly differing: number;
};

export type Looks = {
  readonly id: string;
  readonly outcome: "compared" | "blocked" | "threw" | "not drawn";
  readonly pagesOurs: number;
  readonly pagesWord: number;
  readonly facesStoodIn: number;
  readonly asks: readonly string[];
  readonly interesting: number;
  readonly differing: number;
  // Kept per page as well as summed, because **a document's share dilutes a long
  // document**: one badly wrong page in a document of twenty-two moves its total
  // about two percent, where the same page alone is the whole of a one-pager. A
  // queue read off the total is a queue sorted by how short a document is.
  readonly pages: readonly PageLooks[];
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
  pages: [],
  detail,
});

export const shareOf = (page: PageLooks): number =>
  page.interesting === 0 ? 0 : page.differing / page.interesting;

export const shareOfLooks = (looks: Looks): number => shareOf(looks);

// A page holding almost nothing can be wholly wrong about the little it holds, so
// a page has to draw something before it can lead a ranking. A twentieth of a
// letter page's cells, which is about three lines of text.
const ENOUGH = 250;

/**
 * The worst single page, which is what a queue should be read off. A document's
 * own share answers a different question: how much of this document is wrong,
 * rather than how wrong is the worst thing in it.
 *
 * Falls back to the document's share where no page draws enough to judge.
 */
export function worstPageOf(looks: Looks): number {
  const judged = looks.pages.filter((each) => each.interesting >= ENOUGH);
  if (judged.length === 0) return shareOf(looks);
  return Math.max(...judged.map(shareOf));
}

export type OurPages = {
  readonly pages: readonly RasterImage[];
  readonly facesStoodIn: number;
  readonly asks: readonly string[];
};

// How our side of the comparison is drawn. Passed in rather than chosen here, so
// that this module never reaches for the writer and the writer can go on naming
// the types in this one.
export type DrawOurs = (bytes: Uint8Array, id: string, workspace: Workspace) => Promise<OurPages>;

// A page at a time, so that a long document is never all in memory at once: the
// sweep turns each into a grid a thousandth of its size and lets it go.
async function eachOfOurPages(
  bytes: Uint8Array,
  id: string,
  workspace: Workspace,
  take: (page: RasterImage) => void,
): Promise<Omit<OurPages, "pages">> {
  const pkg = openDocx(bytes);
  // The document's own alternatives are part of how a face is stood in.
  const measuring = substitutingMetrics(
    corpusFaces(),
    WORD_FALLBACK_FACES,
    readFaceAlternatives(pkg),
  );
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
  drawOurs: DrawOurs | null = null,
): Promise<Looks> {
  if (!existsSync(drawnPath)) return empty(id, "not drawn", "no pdf of Word's");

  const ours: PageGrid[] = [];
  let rest: Omit<OurPages, "pages">;
  try {
    if (drawOurs === null) {
      rest = await eachOfOurPages(bytes, id, workspace, (page) => ours.push(gridOf(page)));
    } else {
      const drawn = await drawOurs(bytes, id, workspace);
      for (const page of drawn.pages) ours.push(gridOf(page));
      rest = { facesStoodIn: drawn.facesStoodIn, asks: drawn.asks };
    }
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

  const pages: PageLooks[] = [];
  for (let at = 0; at < Math.max(ours.length, theirs.length); at += 1) {
    pages.push(differenceBetween(ours[at] ?? null, theirs[at] ?? null, tolerances));
  }
  const interesting = pages.reduce((sum, each) => sum + each.interesting, 0);
  const differing = pages.reduce((sum, each) => sum + each.differing, 0);

  return {
    id,
    outcome: "compared",
    pagesOurs: ours.length,
    pagesWord: theirs.length,
    facesStoodIn: rest.facesStoodIn,
    asks: rest.asks,
    interesting,
    differing,
    pages,
    detail: ours.length === theirs.length ? "" : "a different number of pages",
  };
}
