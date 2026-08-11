import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { renderToStaticMarkup } from "react-dom/server";

import type { LaidOutDocument } from "@docx-pages/core";
import { Page, type ImageResolver } from "@docx-pages/viewer";

// A document's pages as html a browser can be told to photograph.
//
// The previewer's page fits itself to the window it is shown in, which is what
// makes it worth looking at beside a pdf viewer and what makes it useless here: a
// drawing to be compared with Word's has to be at a stated size and no other. So
// this writes the same pages with nothing around them, each in a box of whole
// pixels, and the browser is given a window exactly as tall as the lot.

// A browser draws in css pixels, ninety-six to the inch, and Word's pdf is read
// back at the same. So a point is a third again as many pixels on both sides.
const PIXELS_PER_POINT = 96 / 72;

const pixels = (twips: number): number => Math.round((twips / 20) * PIXELS_PER_POINT);

export type WrittenPages = {
  readonly path: string;
  readonly widthPx: number;
  readonly heightPx: number;
  // Where each page starts down the drawing, and how big it is. Every box is a
  // whole number of pixels, so no page's top has to be worked out from the
  // fractional heights of the ones above it.
  readonly pages: readonly {
    readonly topPx: number;
    readonly widthPx: number;
    readonly heightPx: number;
  }[];
};

const html = (stylesheet: string, body: string): string =>
  `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <link rel="stylesheet" href="${stylesheet}" />
    <style>
      html, body { margin: 0; padding: 0; background: #fff; }
      [data-docx-sheet] { overflow: hidden; background: #fff; }
    </style>
  </head>
  <body>${body}</body>
</html>
`;

/**
 * The pages `from` up to `to` of one document, stacked with nothing between them.
 * A run of pages rather than a document because a browser is asked for one
 * photograph the size of its window, and a window as tall as a long document is a
 * window no browser will give.
 */
export function writePages(
  path: string,
  layout: LaidOutDocument,
  imageUrl: ImageResolver,
  stylesheet: string,
  from: number,
  to: number,
): WrittenPages {
  const wanted = layout.pages.slice(from, to);

  const pages = wanted.map((page) => ({
    widthPx: pixels(page.geometry.widthTwips),
    heightPx: pixels(page.geometry.heightTwips),
  }));

  const body = wanted
    .map((page, at) => {
      const box = pages[at];
      const drawn = renderToStaticMarkup(
        <Page layout={layout} page={page} imageUrl={imageUrl} fallbackFonts="serif" />,
      );
      return (
        `<div data-docx-sheet="${String(page.index)}" style="width:${String(box?.widthPx ?? 0)}px;` +
        `height:${String(box?.heightPx ?? 0)}px">${drawn}</div>`
      );
    })
    .join("");

  mkdirSync(dirname(resolve(path)), { recursive: true });
  writeFileSync(resolve(path), html(stylesheet, body));

  let topPx = 0;
  const placed = pages.map((box) => {
    const at = { topPx, ...box };
    topPx += box.heightPx;
    return at;
  });

  return {
    path: resolve(path),
    widthPx: Math.max(...pages.map((each) => each.widthPx), 1),
    heightPx: Math.max(topPx, 1),
    pages: placed,
  };
}
