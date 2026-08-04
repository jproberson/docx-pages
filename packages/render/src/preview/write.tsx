import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { renderToStaticMarkup } from "react-dom/server";

import { layOutDocument, lookupFontMetrics } from "@onepager/core";
import { imageResolver, OnePagerDocument, type FrameStyle } from "@onepager/viewer";

import { referenceCases, suppliedFaces, type ReferenceCase } from "../testing/cases.js";
import { readReferenceDocument } from "../testing/documents.js";

// The authored face is named first; whatever the machine has falls in behind it.
const FALLBACK_FONTS = "Calibri, Helvetica, Arial, sans-serif";

const OUTPUT_DIRECTORY = "samples/preview";

const document = (title: string, body: string): string =>
  `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>${title}</title>
    <style>
      body { margin: 0; padding: 24px; background: #6b6b6b; }
      [data-onepager-page] { background: #fff; box-shadow: 0 2px 16px rgb(0 0 0 / 40%); margin: 0 auto 24px; }
    </style>
  </head>
  <body>${body}</body>
</html>
`;

export function writePreview(each: ReferenceCase, frames: FrameStyle): string {
  const pkg = readReferenceDocument(each);
  const supplied = suppliedFaces();
  const layout = layOutDocument(pkg, (request) => lookupFontMetrics(request, supplied));
  if (layout.kind !== "laid-out") {
    throw new Error(`case ${each.id} is blocked: ${JSON.stringify(layout.blocker)}`);
  }

  const body = renderToStaticMarkup(
    <OnePagerDocument
      layout={layout}
      imageUrl={imageResolver(pkg)}
      frames={frames}
      fallbackFonts={FALLBACK_FONTS}
    />,
  );

  const path = resolve(OUTPUT_DIRECTORY, `${each.id}.html`);
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, document(`One-pager ${each.id}`, body));
  return path;
}

function main(): void {
  const frames: FrameStyle = process.argv.includes("--frames") ? "outlined" : "hidden";
  const cases = referenceCases();

  if (cases.length === 0) {
    process.stdout.write("no reference cases; point ONEPAGER_REFERENCE_MANIFEST at a manifest\n");
    return;
  }

  for (const each of cases) process.stdout.write(`${writePreview(each, frames)}\n`);
}

main();
