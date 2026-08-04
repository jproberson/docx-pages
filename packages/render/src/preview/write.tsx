import { mkdirSync, writeFileSync } from "node:fs";
import { basename, dirname, resolve } from "node:path";
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

const FRAME_STYLES: readonly FrameStyle[] = ["hidden", "outlined"];

const pageName = (id: string, frames: FrameStyle): string =>
  frames === "outlined" ? `${id}.frames.html` : `${id}.html`;

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

  const path = resolve(OUTPUT_DIRECTORY, pageName(each.id, frames));
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, document(`One-pager ${each.id}`, body));
  return path;
}

const escaped = (value: string): string =>
  value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const titleOf = (each: ReferenceCase): string =>
  `${each.id}. ${basename(each.documentPath).replace(/\.docx$/i, "")}`;

// One page holding every document, so a change can be seen in all of them without
// hunting for files. Each one is drawn into its own frame, which keeps the
// document's own styling away from the browser's.
function writeBrowser(cases: readonly ReferenceCase[]): string {
  const options = cases
    .map((each) => `<option value="${escaped(each.id)}">${escaped(titleOf(each))}</option>`)
    .join("");

  const body = `
    <header>
      <label for="document">Document</label>
      <select id="document">${options}</select>
      <label for="frames"><input type="checkbox" id="frames" /> Outline text boxes and shapes</label>
    </header>
    <iframe id="page" title="one-pager"></iframe>
    <script>
      const chooser = document.getElementById("document");
      const frames = document.getElementById("frames");
      const page = document.getElementById("page");

      const show = () => {
        const suffix = frames.checked ? ".frames" : "";
        page.src = chooser.value + suffix + ".html";
        location.hash = chooser.value;
      };

      const restore = () => {
        const wanted = location.hash.slice(1);
        const known = [...chooser.options].some((each) => each.value === wanted);
        if (!known || chooser.value === wanted) return;
        chooser.value = wanted;
        show();
      };

      chooser.addEventListener("change", show);
      frames.addEventListener("change", show);
      window.addEventListener("hashchange", restore);

      const wanted = location.hash.slice(1);
      if ([...chooser.options].some((each) => each.value === wanted)) chooser.value = wanted;
      show();
    </script>`;

  const path = resolve(OUTPUT_DIRECTORY, "index.html");
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, browser(body));
  return path;
}

const browser = (body: string): string =>
  `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>One-pagers</title>
    <style>
      html, body { height: 100%; margin: 0; }
      body { display: flex; flex-direction: column; font: 14px system-ui, sans-serif; }
      header { display: flex; gap: 16px; align-items: center; padding: 10px 16px; background: #222; color: #eee; }
      select { font: inherit; padding: 2px 6px; }
      label { display: flex; gap: 6px; align-items: center; }
      iframe { flex: 1; width: 100%; border: 0; }
    </style>
  </head>
  <body>${body}</body>
</html>
`;

function main(): void {
  const cases = referenceCases();

  if (cases.length === 0) {
    process.stdout.write("no reference cases; point ONEPAGER_REFERENCE_MANIFEST at a manifest\n");
    return;
  }

  for (const each of cases) {
    for (const frames of FRAME_STYLES) process.stdout.write(`${writePreview(each, frames)}\n`);
  }
  process.stdout.write(`${writeBrowser(cases)}\n`);
}

main();
