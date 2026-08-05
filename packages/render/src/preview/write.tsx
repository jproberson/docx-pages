import { copyFileSync, existsSync, mkdirSync, writeFileSync } from "node:fs";
import { basename, dirname, relative, resolve } from "node:path";
import { renderToStaticMarkup } from "react-dom/server";

import {
  layOutDocument,
  lookupFontMetrics,
  type FaceRequest,
  type MetricsLookup,
  type SuppliedFace,
} from "@onepager/core";
import { imageResolver, OnePagerDocument, type FrameStyle } from "@onepager/viewer";

import {
  authoredFaces,
  authoredFonts,
  referenceCases,
  referenceFonts,
  suppliedFaces,
  type ReferenceCase,
} from "../testing/cases.js";
import { readReferenceDocument } from "../testing/documents.js";

// Whatever the machine has falls in behind the face the document names, which the
// stylesheet beside the page supplies.
const FALLBACK_FONTS = "Calibri, Helvetica, Arial, sans-serif";

const OUTPUT_DIRECTORY = "samples/preview";
const FONT_DIRECTORY = "fonts";

// A file the browser can load, named for a face rather than for where it came from.
type FaceFile = {
  readonly name: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly filePath: string;
};

// A document can be drawn in the faces Word fell back on, which is what its own pdf
// shows, or in the ones it was authored in. Each is measured with the same faces it
// is drawn in: holding a run to a width some other face was measured at is what
// crushes the spaces out of it.
type FaceSet = {
  readonly id: string;
  readonly label: string;
  readonly faces: readonly SuppliedFace[];
  readonly files: readonly FaceFile[];
};

const wordFaceSet = (): FaceSet => ({
  id: "word",
  label: "As Word rendered it",
  faces: suppliedFaces(),
  files: referenceFonts().flatMap((font) =>
    font.filePath === null ? [] : [{ ...font, filePath: font.filePath }],
  ),
});

function authoredFaceSet(): FaceSet | null {
  const authored = authoredFonts();
  if (authored.length === 0) return null;

  const named = new Set(authored.map((each) => each.name.toLowerCase()));
  const rest = referenceFonts().flatMap((font) =>
    font.filePath === null || named.has(font.name.toLowerCase())
      ? []
      : [{ ...font, filePath: font.filePath }],
  );

  return {
    id: "authored",
    label: "In the authored face",
    faces: authoredFaces(),
    files: [...authored, ...rest],
  };
}

const stylesheetOf = (set: FaceSet): string => `fonts.${set.id}.css`;

// A line is measured with the widths of the face it is laid out in. Unless the
// browser can draw that same face the text is squeezed back to those widths against
// whatever it has instead, which crushes the spaces between words.
function writeFonts(set: FaceSet): string {
  mkdirSync(resolve(OUTPUT_DIRECTORY, FONT_DIRECTORY), { recursive: true });

  const rules = set.files.flatMap((face) => {
    if (!existsSync(face.filePath)) return [];

    const file = basename(face.filePath).replace(/\s+/g, "-");
    copyFileSync(face.filePath, resolve(OUTPUT_DIRECTORY, FONT_DIRECTORY, file));

    return [
      `@font-face {
  font-family: "${face.name}";
  font-weight: ${face.bold ? "bold" : "normal"};
  font-style: ${face.italic ? "italic" : "normal"};
  src: url("${FONT_DIRECTORY}/${file}");
}`,
    ];
  });

  const path = resolve(OUTPUT_DIRECTORY, stylesheetOf(set));
  writeFileSync(path, `${rules.join("\n")}\n`);
  return path;
}

const document = (title: string, stylesheet: string, body: string): string =>
  `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>${title}</title>
    <link rel="stylesheet" href="${stylesheet}" />
    <style>
      body { margin: 0; padding: 24px; background: #6b6b6b; }
      [data-onepager-page] { background: #fff; box-shadow: 0 2px 16px rgb(0 0 0 / 40%); margin: 0 auto 24px; }
    </style>
  </head>
  <body>${body}
    <script>
      // Word's own pdf is shown fitted to the width it is given, so a page beside
      // it has to be fitted the same way for the two to be worth comparing.
      const first = document.querySelector("[data-onepager-page]");
      const pageWidth = first.offsetWidth;
      const fit = () => {
        document.body.style.zoom = (document.documentElement.clientWidth - 48) / pageWidth;
      };
      window.addEventListener("resize", fit);
      fit();
    </script>
  </body>
</html>
`;

const FRAME_STYLES: readonly FrameStyle[] = ["hidden", "outlined"];

const pageName = (id: string, set: FaceSet, frames: FrameStyle): string =>
  `${id}.${set.id}${frames === "outlined" ? ".frames" : ""}.html`;

export function writePreview(each: ReferenceCase, set: FaceSet, frames: FrameStyle): string {
  const pkg = readReferenceDocument(each);
  const metricsFor = (request: FaceRequest): MetricsLookup => lookupFontMetrics(request, set.faces);
  const layout = layOutDocument(pkg, metricsFor);
  if (layout.kind !== "laid-out") {
    throw new Error(`case ${each.id} is blocked: ${JSON.stringify(layout.blocker)}`);
  }

  const body = renderToStaticMarkup(
    <OnePagerDocument
      layout={layout}
      imageUrl={imageResolver(pkg, metricsFor)}
      frames={frames}
      fallbackFonts={FALLBACK_FONTS}
    />,
  );

  const path = resolve(OUTPUT_DIRECTORY, pageName(each.id, set, frames));
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, document(`One-pager ${each.id}`, stylesheetOf(set), body));
  return path;
}

const escaped = (value: string): string =>
  value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const titleOf = (each: ReferenceCase): string =>
  `${each.id}. ${basename(each.documentPath).replace(/\.docx$/i, "")}`;

// Word's own pdf is the only thing worth judging a rendering against, so it sits
// beside it rather than in another window.
const wordPageOf = (each: ReferenceCase): string =>
  each.renderedPath === null ? "" : relative(OUTPUT_DIRECTORY, each.renderedPath);

// One page holding every document, so a change can be seen in all of them without
// hunting for files. Each one is drawn into its own frame, which keeps the
// document's own styling away from the browser's.
function writeBrowser(cases: readonly ReferenceCase[], sets: readonly FaceSet[]): string {
  const options = cases
    .map(
      (each) =>
        `<option value="${escaped(each.id)}" data-word="${escaped(wordPageOf(each))}">${escaped(titleOf(each))}</option>`,
    )
    .join("");

  const faces = sets
    .map((set) => `<option value="${escaped(set.id)}">${escaped(set.label)}</option>`)
    .join("");

  const body = `
    <header>
      <label for="document">Document</label>
      <select id="document">${options}</select>
      <label for="face">Face</label>
      <select id="face">${faces}</select>
      <label for="beside"><input type="checkbox" id="beside" checked /> Word's own pdf beside it</label>
      <label for="frames"><input type="checkbox" id="frames" /> Outline text boxes and shapes</label>
    </header>
    <main>
      <iframe id="page" title="ours"></iframe>
      <iframe id="word" title="word"></iframe>
    </main>
    <script>
      const chooser = document.getElementById("document");
      const face = document.getElementById("face");
      const frames = document.getElementById("frames");
      const beside = document.getElementById("beside");
      const page = document.getElementById("page");
      const word = document.getElementById("word");

      const show = () => {
        const suffix = frames.checked ? ".frames" : "";
        page.src = chooser.value + "." + face.value + suffix + ".html";
        const pdf = chooser.selectedOptions[0].dataset.word;
        word.hidden = !beside.checked || pdf === "";
        word.src = word.hidden ? "about:blank" : pdf + "#view=FitH&toolbar=0";
        location.hash = chooser.value;
      };

      const restore = () => {
        const wanted = location.hash.slice(1);
        const known = [...chooser.options].some((each) => each.value === wanted);
        if (!known || chooser.value === wanted) return;
        chooser.value = wanted;
        show();
      };

      for (const control of [chooser, face, frames, beside]) {
        control.addEventListener("change", show);
      }
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
      main { flex: 1; display: flex; min-height: 0; }
      iframe { flex: 1; min-width: 0; height: 100%; border: 0; }
      iframe + iframe { border-left: 2px solid #222; }
      iframe[hidden] { display: none; }
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

  const authored = authoredFaceSet();
  const sets = authored === null ? [wordFaceSet()] : [wordFaceSet(), authored];

  for (const set of sets) {
    process.stdout.write(`${writeFonts(set)}\n`);
    for (const each of cases) {
      for (const frames of FRAME_STYLES)
        process.stdout.write(`${writePreview(each, set, frames)}\n`);
    }
  }
  process.stdout.write(`${writeBrowser(cases, sets)}\n`);
}

main();
