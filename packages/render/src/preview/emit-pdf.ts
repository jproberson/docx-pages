import { readFileSync, writeFileSync } from "node:fs";

import {
  layOutDocument,
  lookupFontMetrics,
  writePdf,
  type MetricsResolver,
  type PdfFont,
} from "@docx-pages/core";

import { referenceCases, referenceFonts, suppliedFaces } from "../testing/cases.js";
import { readReferenceDocument } from "../testing/documents.js";

// A reference document written out by `@docx-pages/core`, to be looked at beside
// Word's own pdf of it. **The numbers are not enough on their own**: this project
// has already had a suite agree to the thousandth on every drawing in a document
// whose rendering was visibly wrong, and only the raster said so.
//
//   pnpm pdf <case-id> ours.pdf
//   pdftoppm -r 96 -png ours.pdf ours
//
// What to put beside it is the case's own Word pdf, whose path is printed below.
// Needs the reference manifest, and says so rather than writing nothing quietly.

const faces = suppliedFaces();
const metricsFor: MetricsResolver = (request) => lookupFontMetrics(request, faces);
const fonts: readonly PdfFont[] = referenceFonts().flatMap((font) =>
  font.filePath === null
    ? []
    : [
        {
          name: font.name,
          bold: font.bold,
          italic: font.italic,
          bytes: new Uint8Array(readFileSync(font.filePath)),
        },
      ],
);

const wanted = process.argv[2];
const out = process.argv[3] ?? "out.pdf";
const cases = referenceCases();

if (cases.length === 0) {
  console.log("no reference manifest, so there is nothing to write");
} else if (wanted === undefined) {
  console.log(`name one of: ${cases.map((each) => each.id).join(", ")}`);
} else {
  const each = cases.find((one) => one.id === wanted);
  if (each === undefined) throw new Error(`no reference case called ${wanted}`);

  const pkg = readReferenceDocument(each);
  const layout = layOutDocument(pkg, metricsFor);
  if (layout.kind !== "laid-out") throw new Error(JSON.stringify(layout.blocker));

  writeFileSync(
    out,
    writePdf(layout, { fonts, imageBytes: (part) => pkg.parts.get(part), metricsFor }),
  );
  console.log(`${each.id}: ${String(layout.pages.length)} pages -> ${out}`);
  console.log(`word drew it at: ${each.renderedPath ?? "nowhere measured"}`);
}
