import { readFileSync } from "node:fs";

import { facesUsed, lookupFontMetrics, openDocx, type UsedFace } from "@docx-pages/core";

import { suppliedFaces } from "../testing/cases.js";

export type FaceReport = {
  readonly face: UsedFace;
  readonly known: boolean;
  // Whether the face can be measured as well as placed: a face resolved only by
  // its metrics still cannot break a line, since that needs its advances.
  readonly measurable: boolean;
};

const styleOf = (face: UsedFace): string =>
  [face.bold ? "bold" : "", face.italic ? "italic" : ""].filter(Boolean).join(" ");

export function reportFaces(documentPath: string): readonly FaceReport[] {
  const pkg = openDocx(new Uint8Array(readFileSync(documentPath)));
  const supplied = suppliedFaces();

  return facesUsed(pkg).map((face) => {
    if (face.name === null) return { face, known: false, measurable: false };
    const lookup = lookupFontMetrics(
      { name: face.name, bold: face.bold, italic: face.italic },
      supplied,
    );
    return {
      face,
      known: lookup.kind === "found",
      measurable: lookup.kind === "found" && lookup.advances.kind === "advances",
    };
  });
}

const sizesOf = (face: UsedFace): string =>
  face.sizesPt.map((size) => `${String(size)}pt`).join(" ");

export const describeFaces = (reports: readonly FaceReport[]): string =>
  reports
    .map(({ face, known, measurable }) => {
      const mark = !known ? "MISSING " : measurable ? "ok      " : "no widths ";
      const style = styleOf(face);
      return `${mark}${face.name ?? "(no font named)"}${style === "" ? "" : ` (${style})`} at ${sizesOf(face)}`;
    })
    .join("\n");

// Run against a document before adding it to the manifest: every face it needs is
// listed at once, rather than one blocked layout at a time.
function main(): void {
  const [documentPath] = process.argv.slice(2);
  if (documentPath === undefined) {
    process.stderr.write("usage: node packages/render/dist/fonts/report.js <path to .docx>\n");
    process.exitCode = 1;
    return;
  }

  const reports = reportFaces(documentPath);
  process.stdout.write(`${describeFaces(reports)}\n`);

  const missing = reports.filter((report) => !report.measurable).length;
  if (missing > 0) {
    process.stdout.write(
      `\n${String(missing)} of ${String(reports.length)} faces cannot be measured; ` +
        `add them to the manifest's "fonts" before laying this document out.\n`,
    );
  }
}

if (process.argv[1] !== undefined && process.argv[1].endsWith("report.js")) main();
