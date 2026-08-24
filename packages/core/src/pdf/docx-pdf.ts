import { openDocx } from "../docx/package.js";
import { DocxPagesError } from "../errors.js";
import { layOutDocument } from "../layout/document.js";
import { readSuppliedFace } from "../layout/font-file.js";
import { lookupFontMetrics, type SuppliedFace } from "../layout/font-metrics.js";
import type { MetricsResolver } from "../layout/stack.js";

import { writePdf, type PdfFont, type PdfMetadata } from "./document.js";

export type PdfOfDocxOptions = {
  // Every face the document draws in. There is no falling back here: a document
  // naming a face this does not carry is refused, which is the mode every
  // measurement in this project is made in.
  readonly fonts: readonly PdfFont[];
  readonly metadata?: PdfMetadata;
};

const suppliedFrom = (font: PdfFont): SuppliedFace =>
  readSuppliedFace(
    font.bytes,
    { name: font.name, bold: font.bold ?? false, italic: font.italic ?? false },
    { inFile: font.name },
  );

/**
 * Opens a `.docx`, lays it out over the faces supplied, and writes it out as a
 * pdf. The whole of the usual path, as `DocxDocument` is for the viewer.
 *
 * The faces are read once and used twice, to measure the document and to draw it,
 * so the page written is the page that was measured. **Nothing is stood in for**:
 * the resolver here is the bare one, so a document naming a face `fonts` does not
 * carry is refused at layout rather than measured against something else. That is
 * the mode every measurement in this project is made in, and for a file being
 * written rather than looked at it is the only honest one.
 *
 * A caller who would rather have a page than a refusal lays the document out
 * themselves, with `bestEffortMetrics` or `substitutingMetrics`, hands the bytes
 * of whatever stood in, and calls `writePdf`. Then the substitution is one they
 * made and can report.
 */
export function pdfOfDocx(source: Uint8Array | ArrayBuffer, options: PdfOfDocxOptions): Uint8Array {
  const bytes = source instanceof Uint8Array ? source : new Uint8Array(source);
  const pkg = openDocx(bytes);

  const supplied = options.fonts.map(suppliedFrom);
  const metricsFor: MetricsResolver = (request) => lookupFontMetrics(request, supplied);

  const layout = layOutDocument(pkg, metricsFor);
  if (layout.kind !== "laid-out") {
    throw new DocxPagesError({
      code: "layout-blocked",
      message: "the document could not be laid out over the faces supplied",
      at: "pdf/docx-pdf.pdfOfDocx",
      context: { blocker: layout.blocker.kind, detail: JSON.stringify(layout.blocker) },
    });
  }

  return writePdf(layout, {
    fonts: options.fonts,
    imageBytes: (part) => pkg.parts.get(part),
    metricsFor,
    ...(options.metadata === undefined ? {} : { metadata: options.metadata }),
  });
}
