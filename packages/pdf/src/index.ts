// What `@docx-pages/pdf` promises: a laid-out document written out as a pdf, and
// the one convenience that opens and lays out a `.docx` before doing it.
//
// The surface is small on purpose, as the viewer's is. A caller states the faces
// the document draws in and where its pictures come from; everything else about
// the page has already been settled by `@docx-pages/core`, and this package
// decides nothing about where anything sits.

export type { PdfFont, PdfMetadata, WritePdfOptions } from "./document.js";
