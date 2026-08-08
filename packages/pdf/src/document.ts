import type { MetricsResolver } from "@docx-pages/core";

/**
 * A face handed in by bytes, named the way the document names it. The same shape
 * as the viewer's `DocxFont`, deliberately: an application drawing a document on
 * the screen and writing it out hands the one list to both.
 */
export type PdfFont = {
  readonly name: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly bytes: Uint8Array;
};

// What a reader shows about the file rather than anything drawn in it. Every part
// is the caller's to state: nothing here reads a clock, since a writer that
// touches no disk and no network should not answer differently on two runs.
export type PdfMetadata = {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
  readonly creator?: string;
  readonly producer?: string;
};

export type WritePdfOptions = {
  // Every face the document draws in, by the name it names. A face the document
  // asks for that is not here refuses the document: a pdf embeds what it draws,
  // and standing another face in its place would move every line on the page.
  readonly fonts: readonly PdfFont[];
  // The bytes of a drawing, by the part that holds it. The same resolver shape
  // the viewer takes, and answering `undefined` leaves the picture undrawn rather
  // than refusing the page.
  readonly imageBytes: (part: string) => Uint8Array | undefined;
  // What the layout measured with, which a metafile picture needs: a metafile
  // records text as a face and a string rather than as a drawing of one, so
  // playing it back asks the same resolver the layout asked.
  readonly metricsFor: MetricsResolver;
  readonly metadata?: PdfMetadata;
};
