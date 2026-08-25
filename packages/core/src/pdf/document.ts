import type { LaidOutDocument } from "../layout/document.js";
import type { MetricsResolver } from "../layout/stack.js";

import { contentOf } from "./content.js";
import { pdfFonts } from "./fonts.js";
import { pdfImages } from "./images.js";
import {
  pdfArray,
  pdfDictionary,
  pdfName,
  pdfNumber,
  pdfObjects,
  pdfStream,
  pdfString,
  type PdfEntries,
} from "./objects.js";

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
  // Which face inside the file, where the file holds more than one and the name it
  // answers to is not the name it is stored under.
  //
  // **The two come apart exactly where a face stands in for another.** A stand-in
  // is carried under the name the document asked for, since that is the name the
  // layout measured it as, and a collection asked for that name holds no such face:
  // measured over a corpus, every stand-in reached for was Cambria, whose file
  // holds `Cambria` and `Cambria Math` and nothing called `Aptos Display`. Left
  // out, the name it answers to is used, which is right for a file holding one
  // face and for a face supplied under its own name.
  readonly faceName?: string;
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
  // Symbol faces the layout stood in for, by lowercased name, exactly as the
  // viewer takes them. A run written in one holds positions in that face's own
  // page, and drawing it as the stand-in's letters would draw the wrong ones.
  readonly aliasSymbolFaces?: ReadonlySet<string>;
  /**
   * Told once about each picture that stood on a page and was drawn nowhere on it.
   *
   * **This backend draws fewer formats than the library reads.** A jpeg, a png and
   * a metafile are written into the file; a gif, a bmp, a tiff, an svg or a webp is
   * not, though the viewer draws all of those by handing them to the browser. So
   * `LaidOutDocument.unhonoured` is right to stay quiet about them, since nothing
   * about the document is unhonoured, and the writer is the only thing that knows.
   *
   * Found on 2026-08-24 in a corpus document holding one gif and one png, whose
   * png came out cell for cell and whose gif left a hole in the page that nothing
   * anywhere reported. A library that sells itself on saying what it did not do
   * cannot drop a picture in silence.
   */
  readonly undrawn?: (part: string) => void;
};

const infoOf = (metadata: PdfMetadata): PdfEntries => ({
  Title: metadata.title === undefined ? undefined : pdfString(metadata.title),
  Author: metadata.author === undefined ? undefined : pdfString(metadata.author),
  Subject: metadata.subject === undefined ? undefined : pdfString(metadata.subject),
  Keywords: metadata.keywords === undefined ? undefined : pdfString(metadata.keywords),
  Creator: metadata.creator === undefined ? undefined : pdfString(metadata.creator),
  Producer: metadata.producer === undefined ? undefined : pdfString(metadata.producer),
});

const NOTHING_STATED = (entries: PdfEntries): boolean =>
  Object.values(entries).every((value) => value === undefined);

/**
 * Writes a laid-out document out as a pdf.
 *
 * Decides nothing about where anything sits. Every position here was settled by
 * `layOutDocument` and measured against Word; this turns the one list
 * `drawablesOf` hands back into the operators that draw it, flipped onto a pdf's
 * own way up.
 *
 * The faces are the caller's to bring, and **a face the document draws in that
 * `fonts` does not supply refuses the document** rather than being stood in for.
 * A pdf carries the faces it draws in, and a page written in the wrong one is
 * wrong in a way nobody watching the file being made would see.
 */
export function writePdf(layout: LaidOutDocument, options: WritePdfOptions): Uint8Array {
  const objects = pdfObjects();
  const fonts = pdfFonts(options.fonts);
  const images = pdfImages({
    imageBytes: options.imageBytes,
    metricsFor: options.metricsFor,
    fonts,
    objects,
    ...(options.undrawn === undefined ? {} : { onUndrawn: options.undrawn }),
  });

  // The tree is reserved before the pages it holds, since a page names it and it
  // names every page: one of the two has to be referred to before it is written.
  const tree = objects.reserve();

  // Written before the resources, because writing a page is what settles which
  // faces were drawn in, what each of them was asked to draw, and which pictures
  // were reached for at all.
  const pages = layout.pages.map((page) => {
    const drawn = contentOf(layout, page, {
      fonts,
      images,
      aliasSymbolFaces: options.aliasSymbolFaces ?? null,
    });
    const stream = objects.add(pdfStream({}, drawn.bytes));

    return objects.add(
      pdfDictionary({
        Type: pdfName("Page"),
        Parent: tree,
        // Page size is stated per section, so a document whose sections differ
        // draws its pages at more than one size and each states its own.
        MediaBox: pdfArray([0, 0, drawn.page.widthPt, drawn.page.heightPt].map(pdfNumber)),
        Contents: stream,
      }),
    );
  });

  objects.put(
    tree,
    pdfDictionary({
      Type: pdfName("Pages"),
      Kids: pdfArray(pages),
      Count: pdfNumber(pages.length),
      // Inherited by every page under it, which is what lets one set of faces and
      // one set of pictures answer for the whole document: a logo on every page is
      // written once and named from each.
      Resources: pdfDictionary({
        Font: fonts.resources(objects),
        XObject: images.resources() ?? undefined,
      }),
    }),
  );

  const root = objects.add(pdfDictionary({ Type: pdfName("Catalog"), Pages: tree }));

  // Nothing here reads a clock, so no creation date is written: a writer that
  // touches no disk and no network should answer the same bytes on two runs of the
  // same document.
  const info = infoOf(options.metadata ?? {});
  return objects.bytes({
    root,
    ...(NOTHING_STATED(info) ? {} : { info: objects.add(pdfDictionary(info)) }),
  });
}
