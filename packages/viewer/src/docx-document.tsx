import { useEffect, useState, type ReactElement } from "react";

import {
  bestEffortMetrics,
  isAliasedSymbolFace,
  layOutDocument,
  openDocx,
  readFaceShapes,
  readFontFile,
  unshowableIn,
  writePdf,
  type FaceDefaults,
  type FallbackCharacter,
  type LaidOutDocument,
  type MissingGlyph,
  type PdfFont,
  type Substitution,
  type SuppliedFace,
  type Unhonoured,
  type Unshowable,
} from "@docx-pages/core";

import { imageResolver, type ImageResolver } from "./images.js";
import { offerToBrowser } from "./offer-face.js";
import { Document, type FrameStyle } from "./page.js";

// A face handed in by bytes, named the way the document names it. Everything
// else about it is read out of the bytes.
export type DocxFont = {
  readonly name: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly bytes: Uint8Array;
};

// Everything that keeps a best-effort page honest, handed out after layout: a
// page drawn over any of these is not the page Word would draw, and an
// application that shows one can say so.
export type DocxRenderReport = {
  readonly substitutions: readonly Substitution[];
  readonly fallbackCharacters: readonly FallbackCharacter[];
  readonly missingGlyphs: readonly MissingGlyph[];
  readonly unhonoured: readonly Unhonoured[];
  // Where the pages came out wrong on their own terms, which is not the same list as
  // `unhonoured`: that says what the document asked for and did not get, most of it
  // invisible, and this says the page draws text off the sheet, above its own top, or
  // over other text, which no document can ask for. Empty is a page worth showing.
  readonly unshowable: readonly Unshowable[];
};

export type DocxDocumentProps = {
  readonly source: Uint8Array | ArrayBuffer;
  // What answers for a face the document names that `fonts` does not supply.
  // Nothing here reaches for a font of its own: this module names no font
  // package, so a consumer supplying its own faces never installs one.
  // `@docx-pages/viewer/pack` fills this in from `@docx-pages/fonts`.
  readonly defaults: FaceDefaults;
  // Bytes for the faces `defaults` measures with, so a face stood in for is
  // painted with the very bytes it was measured with. A default face whose
  // bytes are missing is measured here and painted by whatever the browser
  // finds, which is the one way a page is right in its geometry and wrong on
  // the screen.
  readonly defaultBytes?: readonly DocxFont[];
  // Faces supplied for exactness. Every face the document names that is not
  // here falls to `defaults`, and the report says which.
  readonly fonts?: readonly DocxFont[];
  readonly onReport?: (report: DocxRenderReport) => void;
  // Handed out once the document is drawn, with a function that writes the very
  // page on the screen out as a pdf.
  //
  // **It writes the layout that was drawn, not a second one.** The same measured
  // page, the same faces it was painted with, and the same `drawablesOf` the
  // component walks, so a file cannot come out saying something the screen did
  // not: there is nothing between them to disagree.
  readonly onReady?: (writePdfOfPages: () => Uint8Array) => void;
  // What to draw for a document even best effort cannot lay out, which with a
  // full set of defaults in reach is a malformed file rather than a missing font.
  readonly blocked?: (reason: unknown) => ReactElement | null;
  /**
   * What to draw instead of a page that came out unusable, where the caller would
   * rather show nothing than show that.
   *
   * **Laying out is not the same as succeeding.** A document that lays out perfectly
   * well can still put text off the sheet, above the top of its own page, or over
   * other text, and a preview showing one of those is worse than a preview that says
   * it cannot show it: over the corpus, 11 clean documents of 580 draw one, and the
   * raster says not one of those pages is the page Word drew. So the decision is made
   * before anything is painted rather than reported after it.
   *
   * Left out, the pages are drawn whatever they say about themselves, which is what
   * this component did before the check existed. Answering `null` from it draws
   * nothing at all.
   */
  readonly unshowable?: (found: readonly Unshowable[]) => ReactElement | null;
  readonly scale?: number;
  readonly frames?: FrameStyle;
  readonly className?: string;
};

const suppliedFrom = (font: DocxFont): SuppliedFace => {
  const read = readFontFile(font.bytes);
  return {
    name: font.name,
    bold: font.bold ?? false,
    italic: font.italic ?? false,
    metrics: read.metrics,
    advances: read.advances,
    sansSerif: read.sansSerif,
  };
};

const sameFace = (
  font: DocxFont,
  want: { readonly name: string; readonly bold: boolean; readonly italic: boolean },
): boolean =>
  font.name.trim().toLowerCase() === want.name.trim().toLowerCase() &&
  (font.bold ?? false) === want.bold &&
  (font.italic ?? false) === want.italic;

/**
 * Every face the page is painted with, which is every face a file of it must
 * carry.
 *
 * **A face stood in for goes under the name the document asked for**, and not
 * under its own. That is the name the layout measured it as and the name the
 * browser is offered it under, so a file carrying it under the stand-in's own name
 * would be a page written in a face nothing here ever measured.
 *
 * The order is the order a face is looked for in: what the caller supplied for
 * exactness first, so a document naming a face that is both supplied and defaulted
 * is drawn in the supplied one, on the screen and in the file alike.
 */
export function facesPaintedWith(
  fonts: readonly DocxFont[],
  substitutions: readonly Substitution[],
  defaultBytes: readonly DocxFont[],
): readonly PdfFont[] {
  const named = (font: DocxFont): PdfFont => ({
    name: font.name,
    bold: font.bold ?? false,
    italic: font.italic ?? false,
    bytes: font.bytes,
  });

  const painted: PdfFont[] = fonts.map(named);

  for (const substitution of substitutions) {
    const stood = defaultBytes.find((each) => sameFace(each, substitution.used));
    // A stand-in nothing handed the bytes of cannot be carried into a file: on a
    // screen it is painted by whatever the browser finds, and a pdf has no such
    // thing to fall back on. The document is refused rather than written wrongly,
    // which is the writer's own bargain and not one to work around here.
    if (stood === undefined || substitution.requested.name === "") continue;
    painted.push({
      name: substitution.requested.name,
      bold: substitution.requested.bold,
      italic: substitution.requested.italic,
      bytes: stood.bytes,
    });
  }

  return [...painted, ...defaultBytes.map(named)];
}

type Shown =
  | { readonly state: "opening" }
  | { readonly state: "blocked"; readonly reason: unknown }
  | {
      readonly state: "shown";
      readonly layout: LaidOutDocument;
      readonly imageUrl: ImageResolver;
      readonly aliasSymbolFaces: ReadonlySet<string> | null;
      readonly unshowable: readonly Unshowable[];
    };

/**
 * Draws a `.docx` from its bytes alone: opens it, resolves every face it names
 * through whatever `fonts` supplies and `defaults` behind that, lays it out, and
 * paints it with the same bytes it was measured with. Never quiet: `onReport`
 * receives every stand-in, borrowed character, missing glyph and unread feature,
 * and a page drawn over any of them is not the page Word would draw.
 *
 * Exactness is `fonts`: a face supplied there is used as given, and a document
 * whose faces are all supplied lays out as `layOutDocument` would have laid it
 * out with no defaults consulted at all.
 *
 * The faces themselves are the caller's to bring. For the pack's, import this
 * component from `@docx-pages/viewer/pack` instead, which is the only module
 * here that names `@docx-pages/fonts`.
 */
export function DocxDocument(props: DocxDocumentProps): ReactElement | null {
  const [shown, setShown] = useState<Shown>({ state: "opening" });
  const { source, fonts, defaults, defaultBytes, onReport, onReady, blocked } = props;
  const refuse = props.unshowable;

  // Laying out is synchronous once the faces are in hand, which is the whole
  // reason the pack lives behind `@docx-pages/viewer/pack`: fetching it was the
  // only thing here that ever had to be waited for.
  useEffect(() => {
    try {
      const bytes = source instanceof Uint8Array ? source : new Uint8Array(source);
      const pkg = openDocx(bytes);

      const supplied = (fonts ?? []).map(suppliedFrom);
      const faces = bestEffortMetrics(supplied, defaults, readFaceShapes(pkg));
      const layout = layOutDocument(pkg, faces);

      if (layout.kind !== "laid-out") {
        setShown({ state: "blocked", reason: layout.blocker });
        return;
      }

      // Gathered once and used twice, for the screen and for the file, so that the
      // two cannot be handed different faces.
      const drawnWith = facesPaintedWith(fonts ?? [], faces.substitutions(), defaultBytes ?? []);
      for (const face of drawnWith) {
        offerToBrowser(face.name, face.bold ?? false, face.italic ?? false, face.bytes);
      }

      const unshowable = unshowableIn(layout);

      onReport?.({
        substitutions: faces.substitutions(),
        fallbackCharacters: faces.fallbackCharacters(),
        missingGlyphs: faces.missingGlyphs(),
        unhonoured: layout.unhonoured,
        unshowable,
      });

      // Runs in a symbol face that was stood in for are drawn as what their
      // positions mean; one whose real face was supplied draws as written.
      const aliasedFaces = new Set(
        faces
          .substitutions()
          .filter((each) => isAliasedSymbolFace(each.requested.name))
          .map((each) => each.requested.name.trim().toLowerCase()),
      );
      const aliasSymbolFaces = aliasedFaces.size > 0 ? aliasedFaces : null;

      onReady?.(() =>
        writePdf(layout, {
          fonts: drawnWith,
          imageBytes: (part) => pkg.parts.get(part),
          metricsFor: faces.metricsFor,
          ...(aliasSymbolFaces === null ? {} : { aliasSymbolFaces }),
        }),
      );

      setShown({
        state: "shown",
        layout,
        imageUrl: imageResolver(pkg, faces.metricsFor),
        aliasSymbolFaces,
        unshowable,
      });
    } catch (error) {
      setShown({ state: "blocked", reason: error });
    }
    // Neither callback is a dependency: a new closure for the same bytes is not a
    // new document, and laying one out again to hand back the same page would cost
    // the whole of the work twice.
  }, [source, fonts, defaults, defaultBytes]);

  if (shown.state === "opening") return null;
  if (shown.state === "blocked") {
    return blocked?.(shown.reason) ?? <pre>{describeReason(shown.reason)}</pre>;
  }
  // Asked before a page is painted, so a caller that would rather show nothing never
  // has the broken page on the screen first.
  if (refuse !== undefined && shown.unshowable.length > 0) return refuse(shown.unshowable);
  return (
    <Document
      layout={shown.layout}
      imageUrl={shown.imageUrl}
      {...(shown.aliasSymbolFaces === null ? {} : { aliasSymbolFaces: shown.aliasSymbolFaces })}
      {...(props.scale === undefined ? {} : { scale: props.scale })}
      {...(props.frames === undefined ? {} : { frames: props.frames })}
      {...(props.className === undefined ? {} : { className: props.className })}
    />
  );
}

function describeReason(reason: unknown): string {
  if (reason instanceof Error) return reason.message;
  try {
    return JSON.stringify(reason);
  } catch {
    return String(reason);
  }
}
