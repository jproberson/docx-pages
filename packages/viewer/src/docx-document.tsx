import { useEffect, useState, type ReactElement } from "react";

import {
  bestEffortMetrics,
  isAliasedSymbolFace,
  layOutDocument,
  openDocx,
  readFaceShapes,
  readFontFile,
  type FaceDefaults,
  type FallbackCharacter,
  type LaidOutDocument,
  type MissingGlyph,
  type Substitution,
  type SuppliedFace,
  type Unhonoured,
} from "@docx-pages/core";

import { imageResolver, type ImageResolver } from "./images.js";
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
  // What to draw for a document even best effort cannot lay out, which with a
  // full set of defaults in reach is a malformed file rather than a missing font.
  readonly blocked?: (reason: unknown) => ReactElement | null;
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

const registered = new Set<string>();

// Lets the browser draw a face under the name the layout measured it as: a
// supplied face under its own name, and a stood-in face under the name the
// document asked for, so what is painted is the very bytes that were measured.
// A runtime without the FontFace API paints whatever its styles find, at the
// measured widths; that is the one way a page here is right in its geometry and
// wrong on the screen, and it is the runtime's limit rather than a quiet choice.
function offerToBrowser(name: string, bold: boolean, italic: boolean, bytes: Uint8Array): void {
  if (typeof FontFace === "undefined" || typeof document === "undefined") return;
  const key = `${name.toLowerCase()}|${bold ? "b" : ""}${italic ? "i" : ""}`;
  if (registered.has(key)) return;
  registered.add(key);

  const face = new FontFace(name, bytes.slice().buffer, {
    weight: bold ? "bold" : "normal",
    style: italic ? "italic" : "normal",
  });
  document.fonts.add(face);
  face.load().catch(() => {
    // A face the browser refuses stays measured and unpainted, as above.
  });
}

const sameFace = (
  font: DocxFont,
  want: { readonly name: string; readonly bold: boolean; readonly italic: boolean },
): boolean =>
  font.name.trim().toLowerCase() === want.name.trim().toLowerCase() &&
  (font.bold ?? false) === want.bold &&
  (font.italic ?? false) === want.italic;

type Shown =
  | { readonly state: "opening" }
  | { readonly state: "blocked"; readonly reason: unknown }
  | {
      readonly state: "shown";
      readonly layout: LaidOutDocument;
      readonly imageUrl: ImageResolver;
      readonly aliasSymbolFaces: ReadonlySet<string> | null;
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
  const { source, fonts, defaults, defaultBytes, onReport, blocked } = props;

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

      for (const [at, font] of (fonts ?? []).entries()) {
        const face = supplied[at];
        if (face !== undefined) offerToBrowser(face.name, face.bold, face.italic, font.bytes);
      }
      for (const substitution of faces.substitutions()) {
        const stood = (defaultBytes ?? []).find((each) => sameFace(each, substitution.used));
        if (stood !== undefined && substitution.requested.name !== "") {
          offerToBrowser(
            substitution.requested.name,
            substitution.requested.bold,
            substitution.requested.italic,
            stood.bytes,
          );
        }
      }
      for (const each of defaultBytes ?? []) {
        offerToBrowser(each.name, each.bold ?? false, each.italic ?? false, each.bytes);
      }

      onReport?.({
        substitutions: faces.substitutions(),
        fallbackCharacters: faces.fallbackCharacters(),
        missingGlyphs: faces.missingGlyphs(),
        unhonoured: layout.unhonoured,
      });

      // Runs in a symbol face that was stood in for are drawn as what their
      // positions mean; one whose real face was supplied draws as written.
      const aliasedFaces = new Set(
        faces
          .substitutions()
          .filter((each) => isAliasedSymbolFace(each.requested.name))
          .map((each) => each.requested.name.trim().toLowerCase()),
      );
      setShown({
        state: "shown",
        layout,
        imageUrl: imageResolver(pkg, faces.metricsFor),
        aliasSymbolFaces: aliasedFaces.size > 0 ? aliasedFaces : null,
      });
    } catch (error) {
      setShown({ state: "blocked", reason: error });
    }
    // The report callback is deliberately not a dependency: a new closure for
    // the same bytes is not a new document.
  }, [source, fonts, defaults, defaultBytes]);

  if (shown.state === "opening") return null;
  if (shown.state === "blocked") {
    return blocked?.(shown.reason) ?? <pre>{describeReason(shown.reason)}</pre>;
  }
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
