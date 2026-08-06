import { useEffect, useState, type ReactElement } from "react";

import {
  bestEffortMetrics,
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
import { fontUrl, METRIC_TWINS, PACK_FACES, type PackFace } from "@docx-pages/fonts";

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
  // Faces supplied for exactness. Every face the document names that is not
  // here falls to the pack's twins and defaults, and the report says which.
  readonly fonts?: readonly DocxFont[];
  // Stands in for the pack, for a runtime whose fetch cannot reach the pack's
  // own files: hand `defaultFacesFromDisk()` in from `@docx-pages/fonts/node`.
  readonly defaults?: FaceDefaults;
  readonly onReport?: (report: DocxRenderReport) => void;
  // What to draw for a document even best effort cannot lay out, which with the
  // pack in reach is a malformed file rather than a missing font.
  readonly blocked?: (reason: unknown) => ReactElement | null;
  readonly scale?: number;
  readonly frames?: FrameStyle;
  readonly className?: string;
};

type PackEntry = {
  readonly face: PackFace;
  readonly bytes: Uint8Array;
  readonly supplied: SuppliedFace;
};

// The pack is fetched once per page, not once per document.
let packOnce: Promise<readonly PackEntry[]> | null = null;

const loadPack = (): Promise<readonly PackEntry[]> => {
  packOnce ??= Promise.all(
    PACK_FACES.map(async (face) => {
      const response = await fetch(fontUrl(face));
      if (!response.ok)
        throw new Error(`the font at ${fontUrl(face).href} came back ${String(response.status)}`);
      const bytes = new Uint8Array(await response.arrayBuffer());
      const read = readFontFile(bytes);
      return {
        face,
        bytes,
        supplied: {
          name: face.name,
          bold: face.bold,
          italic: face.italic,
          metrics: read.metrics,
          advances: read.advances,
          sansSerif: read.sansSerif,
        },
      };
    }),
  );
  return packOnce;
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

type Shown =
  | { readonly state: "opening" }
  | { readonly state: "blocked"; readonly reason: unknown }
  | { readonly state: "shown"; readonly layout: LaidOutDocument; readonly imageUrl: ImageResolver };

/**
 * Draws a `.docx` from its bytes alone: opens it, resolves every face it names
 * through whatever `fonts` supplies and the pack's twins behind that, lays it
 * out, and paints it with the same bytes it was measured with. Never quiet:
 * `onReport` receives every stand-in, borrowed character, missing glyph and
 * unread feature, and a page drawn over any of them is not the page Word would
 * draw.
 *
 * Exactness is `fonts`: a face supplied there is used as given, and a document
 * whose faces are all supplied lays out as `layOutDocument` would have laid it
 * out with no pack at all.
 */
export function DocxDocument(props: DocxDocumentProps): ReactElement | null {
  const [shown, setShown] = useState<Shown>({ state: "opening" });
  const { source, fonts, defaults, onReport, blocked } = props;

  useEffect(() => {
    let stale = false;

    const open = async (): Promise<void> => {
      try {
        const bytes = source instanceof Uint8Array ? source : new Uint8Array(source);
        const pkg = openDocx(bytes);

        const pack = defaults === undefined ? await loadPack() : null;
        const packDefaults: FaceDefaults =
          defaults ??
          ({
            faces: (pack ?? []).map((entry) => entry.supplied),
            twins: METRIC_TWINS,
            sansSerif: "Liberation Sans",
            serif: "Liberation Serif",
            monospace: "Liberation Mono",
            lastResort: "Caladea",
          } satisfies FaceDefaults);

        const supplied = (fonts ?? []).map(suppliedFrom);
        const faces = bestEffortMetrics(supplied, packDefaults, readFaceShapes(pkg));
        const layout = layOutDocument(pkg, faces);
        if (stale) return;

        if (layout.kind !== "laid-out") {
          setShown({ state: "blocked", reason: layout.blocker });
          return;
        }

        for (const [at, font] of (fonts ?? []).entries()) {
          const face = supplied[at];
          if (face !== undefined) offerToBrowser(face.name, face.bold, face.italic, font.bytes);
        }
        for (const substitution of faces.substitutions()) {
          const entry = pack?.find(
            (each) =>
              each.face.name === substitution.used.name &&
              each.face.bold === substitution.used.bold &&
              each.face.italic === substitution.used.italic,
          );
          if (entry !== undefined && substitution.requested.name !== "") {
            offerToBrowser(
              substitution.requested.name,
              substitution.requested.bold,
              substitution.requested.italic,
              entry.bytes,
            );
          }
        }
        for (const entry of pack ?? []) {
          offerToBrowser(entry.face.name, entry.face.bold, entry.face.italic, entry.bytes);
        }

        onReport?.({
          substitutions: faces.substitutions(),
          fallbackCharacters: faces.fallbackCharacters(),
          missingGlyphs: faces.missingGlyphs(),
          unhonoured: layout.unhonoured,
        });
        setShown({ state: "shown", layout, imageUrl: imageResolver(pkg, faces.metricsFor) });
      } catch (error) {
        if (!stale) setShown({ state: "blocked", reason: error });
      }
    };

    void open();
    return () => {
      stale = true;
    };
    // The report callback is deliberately not a dependency: a new closure for
    // the same bytes is not a new document.
  }, [source, fonts, defaults]);

  if (shown.state === "opening") return null;
  if (shown.state === "blocked") {
    return blocked?.(shown.reason) ?? <pre>{describeReason(shown.reason)}</pre>;
  }
  return (
    <Document
      layout={shown.layout}
      imageUrl={shown.imageUrl}
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
