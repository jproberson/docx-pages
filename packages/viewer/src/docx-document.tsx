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
import {
  fontUrl,
  LAST_RESORT_DEFAULT,
  METRIC_TWINS,
  MONOSPACE_DEFAULT,
  PACK_FACES,
  SANS_SERIF_DEFAULT,
  SERIF_DEFAULT,
  type PackFace,
} from "@docx-pages/fonts";

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
      const url = fontUrl(face);
      // A dev server that prebundles its dependencies moves this module into a
      // deps cache, away from the files it resolves beside itself. Vite's is the
      // one met so far, and the way out is configuration, so say so.
      if (url.href.includes("/.vite/")) {
        throw new Error(
          `the font pack was prebundled away from its own files (${url.href}); add optimizeDeps: { exclude: ["@docx-pages/viewer", "@docx-pages/fonts"] } to vite.config`,
        );
      }
      const response = await fetch(url);
      if (!response.ok)
        throw new Error(`the font at ${url.href} came back ${String(response.status)}`);
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
          sansSerif: face.sansSerif ?? read.sansSerif,
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
  | {
      readonly state: "shown";
      readonly layout: LaidOutDocument;
      readonly imageUrl: ImageResolver;
      readonly aliasSymbolFaces: ReadonlySet<string> | null;
    };

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
            sansSerif: SANS_SERIF_DEFAULT,
            serif: SERIF_DEFAULT,
            monospace: MONOSPACE_DEFAULT,
            lastResort: LAST_RESORT_DEFAULT,
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
