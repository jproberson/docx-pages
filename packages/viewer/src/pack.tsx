import { useEffect, useState, type ReactElement } from "react";

import { readFontFile, type FaceDefaults } from "@docx-pages/core";
import {
  fontUrl,
  LAST_RESORT_DEFAULT,
  METRIC_TWINS,
  MONOSPACE_DEFAULT,
  PACK_FACES,
  SANS_SERIF_DEFAULT,
  SERIF_DEFAULT,
} from "@docx-pages/fonts";

import {
  DocxDocument as DrawDocument,
  type DocxDocumentProps,
  type DocxFont,
} from "./docx-document.js";

// The one module here that names `@docx-pages/fonts`. Everything a consumer
// reaches through `@docx-pages/viewer` itself is free of it, so a project that
// supplies its own faces never resolves the pack and never installs its
// megabytes. Importing this file is what asks for them.

type Pack = {
  readonly defaults: FaceDefaults;
  readonly bytes: readonly DocxFont[];
};

// The pack is fetched once per page, not once per document.
let packOnce: Promise<Pack> | null = null;

const loadPack = (): Promise<Pack> => {
  packOnce ??= (async (): Promise<Pack> => {
    const read = await Promise.all(
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
        const found = readFontFile(bytes);
        return {
          bytes: { name: face.name, bold: face.bold, italic: face.italic, bytes },
          supplied: {
            name: face.name,
            bold: face.bold,
            italic: face.italic,
            metrics: found.metrics,
            advances: found.advances,
            sansSerif: face.sansSerif ?? found.sansSerif,
          },
        };
      }),
    );

    return {
      defaults: {
        faces: read.map((each) => each.supplied),
        twins: METRIC_TWINS,
        sansSerif: SANS_SERIF_DEFAULT,
        serif: SERIF_DEFAULT,
        monospace: MONOSPACE_DEFAULT,
        lastResort: LAST_RESORT_DEFAULT,
      },
      bytes: read.map((each) => each.bytes),
    };
  })();
  return packOnce;
};

// The pack answers for `defaults`, so a caller of this entry never states them.
export type PackedDocxDocumentProps = Omit<DocxDocumentProps, "defaults" | "defaultBytes">;

/**
 * `DocxDocument` with `@docx-pages/fonts` behind it: every face the document
 * names that `fonts` does not supply falls to a metric twin out of the pack, and
 * `onReport` says which. This is the batteries-included entry, and the reason
 * `@docx-pages/fonts` has to be installed alongside the viewer.
 *
 * A project that supplies all its own faces imports `DocxDocument` from
 * `@docx-pages/viewer` instead, states its own `defaults`, and never pulls the
 * pack into its bundle.
 */
export function DocxDocument(props: PackedDocxDocumentProps): ReactElement | null {
  const [pack, setPack] = useState<Pack | null>(null);
  const [failed, setFailed] = useState<unknown>(null);

  useEffect(() => {
    let stale = false;
    loadPack().then(
      (loaded) => {
        if (!stale) setPack(loaded);
      },
      (error: unknown) => {
        if (!stale) setFailed(error);
      },
    );
    return () => {
      stale = true;
    };
  }, []);

  if (failed !== null) {
    return props.blocked?.(failed) ?? <pre>{describeReason(failed)}</pre>;
  }
  if (pack === null) return null;
  return <DrawDocument {...props} defaults={pack.defaults} defaultBytes={pack.bytes} />;
}

function describeReason(reason: unknown): string {
  if (reason instanceof Error) return reason.message;
  try {
    return JSON.stringify(reason);
  } catch {
    return String(reason);
  }
}
