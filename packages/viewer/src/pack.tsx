import { useEffect, useState, type ReactElement } from "react";

import { readPack, type FontPack } from "@docx-pages/fonts";

import { DocxDocument as DrawDocument, type DocxDocumentProps } from "./docx-document.js";
import { loadedOnce } from "./loaded-once.js";

// The one module here that names `@docx-pages/fonts`. Everything a consumer
// reaches through `@docx-pages/viewer` itself is free of it, so a project that
// supplies its own faces never resolves the pack and never installs its
// megabytes. Importing this file is what asks for them.

// The pack is read once for the page rather than once for each document drawn on
// it, and a read that failed is forgotten so that a later mount asks again.
const loadPack = loadedOnce(readPack);

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
  const [pack, setPack] = useState<FontPack | null>(null);
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
