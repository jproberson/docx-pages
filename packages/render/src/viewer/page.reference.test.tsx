import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import { layOutDocument, lookupFontMetrics, type LaidOutDocument } from "@onepager/core";
import { imageResolver, OnePagerPage } from "@onepager/viewer";

import { referenceCases, suppliedFaces, type ReferenceCase } from "../testing/cases.js";
import { readReferenceDocument } from "../testing/documents.js";

const rendered = (
  each: ReferenceCase,
): { readonly html: string; readonly layout: LaidOutDocument } => {
  const pkg = readReferenceDocument(each);
  const supplied = suppliedFaces();
  const layout = layOutDocument(pkg, (request) => lookupFontMetrics(request, supplied));
  if (layout.kind !== "laid-out") throw new Error(`blocked: ${layout.blocker.kind}`);
  return {
    layout,
    html: renderToStaticMarkup(
      <OnePagerPage layout={layout} imageUrl={imageResolver(pkg)} frames="outlined" />,
    ),
  };
};

const CASES = referenceCases();

describe.skipIf(CASES.length === 0)("rendering a real document", () => {
  for (const each of CASES) {
    describe(each.id, () => {
      it("draws every picture the layout resolved", () => {
        const { html, layout } = rendered(each);
        const pictures = [
          ...layout.headerFloats,
          ...layout.bodyFloats,
          ...layout.headerInlines,
          ...layout.bodyInlines,
        ].filter((placed) => placed.content.kind === "picture");

        expect(pictures.length).toBeGreaterThan(0);
        expect(html.split('data-kind="picture"')).toHaveLength(pictures.length + 1);
      });

      it("leaves nothing unresolved or unrecognised", () => {
        expect(rendered(each).html).not.toContain('data-kind="missing-picture"');
        expect(rendered(each).html).not.toContain('data-kind="unknown"');
      });

      it("embeds each picture rather than pointing at a file that will not be there", () => {
        expect(rendered(each).html).not.toMatch(/src="(?!data:image\/)/);
      });

      it("keeps every object inside the page it belongs to", () => {
        const { html } = rendered(each);
        expect(html).toContain("width:612pt");
        expect(html).toContain("height:792pt");
      });
    });
  }
});
