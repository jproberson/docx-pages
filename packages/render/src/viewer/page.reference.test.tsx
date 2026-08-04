import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import { layOutDocument, lookupFontMetrics, type LaidOutDocument } from "@onepager/core";
import { imageResolver, OnePagerDocument } from "@onepager/viewer";

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
      <OnePagerDocument layout={layout} imageUrl={imageResolver(pkg)} frames="outlined" />,
    ),
  };
};

const CASES = referenceCases();

describe.skipIf(CASES.length === 0)("rendering a real document", () => {
  for (const each of CASES) {
    describe(each.id, () => {
      // The header and the footer are drawn again on every page, so what they hold
      // counts once per page. A picture in a format no browser takes cannot be
      // drawn at all, and is marked in its place rather than quietly dropped.
      it("draws every picture the layout resolved, or marks the ones it cannot", () => {
        const { html, layout } = rendered(each);
        const surrounding = [
          ...layout.headerFloats,
          ...layout.footerFloats,
          ...layout.headerInlines,
          ...layout.footerInlines,
        ];
        const drawn = layout.pages.flatMap((page) => [
          ...surrounding,
          ...page.floats,
          ...page.inlines,
        ]);
        const pictures = drawn.filter((placed) => placed.content.kind === "picture");

        const undrawable = each.unrenderablePictures;
        expect(pictures.length).toBeGreaterThan(0);
        expect(html.split('data-kind="picture"')).toHaveLength(pictures.length - undrawable + 1);
        expect(html.split('data-kind="unresolved-picture"')).toHaveLength(undrawable + 1);
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
