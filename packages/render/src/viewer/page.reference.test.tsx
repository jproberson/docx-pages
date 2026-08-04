import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import { layOutDocument, lookupFontMetrics, type LaidOutDocument } from "@onepager/core";
import { drawablesOf, imageResolver, OnePagerDocument } from "@onepager/viewer";

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

// A text layer states its size twice: once as a box on the page, in points, and
// once as the extent of the coordinates its contents are written in. The two have
// to agree, wherever on the page the layer's own origin is.
const TEXT_LAYER =
  /data-kind="text"[^>]*width="([\d.]+)pt" height="([\d.]+)pt" viewBox="[\d.-]+ [\d.-]+ ([\d.]+ [\d.]+)"/g;

const DRAWN_AT = /<text [^>]*?x="([-\d.]+)" y="([-\d.]+)"/g;

// Where the drawing is expected to start each line and each list number, in the
// page's own points.
const placedAt = (layout: LaidOutDocument): readonly string[] =>
  layout.pages.flatMap((page) =>
    drawablesOf(layout, page).flatMap((drawable) =>
      drawable.kind !== "text"
        ? []
        : drawable.boxes.flatMap((box) => [
            ...(box.marker === null || box.marker.text === ""
              ? []
              : [`${String(box.marker.leftPt)},${String(box.marker.baselinePt)}`]),
            ...box.lines.flatMap((line) =>
              line.line.segments.some((segment) => segment.kind === "text")
                ? [`${String(line.leftPt)},${String(line.baselinePt)}`]
                : [],
            ),
          ]),
    ),
  );

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

      // Every colour in these files is a theme slot under a luminance transform,
      // so a shape that resolved one and does not carry it into the markup is a
      // panel or a rule drawn as nothing at all.
      it("paints every shape in the colour the theme resolved for it", () => {
        const { html, layout } = rendered(each);
        const painted = layout.pages.flatMap((page) =>
          [...layout.headerFloats, ...layout.footerFloats, ...page.floats].flatMap((placed) =>
            "paint" in placed.content ? [placed.content.paint] : [],
          ),
        );
        const colors = new Set(
          painted.flatMap((paint) => [paint.fillColor, paint.outline?.color ?? null]),
        );
        colors.delete(null);

        expect(colors.size).toBeGreaterThan(0);
        for (const color of colors) expect(html).toContain(color);
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

      // Everything else the suite pins stops at the laid-out model. A layer sized
      // in the browser's own pixels while the page around it is sized in points
      // draws every glyph at three quarters of its size, three quarters of the way
      // to where it belongs, and nothing above this would notice.
      it("draws text in the points layout measured it in", () => {
        const { html } = rendered(each);
        for (const [, width, height, box] of html.matchAll(TEXT_LAYER)) {
          expect(`${String(width)} ${String(height)}`).toBe(box);
        }
        expect(html).toMatch(TEXT_LAYER);
      });

      it("puts every line where layout put it", () => {
        const { html, layout } = rendered(each);
        const drawn = [...html.matchAll(DRAWN_AT)].map(([, x, y]) => `${String(x)},${String(y)}`);

        expect(drawn.length).toBeGreaterThan(0);
        expect(new Set(drawn)).toStrictEqual(new Set(placedAt(layout)));
      });
    });
  }
});
