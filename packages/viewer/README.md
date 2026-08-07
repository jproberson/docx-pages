# @docx-pages/viewer

Draws a `.docx` in React, at the positions `@docx-pages/core` worked out against
Word. React 18 or later.

The layout decides where everything goes; this package only paints it. Every
position it draws to is one core computed against Word, so the component's job is
to put a glyph where it was told and get out of the way.

```
npm install @docx-pages/viewer
```

That brings `@docx-pages/core` and `@docx-pages/fonts` with it. The font pack is
most of what gets installed, around 7.7 MiB of it, because it carries the face
files themselves.

## Drawing a document

```tsx
import { DocxDocument } from "@docx-pages/viewer";

<DocxDocument source={bytes} onReport={(report) => console.log(report.substitutions)} />;
```

That is the whole consumer. `DocxDocument` opens the bytes, resolves every face
the document names, lays it out, and paints it with the same bytes it measured
with, so what is drawn is the face the widths came from.

## It is never quiet about a stand-in

A face the document names that nothing supplies falls to a metric twin out of
`@docx-pages/fonts`, and a page drawn over one is no longer the page Word would
draw. `onReport` says so, after every layout:

- `substitutions` names every face another one answered for. Lines drawn in it
  may break where Word did not break them.
- `fallbackCharacters` names every character drawn out of a face the document
  never mentioned. The room it takes is Word's, so nothing moves, but the glyph
  in that room is the borrowed one.
- `missingGlyphs` names every character nothing could draw, which is painted as a
  box.
- `unhonoured` names what the document asked for that the layout passed over,
  each saying whether it moves text or only changes paint.

An application that shows a page can show what was doubtful about it.

## Supplying the real faces

Exactness is the `fonts` prop. A face handed in there is used as given, and a
document whose faces are all supplied is laid out as if there were no pack at
all:

```tsx
<DocxDocument source={bytes} fonts={[{ name: "Calibri", bytes: calibri }]} />
```

Bold and italic are separate faces, each with its own bytes and its own
`bold`/`italic` flags.

## Off a disk rather than a fetch

The pack finds its files beside its own module and reaches them with `fetch`,
which node will not do for a `file:` url. Under node, hand the defaults in:

```tsx
import { defaultFacesFromDisk } from "@docx-pages/fonts/node";

<DocxDocument source={bytes} defaults={await defaultFacesFromDisk()} />;
```

For the same reason a dev server that prebundles its dependencies moves the pack
away from its own files. Vite is the one met so far, and the way out is
configuration:

```ts
optimizeDeps: {
  exclude: ["@docx-pages/viewer", "@docx-pages/fonts"];
}
```

`vite build` needs nothing. Without it the component reports a blocker saying
exactly this.

## Drawing a layout laid out by hand

`Document` paints a layout somebody else made, which is what a caller wanting to
choose the faces itself, or to lay out once and paint many times, reaches for. It
needs the resolver the pictures come out of as well as the layout:

```tsx
import { bestEffortMetrics, layOutDocument, openDocx, readFaceShapes } from "@docx-pages/core";
import { defaultFaces } from "@docx-pages/fonts";
import { Document, imageResolver } from "@docx-pages/viewer";

const pkg = openDocx(bytes);
const faces = bestEffortMetrics(suppliedFaces, await defaultFaces(), readFaceShapes(pkg));
const layout = layOutDocument(pkg, faces);
if (layout.kind !== "laid-out") throw new Error(JSON.stringify(layout.blocker));

<Document layout={layout} imageUrl={imageResolver(pkg, faces.metricsFor)} />;
```

Nothing registers the faces with the browser on this path, which `DocxDocument`
does for itself. A face the CSS cannot load is drawn in whatever the browser
substitutes, at the widths core measured, which is the one way a page can be
right in its geometry and wrong on the screen.

`Page` takes the same props and one page besides, for a viewer paginating by
hand:

```tsx
{
  layout.pages.map((page, at) => (
    <Page key={at} page={page} layout={layout} imageUrl={imageResolver(pkg, faces.metricsFor)} />
  ));
}
```

## Pictures

A drawing's bytes live in the `.docx`, so the component is handed a resolver
rather than reaching for them itself. `imageResolver` builds one over a package.
Under it, `imageDataUrl` builds the url for a single part, and `drawablesOf` says
what a page holds to draw, for a caller writing its own painter.

## Frames

`frames="outlined"` draws the frame of every text box and anchored object. It is
for looking at a page beside Word's own and seeing which box a difference is in;
`"hidden"` is the default and is what a reader sees.

## Licence

MIT. The font pack it installs is separately licensed, part of it under the
SIL Open Font License; see `@docx-pages/fonts`.
