# @docx-pages/viewer

Draws a page that `@docx-pages/core` has laid out. React 18 or later.

The layout decides where everything goes; this package only paints it. Every
position it draws to is one core computed against Word, so the component's job is
to put a glyph where it was told and get out of the way.

```
npm install @docx-pages/core @docx-pages/viewer
```

## Drawing

```tsx
import { layOutDocument, openDocx, substitutingMetrics } from "@docx-pages/core";
import { Document } from "@docx-pages/viewer";

const faces = substitutingMetrics(suppliedFaces, ["Cambria"]);
const layout = layOutDocument(openDocx(bytes), faces);
if (layout.kind !== "laid-out") throw new Error(JSON.stringify(layout.blocker));

<Document layout={layout} />;
```

`Page` draws one page on its own, which is what a viewer paginating by hand wants.

## Pictures

A drawing's bytes live in the `.docx`, so the component is handed a resolver rather
than reaching for them itself:

```tsx
import { imageResolver, Document } from "@docx-pages/viewer";

<Document layout={layout} imageUrl={imageResolver(pkg, metricsFor)} />;
```

`imageDataUrl` builds one url, and `drawablesOf` says what a page holds to draw, for
a caller writing its own painter.

## Frames

`frames="outlined"` draws the frame of every text box and anchored object. It is
for looking at a page beside Word's own and seeing which box a difference is in;
`"hidden"` is the default and is what a reader sees.

## The faces have to be loaded too

The component draws text in the face the document names, so that face has to be
available to the browser as well as supplied to the layout. A face the CSS cannot
load is drawn in whatever the browser substitutes, at the widths core measured,
which is the one way a page can be right in its geometry and wrong on the screen.

## Licence

MIT.
