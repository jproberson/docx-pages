# @docx-pages/fonts

Freely redistributable faces whose metrics twin the ones documents usually
name, so `@docx-pages/core` can lay a document out when the named fonts cannot
be supplied: Carlito for Calibri, Caladea for Cambria, and the Liberation
family for Arial, Times New Roman and Courier New. Advance widths match glyph
for glyph, so lines break where the named face would have broken them.

```ts
import { bestEffortMetrics, layOutDocument, openDocx, readFaceShapes } from "@docx-pages/core";
import { defaultFaces } from "@docx-pages/fonts"; // or "@docx-pages/fonts/node" off a disk

const pkg = openDocx(bytes);
const faces = bestEffortMetrics(suppliedFaces, await defaultFaces(), readFaceShapes(pkg));
const layout = layOutDocument(pkg, faces); // every stand-in comes back out of faces.substitutions()
```

A document naming a face with no twin falls to a default of the shape the
document's own font table gives the name, and a name nothing classifies falls
to Caladea, because Word's own last resort for a name it cannot place is
Cambria.

The font files keep their own licenses, committed beside them under `fonts/`.

One consumer note: the pack finds its files beside its own module, so a dev
server that prebundles dependencies has to leave these two alone. Under Vite
that is

```ts
optimizeDeps: {
  exclude: ["@docx-pages/viewer", "@docx-pages/fonts"];
}
```

and `vite build` needs nothing.
