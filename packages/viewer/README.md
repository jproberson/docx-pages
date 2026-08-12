# @docx-pages/viewer

Draws a `.docx` in React, at the positions `@docx-pages/core` worked out against
Word. React 18 or later.

The layout decides where everything goes; this package only paints it. Every
position it draws to is one core computed against Word, so the component's job is
to put a glyph where it was told and get out of the way.

## Which of the two you want

The faces are the whole question. A document names faces, and something has to
answer for the ones you cannot ship, so this package has two entries and they
differ only in what answers.

```
npm install @docx-pages/viewer @docx-pages/fonts   # let the pack answer
npm install @docx-pages/viewer                     # answer for them yourself
```

`@docx-pages/fonts` is an optional peer, so the second install really does leave
it out, and nothing the root entry reaches names it. The pack carries font files
and is around 7.7 MiB; the viewer without it bundles to a couple of hundred
kilobytes.

## Letting the pack answer

```tsx
import { DocxDocument } from "@docx-pages/viewer/pack";

<DocxDocument source={bytes} onReport={(report) => console.log(report.substitutions)} />;
```

That is the whole consumer. It opens the bytes, falls back to a metric twin for
every face the document names that you have not supplied, lays it out, and paints
it with the same bytes it measured with.

## Answering for the faces yourself

The root entry is the same component with nothing behind it, so it asks you what
to fall back to. `defaults` is a `FaceDefaults` from core: the faces to fall back
through, the twin each name maps to, and the shape defaults behind that.

```tsx
import { DocxDocument } from "@docx-pages/viewer";

<DocxDocument
  source={bytes}
  fonts={[{ name: "Calibri", bytes: calibri }]}
  defaults={{
    faces: myFaces,
    twins: { calibri: "My Sans" },
    sansSerif: "My Sans",
    serif: "My Serif",
    monospace: "My Mono",
    lastResort: "My Serif",
  }}
  defaultBytes={[{ name: "My Sans", bytes: mySans }]}
/>;
```

`defaultBytes` is what the browser paints the fallbacks with. Leave it out and a
stood-in face is measured here and painted by whatever the browser finds, which
is the one way a page is right in its geometry and wrong on the screen.

## Supplying the real faces

Either entry takes `fonts`, and that is what exactness means: a face handed in
there is used as given, and a document whose faces are all supplied is laid out
as if nothing were behind it. Bold and italic are separate faces, each with its
own bytes and its own `bold`/`italic` flags.

## It is never quiet about a stand-in

A page drawn over a fallback is no longer the page Word would draw. `onReport`
says so, after every layout:

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

## Under node, and under a dev server

The pack finds its files beside its own module and reaches them with `fetch`,
which node will not do for a `file:` url. Under node, read them off the disk and
hand them to the root entry:

```tsx
import { DocxDocument } from "@docx-pages/viewer";
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

`vite build` needs nothing, and neither does the root entry, which has no files
to be moved away from. Without it `@docx-pages/viewer/pack` reports a blocker
saying exactly this.

## Drawing a layout laid out by hand

`Document` paints a layout somebody else made, which is what a caller wanting to
lay out once and paint many times reaches for. It needs the resolver the pictures
come out of as well as the layout:

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
does for itself.

`Page` takes the same props and one page besides, for a viewer paginating by
hand:

```tsx
{
  layout.pages.map((page, at) => (
    <Page key={at} page={page} layout={layout} imageUrl={imageResolver(pkg, faces.metricsFor)} />
  ));
}
```

## Downloading the page as a pdf

`onReady` hands out a function that writes the page being shown out as a pdf, and
`downloadPdf` offers those bytes to the browser as a file.

```tsx
const [pdf, setPdf] = useState<(() => Uint8Array) | null>(null);

return (
  <>
    <button
      disabled={pdf === null}
      onClick={() => pdf && downloadPdf(pdf(), { fileName: "page.pdf" })}
    >
      Download pdf
    </button>
    <DocxDocument source={bytes} defaults={defaults} onReady={(write) => setPdf(() => write)} />
  </>
);
```

**What is written is what is drawn, and not a second reading of it.** The same
measured layout, the same faces it was painted with, and the same `drawablesOf`
this component walks, so the file cannot say something the screen did not: there
is nothing in between for the two to disagree about. Nothing is laid out twice
either, so the button costs only the writing.

The faces go into the file the way they went onto the screen, which includes the
stand-ins: a face this component stood in for is carried under the name the
document asked for, since that is the name the layout measured it as. A face
nothing supplied the bytes of cannot be carried at all, and writing throws
`font-not-supplied` rather than putting out a page in a face nothing measured. On
screen that same face is drawn by whatever the browser finds, and the report says
so; a file has nobody to say it to.

## Pictures

A drawing's bytes live in the `.docx`, so the component is handed a resolver
rather than reaching for them itself. `imageResolver` builds one over a package.
Under it, `imageDataUrl` builds the url for a single part, and `drawablesOf` says
what a page holds to draw and in what order, for a caller writing its own painter.
That traversal is `@docx-pages/core`'s, offered again here so a caller who has the
viewer need not reach past it: this component and core's own `writePdf` walk the
one list, and so stack a page the same way.

## Frames

`frames="outlined"` draws the frame of every text box and anchored object. It is
for looking at a page beside Word's own and seeing which box a difference is in;
`"hidden"` is the default and is what a reader sees.

## Licence

MIT. `@docx-pages/fonts`, if you install it, carries font files under their own
licences, part of them the SIL Open Font License.
