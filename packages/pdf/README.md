# @docx-pages/pdf

Writes a `.docx` laid out by [`@docx-pages/core`](../core) out as a pdf.

A second backend over the same layout the [viewer](../viewer) draws, not a second
renderer. Where a line sits, how tall a row is and which page a paragraph lands on
were all settled before this package is called, and were measured against Word;
this decides nothing about any of it. Both backends walk the one traversal,
`drawablesOf`, so a page stacks the same way on the screen and in the file.

Isomorphic, like core: it touches no disk and no network, reads no clock, and
answers with a `Uint8Array`. The same bytes in a browser tab and in Node.

## Using it

```ts
import { pdfOfDocx } from "@docx-pages/pdf";

const pdf = pdfOfDocx(bytes, { fonts: [{ name: "Calibri", bytes: calibri }] });
```

That is the whole of it. Underneath, for a caller who has already laid a document
out, or who lays one out once and writes it more than once:

```ts
import { layOutDocument, openDocx, substitutingMetrics } from "@docx-pages/core";
import { writePdf } from "@docx-pages/pdf";

const pkg = openDocx(bytes);
const faces = substitutingMetrics(supplied);
const layout = layOutDocument(pkg, faces);
if (layout.kind !== "laid-out") throw new Error(JSON.stringify(layout.blocker));

const pdf = writePdf(layout, {
  fonts,
  imageBytes: (part) => pkg.parts.get(part),
  metricsFor: faces.metricsFor,
});
```

## Fonts are yours to supply, and there is no falling back

A pdf carries the faces it draws in, so this package needs the bytes of every one
the document names, handed in as `fonts` under the name the document names it.

**A face the document draws in that `fonts` does not supply refuses the
document**, with a `DocxPagesError` whose code is `font-not-supplied`. There is no
stand-in here and there deliberately is none. On a screen a substituted face is a
page that is right in its geometry and wrong in its letters, and the viewer says
so through its report; in a file that nobody is watching being made, the same page
would go out looking finished. The rest of this project would rather refuse than
be wrong quietly, and so does this.

For a page that must be drawn whatever it costs, stand the faces in before laying
out, with `bestEffortMetrics` or `substitutingMetrics`, and hand the bytes of
whatever stood in. Then the substitution is one the caller made and can report,
rather than one this package made and did not.

The whole face is embedded, under Identity-H, with a `ToUnicode` map so the text
can still be selected and searched. See the gaps below on what that costs.

## What is drawn

- Text, at the baselines layout measured, in the faces supplied.
- Paragraph and cell fills, and their borders: single, double, dashed and dotted.
- Shapes and text boxes: their fill, their outline, and their own text.
- Pictures: jpeg passed through as it stands, and metafiles played back as the
  vector drawing they record rather than rasterised.

## What a document asks for and does not get

Named here rather than passed over quietly, which is the same bargain
`LaidOutDocument.unhonoured` makes about the layout.

- **No font subsetting.** The whole of every face goes into the file, every glyph
  of it, including the thousands the document never draws. The output is correct
  and it is larger than it needs to be: a document in one face of Calibri carries
  the whole of Calibri. This is the one gap here that is only about size.
- **No png.** Only jpeg is passed through. A png would have to be inflated and its
  predictors undone to be re-encoded, or handed to an encoder this package does
  not have. A png picture is not drawn at all.
- **No encryption**, and so no permissions and no password.
- **No transparency beyond a flat fill.** A soft mask, a blend mode or a partly
  transparent picture is drawn as though it were opaque.
- **No tagged structure.** The file is not a tagged pdf: text can be selected and
  copied, and a screen reader is given no reading order, no headings and no
  alternative text for a picture.

Everything the layout itself passed over is passed over here too, and is in
`layout.unhonoured` rather than in this list.

## Licence

MIT.
