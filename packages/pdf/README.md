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
import { layOutDocument, lookupFontMetrics, openDocx } from "@docx-pages/core";
import { writePdf } from "@docx-pages/pdf";

const pkg = openDocx(bytes);
const metricsFor = (request) => lookupFontMetrics(request, supplied);
const layout = layOutDocument(pkg, metricsFor);
if (layout.kind !== "laid-out") throw new Error(JSON.stringify(layout.blocker));

const pdf = writePdf(layout, { fonts, imageBytes: (part) => pkg.parts.get(part), metricsFor });
```

`metricsFor` is asked for as well as `fonts` because a metafile picture records
text as a face and a string rather than as a drawing of one: playing it back
measures it, and it is measured with whatever the layout measured with.

A caller who laid the document out over stand-ins passes `aliasSymbolFaces` too,
the same set the viewer takes, so a run written in a symbol face that was stood in
for is drawn as what its positions mean rather than as the stand-in's own letters.

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
- Underlines, as the filled rectangle Word draws one as, where the face's own
  `post` table says to put it. Measured on 2026-08-07 against Word's own pdf: it
  puts the line where the drawn face states and not at a place of its own.
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
- **No png**, and no gif, bmp, tiff or webp. Only jpeg is passed through, because
  only jpeg is already a compression a pdf understands. A png would have to be
  inflated and its predictors undone before it could be written, and the others
  would have to be encoded into something a pdf carries. None of them is drawn.
- **No CMYK jpeg.** A four-channel jpeg is left undrawn beside the png. Word writes
  them inverted often enough that drawing one the wrong way round is worse than
  not drawing it, and which of the two it is cannot be told from the frame header
  alone. Greyscale and colour jpegs, which is nearly all of them, go through.
- **A picture nothing can draw leaves its frame empty.** It is not an error and it
  is not a placeholder: the paint round it is still drawn, and the picture is not.
  The viewer outlines such a frame when it is asked to; a file being written has
  nobody to show an outline to.
- **No underline under a face stating no `post` table.** The run is drawn without
  one rather than under a line in a place nothing measured. Every real face states
  the table, so this is a gap in principle more than in practice.
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
