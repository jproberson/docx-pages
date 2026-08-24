# @docx-pages/core

Reads a `.docx` and works out where Word would put everything, page by page.
Isomorphic: it touches no disk and no network, so the caller hands in the bytes and
the fonts.

Draw the result with `@docx-pages/viewer`, write it out as a pdf with `writePdf`
below, or use your own renderer against the same types. All three walk the one
traversal, `drawablesOf`, so none of them decides where anything sits.

```
npm install @docx-pages/core
```

## Laying a document out

```ts
import { layOutDocument, openDocx, substitutingMetrics } from "@docx-pages/core";

const faces = substitutingMetrics(suppliedFaces, ["Cambria"]);
const layout = layOutDocument(openDocx(bytes), faces);

if (layout.kind === "laid-out") {
  for (const page of layout.pages) {
    // page.body   paragraphs, each with its lines and where they sit
    // page.cells  table cells
    // page.floats anchored objects, in the order Word draws them
    // page.inlines drawings that stand in the flow of the text
  }
}
```

A document is **refused rather than drawn wrongly**. Where `layout.kind` is
`"blocked"`, `layout.blocker` says which of four things went wrong: the style
cascade could not name a run's face (`unresolved-font`), nothing supplies a name
and no fallback answers for it (`unknown-font-metrics`), a face was supplied in a
form the reader cannot measure with (`unmeasurable-text`), or no face on the
machine has a glyph for some character (`unmapped-character`).

## Supplying faces

Nothing is guessed. A face is handed in carrying its metrics and its glyph
advances, both read out of the file:

```ts
import { readFontFile, type SuppliedFace } from "@docx-pages/core";

const read = readFontFile(bytes); // ttf, otf, ttc or woff. Not woff2.
const face: SuppliedFace = {
  name: "Calibri",
  bold: false,
  italic: false,
  metrics: read.metrics,
  advances: read.advances,
  sansSerif: read.sansSerif,
};
```

`readFontFile(bytes, faceName)` picks one face out of a collection by name, since a
`.ttc` holds several.

`substitutingMetrics` resolves a face the way Word does and is never quiet about
it. Read both after laying out, not before, since nothing is known until the layout
has asked for the faces it needs:

- `substitutions()` names every face the document asked for that another one
  answered for. Every line drawn in it may break where Word did not break it.
- `fallbackCharacters()` names every character drawn out of a face the document
  never mentioned. The room it takes is Word's, so nothing moves, but the glyph
  drawn in that room is whatever the renderer finds.

Which face answers for a character its own has no glyph for turns on the kind of
face that asked and then on the character. `WORD_SANS_FALLBACK_FACE`,
`WORD_SERIF_FALLBACK_FACE`, `WORD_EMOJI_FACE` and `WORD_CHARACTER_FALLBACK_FACES`
say which faces are reached for, so a caller can supply them. A face the machine
has not got is passed over rather than refused.

## What the document asked for and did not get

```ts
for (const entry of layout.unhonoured) {
  entry.kind; // "keep-lines-together", "text-columns", "footnote", ...
  entry.effect; // "moves-text" or "changes-paint"
}
```

`moves-text` puts every page below it in doubt; `changes-paint` is wrong only where
it stands. The report is not a list of known-wrong drawings: it is what the
document asks for that nothing here answers, whether or not it showed.

## Writing a pdf

```ts
import { pdfOfDocx } from "@docx-pages/core";

const pdf = pdfOfDocx(bytes, { fonts: [{ name: "Calibri", bytes: calibri }] });
```

Underneath, for a caller who has already laid a document out, or who lays one out
once and writes it more than once:

```ts
const pkg = openDocx(bytes);
const layout = layOutDocument(pkg, metricsFor);
if (layout.kind !== "laid-out") throw new Error(JSON.stringify(layout.blocker));

const pdf = writePdf(layout, { fonts, imageBytes: (part) => pkg.parts.get(part), metricsFor });
```

`metricsFor` is asked for as well as `fonts` because a metafile picture records
text as a face and a string rather than as a drawing of one: playing it back
measures it, and it is measured with whatever the layout measured with. A caller
who laid the document out over stand-ins passes `aliasSymbolFaces` too, the same
set the viewer takes, so a run written in a symbol face that was stood in for is
drawn as what its positions mean rather than as the stand-in's own letters.

### Fonts are yours to supply, and there is no falling back

A pdf carries the faces it draws in, so writing needs the bytes of every one the
document names, handed in under the name the document names it.

**A face the document draws in that `fonts` does not supply refuses the document**,
with a `DocxPagesError` whose code is `font-not-supplied`. There is deliberately no
stand-in. On a screen a substituted face is a page right in its geometry and wrong
in its letters, and the viewer says so through its report; in a file nobody is
watching being made, the same page would go out looking finished.

For a page that must be written whatever it costs, stand the faces in before laying
out, with `bestEffortMetrics` or `substitutingMetrics`, and hand the bytes of
whatever stood in. Then the substitution is one the caller made and can report.

The whole face is embedded, under Identity-H, with a `ToUnicode` map so the text
can still be selected and searched.

### What is drawn

- Text, at the baselines layout measured, in the faces supplied.
- Underlines, as the filled rectangle Word draws one as, where the face's own
  `post` table says to put it. Measured on 2026-08-07 against Word's own pdf: it
  puts the line where the drawn face states and not at a place of its own.
- Paragraph and cell fills, and their borders: single, double, dashed and dotted.
- Shapes and text boxes: their fill, their outline, and their own text.
- Pictures: jpeg and png, and metafiles played back as the vector drawing they
  record rather than rasterised.

A picture costs as little as it can. **A jpeg goes across as it stands**, since it
is already the compression a pdf would have applied. **So does a png that carries
no alpha**: a pdf deflates and predicts its pixels exactly as a png does, so the
`IDAT` stream is already a pdf image stream and is written untouched, with
`/Predictor 15` naming the arrangement. Neither is decoded and neither is
compressed a second time. A png that carries alpha is the one picture whose pixels
have to be opened, because a png keeps what shows through in with the colour and a
pdf keeps it in a separate image; it is inflated, unfiltered, split, and written as
the picture and its soft mask. That is the path that matters, since nearly every
png a real document holds carries alpha.

### What a document asks for and the writer does not draw

Named here rather than passed over quietly, which is the same bargain
`layout.unhonoured` makes about the layout. Everything the layout itself passed
over is in that report rather than in this list.

- **No font subsetting.** The whole of every face goes in, every glyph of it,
  including the thousands the document never draws. The output is correct and
  larger than it needs to be: a reference one-pager comes out at 2.8MB, of which
  2.5MB is five embedded faces and 0.24MB is every picture on it. Word's own pdf of
  the same document is 1.0MB.
- **No gif, bmp, tiff or webp.** Only jpeg and png are written, because only those
  two are already compressions a pdf understands. None of the rest is drawn.
- **No interlaced png**, which holds its rows in seven passes that would have to be
  woven back together. Left undrawn rather than drawn as a smear. Vanishingly rare.
- **No png that is not eight bits to a sample.** No other depth has been met at all,
  so this is a gap in principle rather than in practice.
- **A partly transparent palette is drawn opaque.** Where an indexed png says an
  entry is wholly invisible it is masked out and the picture still crosses
  untouched; where it says an entry is _half_ transparent, honouring it would mean
  opening the pixels, and the picture is drawn solid instead.
- **No CMYK jpeg.** Word writes them inverted often enough that drawing one the
  wrong way round is worse than not drawing it, and which of the two it is cannot be
  told from the frame header alone. Greyscale and colour jpegs go through.
- **A picture nothing can draw leaves its frame empty.** Not an error and not a
  placeholder: the paint round it is still drawn, and the picture is not.
- **No underline under a face stating no `post` table.** The run is drawn without one
  rather than under a line in a place nothing measured.
- **No encryption**, and so no permissions and no password.
- **No transparency beyond a picture's own.** A shape or a run Word draws part way
  through what is behind it is drawn solid, and a blend mode is not read at all.
- **No tagged structure.** Text can be selected and copied; a screen reader is given
  no reading order, no headings and no alternative text for a picture.

## Limits worth knowing

- **woff2 is refused.** It needs brotli, which is not in every runtime.
- **cmap formats 0, 4, 6, 12 and 13.** A face whose best subtable is format 2
  refuses the document rather than guessing a width.
- **No ligatures.** A line is the sum of its characters' advances and whatever
  the face's pairs move them by. The pairs themselves are read, from `kern` and
  from GPOS.
- **Only the Latin slots of `w:rFonts`.** The East Asian and complex-script slots
  are unread, so a run of CJK, Arabic or Devanagari text is measured in whatever
  face the Latin slot named.
- **A face that sets equations needs its MATH table and its outlines**, which all
  but a handful of faces state nothing of. A document holding an equation and
  naming no such face is refused rather than set out of one that cannot say where
  a fraction's bar goes.
- **Bold and italic are separate faces**, each needing its own file.

## Licence

MIT.
