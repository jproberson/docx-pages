# @docx-pages/core

Reads a `.docx` and works out where Word would put everything, page by page.
Isomorphic: it touches no disk and no network, so the caller hands in the bytes and
the fonts.

Draw the result with `@docx-pages/viewer`, or with your own renderer against the
same types.

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
  entry.kind; // "character-spacing", "keep-with-next", "footnote", ...
  entry.effect; // "moves-text" or "changes-paint"
}
```

`moves-text` puts every page below it in doubt; `changes-paint` is wrong only where
it stands. The report is not a list of known-wrong drawings: it is what the
document asks for that nothing here answers, whether or not it showed.

## Limits worth knowing

- **woff2 is refused.** It needs brotli, which is not in every runtime.
- **cmap formats 4 and 12 only.** Anything else refuses the document rather than
  guessing a width.
- **No kerning and no ligatures.** A line is the sum of its characters' advances.
- **Bold and italic are separate faces**, each needing its own file.

## Licence

MIT.
