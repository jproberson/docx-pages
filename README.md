# docx-pages

Reads a `.docx`, works out where Word would put everything, page by page, and draws
it in the browser.

The point is **agreement with Word** rather than a plausible rendering. A floating
image's position is not stored in a `.docx` at all: the file holds a recipe
("anchor to this paragraph, 2.07 inches down") and the coordinate is computed at
render time from where the text above it lands. A different line breaker means a
different coordinate, which is why documents that rely on anchored objects come out
wrong in engines that delegate layout to a browser.

So nearly every rule here was measured against Word before it was built, and the
measurement is what the tests hold on to.

## The packages

| Package                                 | Runs on   | What it is                                                                                                                                                     |
| --------------------------------------- | --------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [`@docx-pages/core`](packages/core)     | anywhere  | Reads the package, lays it out, plays metafiles. Touches no disk and no network.                                                                               |
| [`@docx-pages/viewer`](packages/viewer) | React 18+ | Draws a laid-out page, or a whole `.docx` from its bytes.                                                                                                      |
| [`@docx-pages/pdf`](packages/pdf)       | anywhere  | Writes a laid-out page out as a pdf. Touches no disk and no network either.                                                                                    |
| [`@docx-pages/fonts`](packages/fonts)   | anywhere  | Freely redistributable twins of the faces documents usually name, for laying out without the real fonts. Optional: only `@docx-pages/viewer/pack` asks for it. |

The viewer and the pdf writer are two backends over the one layout, not two
renderings. Both walk the same `drawablesOf`, so what is drawn over what is
answered once; neither of them decides where anything sits.

A fifth package, `render`, is not published. It is the Word oracle: it drives Word
over AppleScript, reads Word's own pdf back, and holds the test suites. It lives in
this repository and consumes the library's source rather than installing it.

## Using it

```
npm install @docx-pages/viewer @docx-pages/fonts
```

```tsx
import { DocxDocument } from "@docx-pages/viewer/pack";

<DocxDocument source={bytes} onReport={(report) => console.log(report.substitutions)} />;
```

That is the whole consumer. Every face the document names that nothing supplies
falls to a metric twin from `@docx-pages/fonts`: Carlito for Calibri, Caladea
for Cambria, the Liberation family for Arial, Times New Roman and Courier New,
each matching the named face's advance widths glyph for glyph, so lines still
break where they would have broken. A name with no twin falls to a default of
the shape the document's own font table gives it. Nothing is quiet about any of
this: `onReport` names every face stood in for, every character borrowed from
another face, and every character nothing could draw. Supply the real fonts
through the `fonts` prop and none of it happens.

`@docx-pages/fonts` is an optional peer of the viewer, and the `/pack` entry above
is the only one that asks for it. A project supplying its own faces installs the
viewer alone, imports `DocxDocument` from `@docx-pages/viewer`, and states what to
fall back to itself; the pack's megabytes never enter its bundle.

Underneath, for a caller drawing by hand or running without React:

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

For exactness rather than resilience, hand `layOutDocument` a
`substitutingMetrics` over only the faces you supply, or a bare resolver: a
document asking for anything more is then **refused rather than drawn wrongly**,
with `layout.blocker` saying why. That is the mode every measurement in this
repository is made in.

### Fonts are yours to supply, and the pack answers when you do not

`core` reads font files but never goes looking for one. A face is handed in as a
`SuppliedFace` carrying its metrics and its glyph advances, both read out of the
file with `readFontFile`. The reader takes ttf, otf, ttc and woff, and refuses
woff2, which needs brotli. `@docx-pages/fonts` is a set of such faces that may
be shipped where the named ones may not, and everything below applies to a face
out of the pack exactly as it applies to one of yours.

Where a document names a face nothing supplies, `substitutingMetrics` stands the
nearest usable one in its place and **says so**: whatever it stood in for comes
back out of `substitutions()`, and every character it had to draw out of another
face comes back out of `fallbackCharacters()`. A page drawn on the back of one is
no longer the page Word would draw, and the library will not pretend otherwise.

## What a document asks for and does not get

`LaidOutDocument.unhonoured` names everything the layout passed over, each with
whether it moves text or only changes paint. It is not a list of things known to be
drawn wrongly: it is what a document asks for that nothing here answers, whether or
not it showed.

The ones worth knowing before you start, every one of them named in the report
rather than passed over quietly: kerning and `keep-with-next` are read from the
document and not yet acted on; a footnote takes no room at the
foot of its page; only the last section's geometry is read, so a document that
changes page size or margins part way through lays the rest out on the wrong
page; and a cell spanning its neighbours is laid out at its own width.

Two gaps are **not** named in the report, which makes them the ones to watch: the
old `w:pict` and `w:object` drawing form, which nothing here reads, and a
paragraph whose indents together exceed the frame they stand in, which is drawn to
a rule that was never measured against Word.

## The evidence

Word's own answers are committed, so the claim above is one anybody can check.

**Authored documents** are written by this repository to ask Word a single
question each: where a tab lands at every alignment, how tall a row is, what an
outline does to a box that fits itself to its text. They invent their own words
and hold nobody's content, so both the documents and Word's answers about them
ship here, in `packages/render/src/authored/measured.json`. The suite lays each
one out and holds the result to what Word said. On a machine with Word and
Calibri:

```
pnpm verify     # format, lint, typecheck, test
pnpm measure    # rewrite measured.json by asking Word again
pnpm browse     # each rendering beside Word's own page
```

Word is asked two ways, and they answer different quantities. Over AppleScript it
gives the position and page of every paragraph, including the empty ones, which
draw nothing and so cannot be measured from a rendering at all; those answers are
rounded to the whole point. Word's own pdf says where something was drawn, to more
decimal places than anything here needs, and is the oracle for anything about a
position.

A second suite runs against real documents, which are private and are no part of
this repository. It finds them through a manifest and reports nothing to run
without one, so a clone skips it.

## Licence

MIT.
