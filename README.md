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

| Package                                 | Runs on   | What it is                                                                       |
| --------------------------------------- | --------- | -------------------------------------------------------------------------------- |
| [`@docx-pages/core`](packages/core)     | anywhere  | Reads the package, lays it out, plays metafiles. Touches no disk and no network. |
| [`@docx-pages/viewer`](packages/viewer) | React 18+ | Draws a laid-out page.                                                           |

A third package, `render`, is not published. It is the Word oracle: it drives Word
over AppleScript, reads Word's own pdf back, and holds the test suites. It lives in
this repository and consumes the library's source rather than installing it.

## Using it

```
npm install @docx-pages/core @docx-pages/viewer
```

```tsx
import { layOutDocument, openDocx, substitutingMetrics } from "@docx-pages/core";
import { Document } from "@docx-pages/viewer";

// Every face the document names has to be supplied. Nothing is guessed: which
// face answers for a name this machine has not got is a question about the
// machine, not about the file, so the library never invents one.
const faces = substitutingMetrics(suppliedFaces, ["Cambria"]);

const layout = layOutDocument(openDocx(bytes), faces);
if (layout.kind !== "laid-out") {
  // A document is refused rather than drawn wrongly. `layout.blocker` says why.
  throw new Error(JSON.stringify(layout.blocker));
}

<Document layout={layout} />;
```

### Fonts are yours to supply

`core` reads font files but never goes looking for one. A face is handed in as a
`SuppliedFace` carrying its metrics and its glyph advances, both read out of the
file with `readFontFile`. The reader takes ttf, otf, ttc and woff, and refuses
woff2, which needs brotli.

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

Seventeen such gaps are ranked in [`docs/gaps.md`](docs/gaps.md), with the whole
backlog in [`docs/remaining.md`](docs/remaining.md). The ones worth knowing before
you start: character spacing, kerning and `keep-with-next` are read from the
document and reported, and not yet acted on.

## The evidence

Two suites, and only one of them can be published.

**Authored documents** are written by this repository to ask Word a single question
each, so Word's own answers about them are committed, in
`packages/render/src/authored/measured.json`. Nothing in them is anyone's
collateral: they invent their own words. This is the evidence anyone can reproduce,
on a machine with Word and Calibri:

```
pnpm verify     # format, lint, typecheck, test
pnpm measure    # rewrite measured.json by asking Word again
pnpm browse     # each rendering beside Word's own page
```

**Reference documents** are real one-pagers. They are labelled internal collateral
and are never committed, so the suites that use them find them through a manifest
and report nothing to run without one. A clone will skip them.

Two oracles answer, and they are not the same quantity. Word over AppleScript
reports empty paragraphs, which draw nothing and cannot be measured from a
rendering at all, and its answers are rounded to the whole point. Word's own pdf
says where something was drawn, to more decimal places than anything here needs.

## Licence

MIT.
