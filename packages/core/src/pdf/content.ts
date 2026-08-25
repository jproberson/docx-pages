import { strToU8 } from "fflate";

import { DocxPagesError } from "../errors.js";
import type { LaidOutDocument } from "../layout/document.js";
import { drawablesOf, faceAskedFor, type PageDrawing } from "../layout/drawables.js";

import { formatPdfNumber } from "./objects.js";
import { pdfPageOf, turnedAboutInPdf, upFromTop, type PdfPage } from "./coordinates.js";
import type { PdfFonts } from "./fonts.js";
import type { PdfImages } from "./images.js";
import { paintedObject, type ObjectDrawable } from "./objects-paint.js";
import { paintLayer } from "./paint.js";
import { drawnGlyphs, textOfBoxes, type TextOptions } from "./text.js";

// A page's drawing, as the operators a pdf writes it in.
//
// The order things are drawn in is not decided here. `drawablesOf` decides it, in
// core, and the viewer walks the very same list: what is drawn over what is one
// question with one answer, and answering it twice is how two backends come to
// disagree about a page neither of them is wrong about on its own.

// What has already been set, so that a run of text in one face and colour does not
// state both again for every line. A pdf reader carries this state itself, and
// restating it is only bytes; leaving it stale is a page drawn in the wrong
// colour, so the two are not the same kind of mistake and the cache is dropped
// wherever the state could have been restored under it.
type Graphics = {
  fillColor: string | null;
  strokeColor: string | null;
  lineWidth: number | null;
  dash: string | null;
  font: string | null;
  characterSpacing: number | null;
  characterScale: number | null;
};

const nothingSet = (): Graphics => ({
  fillColor: null,
  strokeColor: null,
  lineWidth: null,
  dash: null,
  font: null,
  characterSpacing: null,
  characterScale: null,
});

export type Content = {
  readonly save: () => void;
  readonly restore: () => void;
  readonly fillColor: (color: string) => void;
  readonly strokeColor: (color: string) => void;
  readonly lineWidth: (widthPt: number) => void;
  readonly dash: (pattern: readonly number[] | null) => void;
  readonly rectangle: (leftPt: number, bottomPt: number, widthPt: number, heightPt: number) => void;
  readonly line: (fromXPt: number, fromYPt: number, toXPt: number, toYPt: number) => void;
  // The path a shape that is not a rectangle is laid down point by point as. A
  // curve takes its two control points and then where it ends, which is what a
  // quarter of an ellipse and the corner of a rounded rectangle are both made of.
  readonly moveTo: (xPt: number, yPt: number) => void;
  readonly lineTo: (xPt: number, yPt: number) => void;
  readonly curveTo: (
    firstXPt: number,
    firstYPt: number,
    secondXPt: number,
    secondYPt: number,
    toXPt: number,
    toYPt: number,
  ) => void;
  readonly closePath: () => void;
  readonly fill: () => void;
  readonly stroke: () => void;
  // Both at once, which a rectangle asking for a colour behind it and a line round
  // it takes rather than laying the path down twice.
  readonly fillAndStroke: () => void;
  readonly clip: () => void;
  readonly transform: (matrix: readonly number[]) => void;
  readonly drawObject: (resource: string) => void;
  readonly beginText: () => void;
  readonly endText: () => void;
  readonly font: (resource: string, sizePt: number) => void;
  readonly characterSpacing: (spacingPt: number) => void;
  readonly characterScale: (scale: number) => void;
  readonly textPosition: (xPt: number, yPt: number) => void;
  // Where a run stands and which way up it is. A metafile plays under a flipped
  // transform, since its own units count down the page, and text drawn under one
  // would come out mirrored without a flip of its own here to undo it.
  readonly textMatrix: (matrix: readonly number[]) => void;
  readonly showGlyphs: (glyphs: Uint8Array) => void;
  readonly bytes: () => Uint8Array;
};

const numbers = (values: readonly number[]): string => values.map(formatPdfNumber).join(" ");

const HEX_COLOR = /^#?([0-9a-fA-F]{6})$/;

// A colour reaches here as the six hex digits layout resolved it to. A pdf states
// each channel as a fraction of one instead, to as many places as it takes.
function channelsOf(color: string, at: string): readonly number[] {
  const digits = HEX_COLOR.exec(color)?.[1];
  if (digits === undefined) {
    throw new DocxPagesError({
      code: "pdf-colour-unreadable",
      message: "a colour reaching the writer is six hex digits, and this one is not",
      at,
      context: { color },
    });
  }
  return [0, 2, 4].map((offset) => Number.parseInt(digits.slice(offset, offset + 2), 16) / 255);
}

const AT = "pdf/content.contentOf";

export function content(): Content {
  const parts: string[] = [];
  let set = nothingSet();

  const write = (text: string): void => {
    parts.push(text);
  };

  return {
    save: () => {
      write("q");
      set = nothingSet();
    },
    restore: () => {
      write("Q");
      set = nothingSet();
    },
    fillColor: (color) => {
      if (set.fillColor === color) return;
      set.fillColor = color;
      write(`${numbers(channelsOf(color, AT))} rg`);
    },
    strokeColor: (color) => {
      if (set.strokeColor === color) return;
      set.strokeColor = color;
      write(`${numbers(channelsOf(color, AT))} RG`);
    },
    lineWidth: (widthPt) => {
      if (set.lineWidth === widthPt) return;
      set.lineWidth = widthPt;
      write(`${formatPdfNumber(widthPt)} w`);
    },
    dash: (pattern) => {
      const stated = pattern === null || pattern.length === 0 ? "[] 0" : `[${numbers(pattern)}] 0`;
      if (set.dash === stated) return;
      set.dash = stated;
      write(`${stated} d`);
    },
    rectangle: (leftPt, bottomPt, widthPt, heightPt) => {
      write(`${numbers([leftPt, bottomPt, widthPt, heightPt])} re`);
    },
    line: (fromXPt, fromYPt, toXPt, toYPt) => {
      write(`${numbers([fromXPt, fromYPt])} m ${numbers([toXPt, toYPt])} l`);
    },
    moveTo: (xPt, yPt) => {
      write(`${numbers([xPt, yPt])} m`);
    },
    lineTo: (xPt, yPt) => {
      write(`${numbers([xPt, yPt])} l`);
    },
    curveTo: (firstXPt, firstYPt, secondXPt, secondYPt, toXPt, toYPt) => {
      write(`${numbers([firstXPt, firstYPt, secondXPt, secondYPt, toXPt, toYPt])} c`);
    },
    closePath: () => {
      write("h");
    },
    fill: () => {
      write("f");
    },
    stroke: () => {
      write("S");
    },
    fillAndStroke: () => {
      write("B");
    },
    // Cuts everything drawn after it to the path just laid down. `n` ends the path
    // without drawing it, which is what makes the rectangle a window rather than a
    // line round one.
    clip: () => {
      write("W n");
    },
    transform: (matrix) => {
      write(`${numbers(matrix)} cm`);
      set = nothingSet();
    },
    drawObject: (resource) => {
      write(`/${resource} Do`);
    },
    beginText: () => {
      write("BT");
    },
    endText: () => {
      write("ET");
    },
    font: (resource, sizePt) => {
      const stated = `/${resource} ${formatPdfNumber(sizePt)} Tf`;
      if (set.font === stated) return;
      set.font = stated;
      write(stated);
    },
    characterSpacing: (spacingPt) => {
      if (set.characterSpacing === spacingPt) return;
      set.characterSpacing = spacingPt;
      write(`${formatPdfNumber(spacingPt)} Tc`);
    },
    // **A pdf scales the glyph and not the spacing beside it**, which is the order
    // Word draws in as well: `Tz` is applied to the glyph's own advance and `Tc` is
    // added after it. Stated as a percentage, as the file states it.
    characterScale: (scale) => {
      if (set.characterScale === scale) return;
      set.characterScale = scale;
      write(`${formatPdfNumber(scale * 100)} Tz`);
    },
    // The whole matrix rather than a move, since every run is written at the place
    // layout put it rather than relative to the run before.
    textPosition: (xPt, yPt) => {
      write(`1 0 0 1 ${numbers([xPt, yPt])} Tm`);
    },
    textMatrix: (matrix) => {
      write(`${numbers(matrix)} Tm`);
    },
    showGlyphs: (glyphs) => {
      let hex = "";
      for (const byte of glyphs) hex += byte.toString(16).padStart(2, "0");
      write(`<${hex}> Tj`);
    },
    bytes: () => strToU8(`${parts.join("\n")}\n`, true),
  };
}

export type PageContent = {
  readonly page: PdfPage;
  readonly bytes: Uint8Array;
};

export type ContentOptions = {
  readonly fonts: PdfFonts;
  readonly images: PdfImages;
  readonly aliasSymbolFaces: ReadonlySet<string> | null;
};

// A picture is drawn over whatever fills its frame and under its own outline, so
// the two halves of its paint go either side of the bitmap. A picture that could
// not be drawn still gets both, which is what the frame of an unresolved one is.
function drawObject(out: Content, page: PdfPage, images: PdfImages, at: ObjectDrawable): void {
  // A drawing turned in the flow is turned about the middle of the box the flow
  // gave it, which is the box the layout answers and the box it stood in before it
  // was turned. Everything the object draws goes under the one matrix: its paint,
  // its picture and the outline over the top all turn together, as a reader of the
  // page would expect them to and as the viewer's own `rotate` does.
  if (at.turnDegrees !== 0) {
    out.save();
    out.transform(
      turnedAboutInPdf(
        at.turnDegrees,
        at.leftPt + at.widthPt / 2,
        upFromTop(page, at.topPt + at.heightPt / 2),
      ),
    );
    drawUnturnedObject(out, page, images, at);
    out.restore();
    return;
  }

  drawUnturnedObject(out, page, images, at);
}

function drawUnturnedObject(
  out: Content,
  page: PdfPage,
  images: PdfImages,
  at: ObjectDrawable,
): void {
  const { content: what } = at;

  switch (what.kind) {
    case "picture":
      paintedObject(out, page, at, { ...what.paint, outline: null });
      images.draw(out, page, at, what.part, what.crop);
      paintedObject(out, page, at, { ...what.paint, fillColor: null });
      return;
    case "text-box":
    case "shape":
      paintedObject(out, page, at, what.paint);
      return;
    // A group is flattened into one item per shape inside it by `drawablesOf`,
    // before any backend sees it, which is where the rule about what covers what
    // belongs. Nothing reaches here as a group, and drawing one would draw its
    // children twice.
    case "group":
      return;
    // A picture whose relationship names no part in the package, and a drawing of
    // a kind nothing here reads. Both draw nothing; the viewer outlines them when
    // it is asked to, and a file being written has nobody to show an outline to.
    case "missing-picture":
    case "unknown":
      return;
  }
}

/**
 * Everything one page draws, as a single content stream.
 *
 * A stream per page rather than one per layer: a pdf reader plays a page's streams
 * end to end as though they were one, so splitting them would buy nothing and cost
 * the graphics state a reader carries across the join.
 */
export function contentOf(
  layout: LaidOutDocument,
  page: PageDrawing,
  options: ContentOptions,
): PageContent {
  const pdfPage = pdfPageOf(page.geometry);
  const out = content();
  const text: TextOptions = { page: pdfPage, fonts: options.fonts };

  for (const drawable of drawablesOf(layout, page, {
    aliasSymbolFaces: options.aliasSymbolFaces,
    // The faces are this backend's to hold; where the line under a run goes is
    // not this backend's to decide.
    underlineFor: (mark) => options.fonts.faceFor(faceAskedFor(mark)).underlineAt(mark.fontSizePt),
  })) {
    switch (drawable.kind) {
      case "text": {
        const { clipTo } = drawable;
        if (clipTo === null) {
          textOfBoxes(out, text, drawable);
          break;
        }
        // Word draws the text a shape or an exact row has no room for and then
        // cuts it off there. Written and clipped rather than left out, which is
        // what Word's own pdf holds: the text is in the file and painted nowhere.
        out.save();
        // A shape's text turns with the shape, about the middle of the very
        // rectangle it is cut to. The turn goes on before the rectangle does, so
        // the cut turns with the text: the shape's box is what Word cuts at, and
        // that box is no longer square with the page once the shape has been
        // turned.
        if (drawable.turnDegrees !== 0) {
          out.transform(
            turnedAboutInPdf(
              drawable.turnDegrees,
              clipTo.leftPt + clipTo.widthPt / 2,
              upFromTop(pdfPage, clipTo.topPt + clipTo.heightPt / 2),
            ),
          );
        }
        out.rectangle(
          clipTo.leftPt,
          pdfPage.heightPt - clipTo.topPt - clipTo.heightPt,
          clipTo.widthPt,
          clipTo.heightPt,
        );
        out.clip();
        textOfBoxes(out, text, drawable);
        out.restore();
        break;
      }
      case "glyphs":
        drawnGlyphs(out, text, drawable);
        break;
      case "paint":
        paintLayer(out, pdfPage, drawable.painted, drawable.highlights);
        break;
      case "object": {
        const { clipTo } = drawable;
        if (clipTo === null) {
          drawObject(out, pdfPage, options.images, drawable);
          break;
        }
        // A line cuts off a drawing taller than itself, and Word writes the whole
        // drawing under that cut rather than leaving out what does not show.
        out.save();
        out.rectangle(
          clipTo.leftPt,
          pdfPage.heightPt - clipTo.topPt - clipTo.heightPt,
          clipTo.widthPt,
          clipTo.heightPt,
        );
        out.clip();
        drawObject(out, pdfPage, options.images, drawable);
        out.restore();
        break;
      }
    }
  }

  return { page: pdfPage, bytes: out.bytes() };
}
