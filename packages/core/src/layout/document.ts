import { readAnchors, type FloatingAnchor } from "../docx/anchors.js";
import { readInlines } from "../docx/inlines.js";
import { blockParagraphs, readBlocks, type Block } from "../docx/blocks.js";
import { defaultFooterPart, defaultHeaderPart, readRelationships } from "../docx/relationships.js";
import { MAIN_DOCUMENT_PART, type DocxPackage } from "../docx/package.js";
import {
  readUnhonoured,
  withFallbackCharacters,
  withMissingGlyphs,
  withSubstitutedFaces,
  type Unhonoured,
} from "../docx/fidelity.js";
import {
  readSectionGeometry,
  sectionsClosedBy,
  type SectionClose,
  type SectionGeometry,
} from "../docx/section.js";
import {
  honoursAWrapOnTheLeft,
  readDocumentSettings,
  takesTheRightOnEqualSides,
  type DocumentSettings,
} from "../docx/settings.js";
import { readStyleTable, type StyleTable } from "../docx/styles.js";
import type { SubstitutingMetrics } from "./substitution.js";
import { readTheme, type Theme } from "../docx/theme.js";
import {
  measureStack,
  shiftBoxes,
  shiftCells,
  type BandResolver,
  type LayoutBlocker,
  type MetricsResolver,
  type ParagraphBox,
  type PlacedCell,
} from "./stack.js";
import { breakStack, type PageStack } from "./pages.js";
import { placeFloat, type FloatSize, type PartResolver, type PlacedFloat } from "./floats.js";
import { placeInlines, type PlacedInline } from "./inlines.js";
import { layOutTextBox } from "./text-boxes.js";
import { emuToPoints, twipsToPoints } from "./units.js";
import type { BandSide, OutlinePoint, WrapBand } from "./wrapping.js";

// The header and the footer are drawn again on every page, so only the body is
// broken up: a page is the run of it that fitted between the two.
export type LaidOutPage = {
  readonly index: number;
  readonly body: readonly ParagraphBox[];
  readonly cells: readonly PlacedCell[];
  readonly floats: readonly PlacedFloat[];
  readonly inlines: readonly PlacedInline[];
};

export type LaidOutDocument = {
  readonly kind: "laid-out";
  readonly page: SectionGeometry;
  readonly headerTopPt: number;
  readonly headerHeightPt: number;
  readonly bodyTopPt: number;
  readonly bodyBottomPt: number;
  readonly footerTopPt: number;
  readonly header: readonly ParagraphBox[];
  readonly footer: readonly ParagraphBox[];
  readonly headerCells: readonly PlacedCell[];
  readonly footerCells: readonly PlacedCell[];
  readonly headerFloats: readonly PlacedFloat[];
  readonly footerFloats: readonly PlacedFloat[];
  readonly headerInlines: readonly PlacedInline[];
  readonly footerInlines: readonly PlacedInline[];
  readonly pages: readonly LaidOutPage[];
  // What the document asked for that this project passed over. A page with
  // anything here is a page Word would have drawn differently, and the entry says
  // whether that is text in the wrong place or only the wrong paint.
  readonly unhonoured: readonly Unhonoured[];
};

export type DocumentLayout =
  LaidOutDocument | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

type FloatsInPart = {
  readonly floats: readonly PlacedFloat[];
  readonly part: string;
};

type FilledFloats =
  | { readonly kind: "filled"; readonly floats: readonly (readonly PlacedFloat[])[] }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

// A text box is placed as a frame first and only then holds text, since its own
// content is laid out against the rectangle the anchor resolved to.
function fillTextBoxes(
  parts: readonly FloatsInPart[],
  styles: StyleTable,
  metricsFor: MetricsResolver,
  settings: DocumentSettings,
): FilledFloats {
  const filled: (readonly PlacedFloat[])[] = [];

  for (const { floats, part } of parts) {
    const placed: PlacedFloat[] = [];
    for (const float of floats) {
      if (float.content.kind !== "text-box") {
        placed.push(float);
        continue;
      }

      const laid = layOutTextBox({
        body: float.content.body,
        rect: float,
        styles,
        metricsFor,
        settings,
        part,
      });
      if (laid.kind === "blocked") return { kind: "blocked", blocker: laid.blocker };
      placed.push({ ...float, content: { ...float.content, text: laid.text } });
    }
    filled.push(placed);
  }

  return { kind: "filled", floats: filled };
}

type FloatFrame = {
  readonly page: SectionGeometry;
  readonly styles: StyleTable;
  readonly metricsFor: MetricsResolver;
  readonly theme: Theme;
  readonly settings: DocumentSettings;
  readonly part: string;
  // Where the story's own text column starts, which is what a column-relative
  // offset is measured from.
  readonly columnTopPt: number;
  readonly marginTopPt: number;
  // The foot of the text an object is drawn up to when it hangs past it, or null in
  // a story nothing is pulled up in. Only the body has one.
  readonly bottomPt: number | null;
};

// "Resize shape to fit text": the box is as tall as its text, and as wide as it
// too when the text does not wrap inside it. A box holding no text still fits
// itself to the paragraph mark standing in it, which Word makes one pilcrow wide
// and one line tall: measured against Word, a 270 x 0.05pt box holding one empty
// 10pt paragraph came back 5.9 x 22.2pt, and the band it wraps text out of moved
// with it.
function fittedSizePt(anchor: FloatingAnchor, frame: FloatFrame): FloatSize {
  const stored = {
    widthPt: emuToPoints(anchor.widthEmu),
    heightPt: emuToPoints(anchor.heightEmu),
  };
  const { content } = anchor;
  if (content.kind !== "text-box" || !content.body.fitsText) return stored;

  const laid = layOutTextBox({
    body: content.body,
    rect: { leftPt: 0, topPt: 0, ...stored },
    styles: frame.styles,
    metricsFor: frame.metricsFor,
    settings: frame.settings,
    part: frame.part,
  });
  if (laid.kind === "blocked") return stored;

  // The box is as tall as its text, its insets, and the outline that runs round the
  // whole of it, and as wide as the same three when the text does not wrap. Only a
  // width the file states counts: an outline that states none is still drawn, as a
  // hairline, but the box does not grow for it. Measured against Word over outlines
  // of nothing, three quarters of a point, one and a half, three and six, in both
  // wrap modes: every stated width grew the box by the whole of itself on each axis
  // being fitted, and the unstated one grew it on neither.
  const { insets, wraps } = content.body;
  const outline = content.paint.outline;
  const outlinePt = outline === null || !outline.widthStated ? 0 : outline.widthPt;
  return {
    widthPt: wraps
      ? stored.widthPt
      : laid.text.contentWidthPt +
        emuToPoints(insets.leftEmu) +
        emuToPoints(insets.rightEmu) +
        outlinePt,
    heightPt:
      laid.text.contentHeightPt +
      emuToPoints(insets.topEmu) +
      emuToPoints(insets.bottomEmu) +
      outlinePt,
  };
}

const placeFloatIn = (
  anchor: FloatingAnchor,
  paragraphTopPt: number,
  frame: FloatFrame,
  resolvePart: PartResolver,
): PlacedFloat =>
  drawnUpToTheFoot(
    placeFloat({
      anchor,
      page: frame.page,
      paragraphTopPt,
      bodyTopPt: frame.columnTopPt,
      marginTopPt: frame.marginTopPt,
      resolvePart,
      theme: frame.theme,
      settings: frame.settings,
      sizePt: fittedSizePt(anchor, frame),
    }),
    paragraphTopPt,
    frame,
  );

// **An object hanging past the foot of the text is drawn up so that its own foot
// rests there, and never higher than the paragraph anchoring it.** Measured on
// 2026-08-07 by the authored `objects-past-the-foot` document: a box hung 74pt
// below its paragraph with 156pt of room under it was drawn 38pt higher than it
// asked for, its foot exactly on the bottom of the text, and the same box with
// 100pt of room, where drawing it up that far would have taken it above its own
// anchor, moved to the next page with the paragraph instead.
//
// So this and the break rule in `breakStack` are one rule read from two sides, and
// what decides between them is whether the anchor itself leaves room: an object
// moves on exactly when it will not fit even standing at its paragraph's own top.
//
// An object wrapping nothing is not drawn up at all. It hangs where it was put,
// however far past the foot, which the same document says over three cases.
function drawnUpToTheFoot(
  float: PlacedFloat,
  paragraphTopPt: number,
  frame: FloatFrame,
): PlacedFloat {
  const bottomPt = frame.bottomPt;
  if (bottomPt === null || float.anchor.wrap === "none") return float;
  if (float.topPt + float.heightPt <= bottomPt) return float;

  const topPt = Math.max(paragraphTopPt, bottomPt - float.heightPt);
  return topPt >= float.topPt ? float : { ...float, topPt };
}

// Text stays off the part of an object its wrap covers by the distances its
// anchor asks for; an object wrapped top and bottom takes the whole width of the
// page with it.
function bandFor(float: PlacedFloat, frame: FloatFrame): WrapBand {
  const { area, distances, wrap } = float.anchor;
  const spansPage = wrap === "topAndBottom";
  const outline = spansPage ? undefined : outlineOf(float);
  const leftPt = spansPage
    ? 0
    : float.leftPt + float.widthPt * area.left - emuToPoints(distances.leftEmu);
  const rightPt = spansPage
    ? twipsToPoints(frame.page.widthTwips)
    : float.leftPt + float.widthPt * area.right + emuToPoints(distances.rightEmu);
  const side = spansPage ? undefined : sideOf(float, leftPt, rightPt, frame);

  return {
    leftPt,
    rightPt,
    topPt: float.topPt + float.heightPt * area.top - emuToPoints(distances.topEmu),
    bottomPt: float.topPt + float.heightPt * area.bottom + emuToPoints(distances.bottomEmu),
    ...(side === undefined ? {} : { side }),
    ...(outline === undefined ? {} : { outline }),
    ...(wrap === "tight" || wrap === "through" ? { outlined: true } : {}),
  };
}

/**
 * Which side of an object a line may sit on, once the object stands where it will.
 *
 * `largest` is the side of the column the object leaves the most room on, measured
 * against the two margins. Word takes the left of two equal sides in a document
 * declaring a compatibility mode and the right in one declaring none, which is one
 * of the two things the setting decides here; the other is that an old document
 * does not keep text off the right of an object wrapped on its left at all.
 *
 * The room is measured off the band rather than off the object's own frame. The
 * two are the same in everything measured, since every case put to Word held its
 * text off itself by the same distance on both sides.
 */
function sideOf(
  float: PlacedFloat,
  leftPt: number,
  rightPt: number,
  frame: FloatFrame,
): BandSide | undefined {
  const { settings } = frame;
  const side = float.anchor.side;
  if (side === "bothSides") return undefined;
  if (side === "left") return honoursAWrapOnTheLeft(settings) ? "left" : undefined;
  if (side === "right") return "right";

  const columnLeftPt = twipsToPoints(frame.page.margin.leftTwips);
  const columnRightPt = twipsToPoints(frame.page.widthTwips - frame.page.margin.rightTwips);
  const roomLeftPt = leftPt - columnLeftPt;
  const roomRightPt = columnRightPt - rightPt;
  if (Math.abs(roomLeftPt - roomRightPt) < EQUAL_SIDES_PT) {
    return takesTheRightOnEqualSides(settings) ? "right" : "left";
  }
  return roomLeftPt > roomRightPt ? "left" : "right";
}

// Two sides are the same side when they are this near each other. Positions here
// are sums of exact ratios, so this absorbs the last bits of one rather than any
// quantity Word would tell apart.
const EQUAL_SIDES_PT = 1e-9;

// A polygon that is its own rectangle says nothing the band does not already, so
// only a shape narrower than that somewhere is carried into the layout.
function outlineOf(float: PlacedFloat): readonly OutlinePoint[] | undefined {
  const { corners } = float.anchor.area;
  const rectangular = corners.every(
    (corner) =>
      (corner.x === float.anchor.area.left || corner.x === float.anchor.area.right) &&
      (corner.y === float.anchor.area.top || corner.y === float.anchor.area.bottom),
  );
  if (rectangular) return undefined;

  return corners.map((corner) => ({
    xPt: float.leftPt + float.widthPt * corner.x,
    yPt: float.topPt + float.heightPt * corner.y,
  }));
}

// Only geometry matters here, so a picture's part is left unresolved.
const bandsIn =
  (frame: FloatFrame): BandResolver =>
  (paragraph, topPt) =>
    readAnchors(paragraph)
      .filter((anchor) => anchor.wrap !== "none")
      .map((anchor) =>
        bandFor(
          placeFloatIn(anchor, topPt, frame, () => null),
          frame,
        ),
      );

type Story = {
  readonly kind: "measured";
  readonly part: string | null;
  readonly blocks: readonly Block[];
  readonly boxes: readonly ParagraphBox[];
  readonly cells: readonly PlacedCell[];
  readonly heightPt: number;
};

type StoryMeasurement = Story | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

type StoryFrame = {
  readonly styles: StyleTable;
  readonly metricsFor: MetricsResolver;
  readonly settings: DocumentSettings;
  readonly leftPt: number;
  readonly widthPt: number;
};

// Which of the body's paragraphs close a section, named by the index the stack
// knows each one as. Only a paragraph standing in the body itself closes one, so
// the blocks are read at their own level and no further.
function sectionsClosedIn(
  pkg: DocxPackage,
  blocks: readonly Block[],
): ReadonlyMap<number, SectionClose> {
  const paragraphs = blocks.flatMap((block) =>
    block.kind === "paragraph" ? [block.paragraph] : [],
  );
  const closes = sectionsClosedBy(
    pkg,
    paragraphs.map((each) => each.element),
  );
  return new Map(
    paragraphs.flatMap((each) => {
      const close = closes.get(each.element);
      return close === undefined ? [] : [[each.index, close] as const];
    }),
  );
}

const EMPTY_STORY: Story = {
  kind: "measured",
  part: null,
  blocks: [],
  boxes: [],
  cells: [],
  heightPt: 0,
};

function measureStory(
  pkg: DocxPackage,
  part: string | null,
  frame: StoryFrame,
  originPt: number,
  bandsFor: BandResolver,
): StoryMeasurement {
  if (part === null) return EMPTY_STORY;

  const blocks = readBlocks(pkg, part);
  const measured = measureStack({ ...frame, blocks, part, originPt, bandsFor });
  if (measured.kind === "blocked") return measured;
  return {
    kind: "measured",
    part,
    blocks,
    boxes: measured.boxes,
    cells: measured.cells,
    heightPt: measured.heightPt,
  };
}

/**
 * Lays a whole document out, and says what it met in the document that it could
 * not honour.
 *
 * Faces are the one thing that list cannot be read out of the package: whether a
 * face was stood in for is only known once the layout has asked for it. So a
 * resolver that stands them in is taken whole rather than as its function, and
 * what it stood in for goes into the same list as everything else. A plain
 * resolver stands in for nothing and reports nothing.
 */
export function layOutDocument(
  pkg: DocxPackage,
  metrics: MetricsResolver | SubstitutingMetrics,
): DocumentLayout {
  const faces = typeof metrics === "function" ? null : metrics;
  const metricsFor = typeof metrics === "function" ? metrics : metrics.metricsFor;
  const page = readSectionGeometry(pkg);
  const styles = readStyleTable(pkg);
  const theme = readTheme(pkg);
  const settings = readDocumentSettings(pkg);
  const headerTopPt = twipsToPoints(page.margin.headerTwips);
  const pageHeightPt = twipsToPoints(page.heightTwips);
  const leftPt = twipsToPoints(page.margin.leftTwips);
  const widthPt = twipsToPoints(page.widthTwips - page.margin.leftTwips - page.margin.rightTwips);
  const frame: StoryFrame = { styles, metricsFor, settings, leftPt, widthPt };

  const floatFrame = (
    part: string | null,
    columnTopPt: number,
    marginTopPt: number,
    bottomPt: number | null = null,
  ): FloatFrame => ({
    page,
    styles,
    metricsFor,
    theme,
    settings,
    part: part ?? MAIN_DOCUMENT_PART,
    columnTopPt,
    marginTopPt,
    bottomPt,
  });

  // A header's own objects are placed against the top of the body, which is not
  // known until the header has been measured. So it is measured once against the
  // margin the page asks for, and again against the body top that came of it.
  const headerPart = defaultHeaderPart(pkg);
  const measureHeader = (marginTopPt: number): StoryMeasurement =>
    measureStory(
      pkg,
      headerPart,
      frame,
      headerTopPt,
      bandsIn(floatFrame(headerPart, headerTopPt, marginTopPt)),
    );
  const bodyTopUnder = (story: StoryMeasurement): number =>
    Math.max(
      twipsToPoints(page.margin.topTwips),
      headerTopPt + (story.kind === "blocked" ? 0 : story.heightPt),
    );

  const firstPass = measureHeader(twipsToPoints(page.margin.topTwips));
  if (firstPass.kind === "blocked") return firstPass;
  const bodyTopPt = bodyTopUnder(firstPass);
  const header = measureHeader(bodyTopPt);
  if (header.kind === "blocked") return header;

  const headerHeightPt = header.heightPt;

  // The footer hangs from the bottom edge, so it is measured at the origin and then
  // dropped to the height it turned out to need.
  const footerPart = defaultFooterPart(pkg);
  const measuredFooter = measureStory(
    pkg,
    footerPart,
    frame,
    0,
    bandsIn(floatFrame(footerPart, 0, bodyTopPt)),
  );
  if (measuredFooter.kind === "blocked") return measuredFooter;
  const footerTopPt =
    pageHeightPt - twipsToPoints(page.margin.footerTwips) - measuredFooter.heightPt;
  const footer: Story = {
    ...measuredFooter,
    boxes: shiftBoxes(measuredFooter.boxes, footerTopPt),
    cells: shiftCells(measuredFooter.cells, footerTopPt),
  };

  const marginBottomPt = pageHeightPt - twipsToPoints(page.margin.bottomTwips);
  const bodyBottomPt =
    footer.part === null ? marginBottomPt : Math.min(marginBottomPt, footerTopPt);

  const bodyBlocks = readBlocks(pkg);
  const bodyStack = measureStack({
    ...frame,
    blocks: bodyBlocks,
    part: MAIN_DOCUMENT_PART,
    originPt: bodyTopPt,
    bandsFor: bandsIn(floatFrame(MAIN_DOCUMENT_PART, bodyTopPt, bodyTopPt)),
    sectionsClosed: sectionsClosedIn(pkg, bodyBlocks),
  });

  if (bodyStack.kind === "blocked") return { kind: "blocked", blocker: bodyStack.blocker };

  // A drawing belongs to the page its anchoring paragraph landed on, and is placed
  // against the top that paragraph has there.
  const drawingsFor = (
    blocks: readonly Block[],
    boxOf: ReadonlyMap<number, ParagraphBox>,
    floats: FloatFrame,
  ): { readonly floats: readonly PlacedFloat[]; readonly inlines: readonly PlacedInline[] } => {
    const part = floats.part;
    const relationships = readRelationships(pkg, part);
    const resolvePart = (relationshipId: string): string | null => {
      const target = relationships.get(relationshipId)?.part;
      return target !== undefined && pkg.parts.has(target) ? target : null;
    };

    const anchored = blockParagraphs(blocks).flatMap((paragraph) => {
      const box = boxOf.get(paragraph.index);
      return box === undefined ? [] : [{ paragraph, box }];
    });

    return {
      floats: anchored.flatMap(({ paragraph, box }) =>
        readAnchors(paragraph).map((anchor) =>
          placeFloatIn(anchor, box.anchorTopPt, floats, resolvePart),
        ),
      ),
      inlines: anchored.flatMap(({ paragraph, box }) =>
        placeInlines({
          drawings: readInlines(paragraph),
          box,
          resolvePart,
          theme: floats.theme,
        }),
      ),
    };
  };

  const drawingsIn = (story: Story, columnTopPt: number) =>
    story.part === null
      ? { floats: [], inlines: [] }
      : drawingsFor(
          story.blocks,
          boxesOf(story.boxes),
          floatFrame(story.part, columnTopPt, bodyTopPt),
        );

  const headerDrawings = drawingsIn(header, headerTopPt);
  const footerDrawings = drawingsIn(footer, footerTopPt);
  const broken = breakStack({
    boxes: bodyStack.boxes,
    cells: bodyStack.cells,
    untornRows: bodyStack.untornRows,
    anchoredObjects: bodyStack.anchoredObjects,
    topPt: bodyTopPt,
    bottomPt: bodyBottomPt,
  });
  const bodyDrawings = pageBoxes(broken).map((boxOf) =>
    drawingsFor(
      bodyBlocks,
      boxOf,
      floatFrame(MAIN_DOCUMENT_PART, bodyTopPt, bodyTopPt, bodyBottomPt),
    ),
  );

  const filled = fillTextBoxes(
    [
      { floats: headerDrawings.floats, part: header.part ?? MAIN_DOCUMENT_PART },
      { floats: footerDrawings.floats, part: footer.part ?? MAIN_DOCUMENT_PART },
      ...bodyDrawings.map((drawings) => ({ floats: drawings.floats, part: MAIN_DOCUMENT_PART })),
    ],
    styles,
    metricsFor,
    settings,
  );
  if (filled.kind === "blocked") return filled;
  const [headerFloats = [], footerFloats = [], ...pageFloats] = filled.floats;

  return {
    kind: "laid-out",
    page,
    unhonoured: withMissingGlyphs(
      withFallbackCharacters(
        withSubstitutedFaces(readUnhonoured(pkg), faces === null ? [] : faces.substitutions()),
        faces === null ? [] : faces.fallbackCharacters(),
      ),
      faces === null ? [] : (faces.missingGlyphs?.() ?? []),
    ),
    headerFloats,
    footerFloats,
    headerInlines: headerDrawings.inlines,
    footerInlines: footerDrawings.inlines,
    headerTopPt,
    headerHeightPt,
    bodyTopPt,
    bodyBottomPt,
    footerTopPt,
    header: header.boxes,
    footer: footer.boxes,
    headerCells: header.cells,
    footerCells: footer.cells,
    pages: broken.map((each) => ({
      index: each.index,
      body: each.boxes,
      cells: each.cells,
      floats: pageFloats[each.index] ?? [],
      inlines: bodyDrawings[each.index]?.inlines ?? [],
    })),
  };
}

const boxesOf = (boxes: readonly ParagraphBox[]): ReadonlyMap<number, ParagraphBox> =>
  new Map(boxes.map((box) => [box.index, box]));

// A paragraph the break ran through is on two pages; its drawings belong to the
// first of them, where the paragraph starts.
function pageBoxes(pages: readonly PageStack[]): readonly ReadonlyMap<number, ParagraphBox>[] {
  const seen = new Set<number>();
  return pages.map((page) => {
    const boxes = new Map<number, ParagraphBox>();
    for (const box of page.boxes) {
      if (seen.has(box.index)) continue;
      seen.add(box.index);
      boxes.set(box.index, box);
    }
    return boxes;
  });
}
