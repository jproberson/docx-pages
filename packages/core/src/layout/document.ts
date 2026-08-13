import { readAnchors, type FloatingAnchor } from "../docx/anchors.js";
import { readInlines } from "../docx/inlines.js";
import { pictureBulletOf, wearingPictureBullet } from "../docx/picture-bullets.js";
import { blockParagraphs, readBlocks, type Block } from "../docx/blocks.js";
import { readRelationships } from "../docx/relationships.js";
import { MAIN_DOCUMENT_PART, type DocxPackage } from "../docx/package.js";
import {
  readUnhonoured,
  withFallbackCharacters,
  withMissingGlyphs,
  withSubstitutedFaces,
  type Unhonoured,
} from "../docx/fidelity.js";
import {
  bodySections,
  NO_STORIES,
  readSectionGeometry,
  sectionsClosedBy,
  storyFor,
  type BodySection,
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
  type Frame,
  type LayoutBlocker,
  type MetricsResolver,
  type ParagraphBox,
  type PlacedCell,
} from "./stack.js";
import { columnsAcross } from "./columns.js";
import { anchorLineFootPt, breakStack, type PageBody, type PageStack } from "./pages.js";
import {
  placeFloat,
  type FloatSize,
  type PartResolver,
  type PlacedContent,
  type PlacedFloat,
  type PlacedGroupChild,
} from "./floats.js";
import { placeInlines, type PlacedInline } from "./inlines.js";
import { layOutTextBox } from "./text-boxes.js";
import { emuToPoints, twipsToPoints } from "./units.js";
import type { BandSide, OutlinePoint, WrapBand } from "./wrapping.js";

// A page and everything drawn on it: the run of the body that fitted, and the
// header and footer that page draws.
//
// **The header and the footer belong to the page and not to the document.** They
// stood on the document until 2026-08-10, which gave one story to every page: a
// section naming a first-page header of its own, which 408 of the 718 corpus
// documents do, had it drawn nowhere, and a document whose sections name different
// defaults had one section's given to all of them.
export type LaidOutPage = {
  readonly index: number;
  // The page the section whose text opened this one makes. Page size and margins
  // are stated per section, so they belong to the page rather than to the document:
  // a document whose sections differ draws its pages at more than one size.
  readonly geometry: SectionGeometry;
  readonly body: readonly ParagraphBox[];
  readonly cells: readonly PlacedCell[];
  readonly floats: readonly PlacedFloat[];
  readonly inlines: readonly PlacedInline[];
  readonly headerTopPt: number;
  readonly headerHeightPt: number;
  readonly footerTopPt: number;
  // What this page keeps for the body, which is not what the document keeps: the pair
  // on `LaidOutDocument` is the opening page's, and a page of another section, or one
  // drawing a header where the first drew none, starts its body somewhere else
  // entirely. Anything asking whether a line came out above the top of its own page
  // has to ask the page.
  readonly bodyTopPt: number;
  readonly bodyBottomPt: number;
  readonly header: readonly ParagraphBox[];
  readonly footer: readonly ParagraphBox[];
  readonly headerCells: readonly PlacedCell[];
  readonly footerCells: readonly PlacedCell[];
  readonly headerFloats: readonly PlacedFloat[];
  readonly footerFloats: readonly PlacedFloat[];
  readonly headerInlines: readonly PlacedInline[];
  readonly footerInlines: readonly PlacedInline[];
};

export type LaidOutDocument = {
  readonly kind: "laid-out";
  readonly page: SectionGeometry;
  readonly headerTopPt: number;
  readonly bodyTopPt: number;
  readonly bodyBottomPt: number;
  readonly pages: readonly LaidOutPage[];
  // What the document asked for that this project passed over. A page with
  // anything here is a page Word would have drawn differently, and the entry says
  // whether that is text in the wrong place or only the wrong paint.
  readonly unhonoured: readonly Unhonoured[];
};

export type DocumentLayout =
  LaidOutDocument | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

// The paragraphs a box holds, at whatever depth it is buried. A group holds shapes and
// a shape holds text, and a group inside a group is the same again: the labels on a
// diagram are shapes of exactly that kind.
const boxesHeldBy = (content: PlacedContent): readonly ParagraphBox[] => {
  if (content.kind === "group")
    return content.children.flatMap((child) => boxesHeldBy(child.content));
  return content.kind === "text-box" && content.text !== null ? content.text.boxes : [];
};

/**
 * Every paragraph whose text this page draws: the flow's, the header's and the footer's,
 * and every one standing in a box, a shape, or a shape inside a group.
 *
 * **Anything comparing a page against something else has to walk this and not
 * `page.body`.** `drawablesOf` flattens a group so no renderer ever learns that groups
 * exist, and every reading that went looking for text by hand has had the same hole in
 * it instead: `pdf/agreement.ts` walked top-level text boxes and no group, so on
 * 2026-08-12 a page whose title block is a group of two shapes was reported as content
 * Word drew and we did not, while the raster said the page was Word's cell for cell.
 * Before that, `corpus/inspect.ts` read a page built out of text boxes as an empty page
 * and six documents were said never to have been read at all.
 */
export function paragraphBoxesOn(page: LaidOutPage): readonly ParagraphBox[] {
  const objects = [
    ...page.headerFloats,
    ...page.footerFloats,
    ...page.floats,
    ...page.headerInlines,
    ...page.inlines,
    ...page.footerInlines,
  ];
  return [
    ...page.header,
    ...page.footer,
    ...page.body,
    ...objects.flatMap((object) => boxesHeldBy(object.content)),
  ];
}

type DrawingsInPart = {
  readonly floats: readonly PlacedFloat[];
  readonly inlines: readonly PlacedInline[];
  readonly part: string;
};

type FilledDrawings =
  | {
      readonly kind: "filled";
      readonly floats: readonly (readonly PlacedFloat[])[];
      readonly inlines: readonly (readonly PlacedInline[])[];
    }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

type Rect = {
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
};

type FilledContent =
  | { readonly kind: "filled"; readonly content: PlacedContent }
  | { readonly kind: "blocked"; readonly blocker: LayoutBlocker };

/**
 * A text box is placed as a frame first and only then holds text, since its own
 * content is laid out against the rectangle the anchor resolved to.
 *
 * **A shape inside a group holds text the same way**, and its rectangle is the
 * fraction of the group's box it stands in. The labels on a diagram are shapes of
 * exactly that kind, so a group whose text was never laid out draws its boxes and
 * none of its words.
 */
function fillContent(
  content: PlacedContent,
  rect: Rect,
  styles: StyleTable,
  metricsFor: MetricsResolver,
  settings: DocumentSettings,
  part: string,
): FilledContent {
  if (content.kind === "group") {
    const children: PlacedGroupChild[] = [];
    for (const child of content.children) {
      const filled = fillContent(
        child.content,
        {
          leftPt: rect.leftPt + child.leftFraction * rect.widthPt,
          topPt: rect.topPt + child.topFraction * rect.heightPt,
          widthPt: child.widthFraction * rect.widthPt,
          heightPt: child.heightFraction * rect.heightPt,
        },
        styles,
        metricsFor,
        settings,
        part,
      );
      if (filled.kind === "blocked") return filled;
      children.push({ ...child, content: filled.content });
    }
    return { kind: "filled", content: { kind: "group", children } };
  }

  if (content.kind !== "text-box") return { kind: "filled", content };

  const laid = layOutTextBox({ body: content.body, rect, styles, metricsFor, settings, part });
  if (laid.kind === "blocked") return { kind: "blocked", blocker: laid.blocker };
  return { kind: "filled", content: { ...content, text: laid.text } };
}

// **An inline drawing is filled here too, and until 2026-08-10 nothing filled
// one.** It never showed, because the only text a shape held was in a box
// somebody had anchored; a group of shapes is inline as often as not, and the
// labels on a diagram are shapes inside it holding text.
function fillTextBoxes(
  parts: readonly DrawingsInPart[],
  styles: StyleTable,
  metricsFor: MetricsResolver,
  settings: DocumentSettings,
): FilledDrawings {
  const filledFloats: (readonly PlacedFloat[])[] = [];
  const filledInlines: (readonly PlacedInline[])[] = [];

  for (const { floats, inlines, part } of parts) {
    const placedFloats: PlacedFloat[] = [];
    for (const float of floats) {
      const content = fillContent(float.content, float, styles, metricsFor, settings, part);
      if (content.kind === "blocked") return { kind: "blocked", blocker: content.blocker };
      placedFloats.push({ ...float, content: content.content });
    }
    filledFloats.push(placedFloats);

    const placedInlines: PlacedInline[] = [];
    for (const inline of inlines) {
      const content = fillContent(inline.content, inline, styles, metricsFor, settings, part);
      if (content.kind === "blocked") return { kind: "blocked", blocker: content.blocker };
      placedInlines.push({ ...inline, content: content.content });
    }
    filledInlines.push(placedInlines);
  }

  return { kind: "filled", floats: filledFloats, inlines: filledInlines };
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
  lineFootPt: number,
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
    lineFootPt,
    frame,
  );

// **An object hanging past the foot of the text is drawn up so that its own foot
// rests there, and never over the line anchoring it.** Measured on 2026-08-07 by
// the authored `objects-past-the-foot` document: a box hung 74pt below its
// paragraph with 156pt of room under it was drawn 38pt higher than it asked for,
// its foot exactly on the bottom of the text, and the same box with 100pt of room,
// where drawing it up that far would have taken it above its own anchor, moved to
// the next page with the paragraph instead.
//
// How far up it may come is the foot of that line and not the paragraph's own top,
// measured on 2026-08-08 by `objects-and-the-footer`: a box that would have had to
// rise to 2pt above the foot of its line moved instead of being drawn up.
//
// So this and the break rule in `breakStack` are one rule read from two sides, and
// what decides between them is whether the line itself leaves room: an object moves
// on exactly when it will not fit even standing at the foot of its own line.
//
// An object wrapping nothing is not drawn up at all. It hangs where it was put,
// however far past the foot, which the same document says over three cases.
function drawnUpToTheFoot(float: PlacedFloat, lineFootPt: number, frame: FloatFrame): PlacedFloat {
  const bottomPt = frame.bottomPt;
  if (bottomPt === null || float.anchor.wrap === "none") return float;
  if (float.topPt + float.heightPt <= bottomPt) return float;

  const topPt = Math.max(lineFootPt, bottomPt - float.heightPt);
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

// Only geometry matters here, so a picture's part is left unresolved, and the line
// an object is anchored to has not been laid out yet: measuring is one column with
// no foot for one to be drawn up to, so the anchor's own top stands in for it.
const bandsIn =
  (frame: FloatFrame): BandResolver =>
  (paragraph, topPt) =>
    readAnchors(paragraph)
      .filter((anchor) => anchor.wrap !== "none")
      .map((anchor) =>
        bandFor(
          placeFloatIn(anchor, topPt, topPt, frame, () => null),
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

// A story as one page draws it: where it hangs, and the drawings its own
// paragraphs anchor.
type LaidStory = Story & {
  readonly topPt: number;
  readonly floats: readonly PlacedFloat[];
  readonly inlines: readonly PlacedInline[];
};

// What a page naming no header or footer draws, which is nothing at all.
const NOTHING_DRAWN: LaidStory = {
  kind: "measured",
  part: null,
  blocks: [],
  boxes: [],
  cells: [],
  heightPt: 0,
  topPt: 0,
  floats: [],
  inlines: [],
};

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

// Which section each of the body's own blocks stands in. A section runs up to and
// including the paragraph carrying its properties, so the section advances after
// that block rather than at it.
function sectionOfEachBlock(
  pkg: DocxPackage,
  blocks: readonly Block[],
): ReadonlyMap<Block, BodySection> {
  const paragraphs = blocks.flatMap((block) =>
    block.kind === "paragraph" ? [block.paragraph] : [],
  );
  const sections = bodySections(
    pkg,
    paragraphs.map((each) => each.element),
  );

  const standing = new Map<Block, BodySection>();
  let at = 0;
  for (const block of blocks) {
    const section = sections[at] ?? sections[sections.length - 1];
    if (section === undefined) break;
    standing.set(block, section);
    if (block.kind === "paragraph" && block.paragraph.element === section.endsAt) at += 1;
  }
  return standing;
}

// Which section each of the body's paragraphs stands in, by the index the stack
// knows it as. The indices run in document order through a table's cells as well as
// over the body's own paragraphs, so a section takes in every index from the block
// it opens at up to the block the next section opens at.
type SectionsByParagraph = {
  readonly at: (index: number) => BodySection | undefined;
  // The paragraph a section starts at, which is what says whether a page is the
  // section's first: a section that started part way down a page had its first page
  // there, whatever the page below it holds.
  readonly opensAt: (index: number) => number | null;
};

function sectionAtEachParagraph(
  blocks: readonly Block[],
  sectionOf: ReadonlyMap<Block, BodySection>,
): SectionsByParagraph {
  const opens: { readonly index: number; readonly section: BodySection }[] = [];
  for (const block of blocks) {
    const section = sectionOf.get(block);
    const index = blockParagraphs([block])[0]?.index;
    if (section === undefined || index === undefined) continue;
    if (opens[opens.length - 1]?.section === section) continue;
    opens.push({ index, section });
  }

  const governing = (index: number): { index: number; section: BodySection } | null => {
    let found: { index: number; section: BodySection } | null = null;
    for (const each of opens) {
      if (each.index > index) break;
      found = each;
    }
    return found;
  };

  return {
    at: (index) => governing(index)?.section,
    opensAt: (index) => governing(index)?.index ?? null,
  };
}

// Where a section's text runs across the page. A continuous section changes this
// partway down a page, which is the whole of what one does that a page break does
// not: two sections an inch apart in their left margins were drawn on one page, at
// 72pt and at 144pt.
const frameOfSection = (section: BodySection): Frame => ({
  leftPt: twipsToPoints(section.geometry.margin.leftTwips),
  widthPt: twipsToPoints(
    section.geometry.widthTwips -
      section.geometry.margin.leftTwips -
      section.geometry.margin.rightTwips,
  ),
});

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
  const leftPt = twipsToPoints(page.margin.leftTwips);
  const widthPt = twipsToPoints(page.widthTwips - page.margin.leftTwips - page.margin.rightTwips);
  const frame: StoryFrame = { styles, metricsFor, settings, leftPt, widthPt };

  const pictureBullet = pictureBulletOf(pkg);
  const bodyBlocks =
    pictureBullet === null ? readBlocks(pkg) : wearingPictureBullet(readBlocks(pkg), pictureBullet);
  const bodySectionOf = sectionOfEachBlock(pkg, bodyBlocks);
  const sections = sectionAtEachParagraph(bodyBlocks, bodySectionOf);
  const sectionAt = sections.at;
  const geometryAt = (index: number | null): SectionGeometry =>
    (index === null ? undefined : sectionAt(index)?.geometry) ?? page;
  // The page the body opens on, which the first section makes and not the last.
  // The whole stack is measured from its top, and every page after the first is
  // shifted back to the top of its own.
  const openingPage = geometryAt(blockParagraphs(bodyBlocks)[0]?.index ?? null);

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

  // **A page draws its own section's header and footer, and the page a section
  // opens draws that section's first-page ones where it says `w:titlePg`.** What
  // stood here read the last `w:headerReference` of type `default` anywhere in the
  // part and gave that one story to every page of the document.
  //
  // 408 of the 718 corpus documents state `w:titlePg` and 226 hold sections naming
  // different defaults, so the old reading was one section's answer handed to
  // everybody. A one-page corpus document whose section names nothing but a
  // first-page header was drawn under the *next* section's default, which is a
  // background image the size of the sheet: the page came out under a grey grid,
  // without the logo and rule its own header holds.
  const storiesOf = (section: BodySection | undefined, opensItsSection: boolean) => ({
    header: storyFor(section?.headers ?? NO_STORIES, opensItsSection, section?.titlePage ?? false),
    footer: storyFor(section?.footers ?? NO_STORIES, opensItsSection, section?.titlePage ?? false),
  });

  // **How far down a page its header starts is its own section's business**, as
  // the top margin and the footer already were. Measured on 2026-08-10 against
  // Word's own drawing of 506 corpus documents that hang a picture in a header:
  // 404 of their 556 header pictures were exactly 31.5pt below where Word drew
  // them, and 150 were exactly right. The 404 are a section keeping 90 twips for
  // its header under a document whose last section keeps 720, and 36.0 less 4.5 is
  // the 31.5. Their body already stood where Word put it, so nothing but the
  // header itself was out.
  const headerTopOf = (geometry: SectionGeometry): number =>
    twipsToPoints(geometry.margin.headerTwips);

  // A header's own objects are placed against the top of the body, which is not
  // known until the header has been measured. So it is measured once against the
  // margin the page asks for, and again against the body top that came of it. Only
  // the first pass is needed to say how tall a header is, and a document draws few
  // enough of them that measuring each one once is nothing.
  //
  // Keyed by the top as well as by the part, since a part drawn under two sections
  // that keep different room for a header is two different measurements.
  const headerHeights = new Map<string, StoryMeasurement>();
  const heightOfHeader = (part: string | null, topPt: number): StoryMeasurement => {
    const key = `${String(topPt)}|${part ?? ""}`;
    const known = part === null ? undefined : headerHeights.get(key);
    if (known !== undefined) return known;
    const measured = measureStory(
      pkg,
      part,
      frame,
      topPt,
      bandsIn(floatFrame(part, topPt, twipsToPoints(page.margin.topTwips))),
    );
    if (part !== null) headerHeights.set(key, measured);
    return measured;
  };

  // Where a section's own page starts the body: at its top margin, or under a
  // header that reaches past it.
  //
  // **A page that draws no header starts its body at the top margin**, however far
  // down the page the room kept for a header would have reached. The room a header
  // would have taken is not taken by a header that is not there, which is what the
  // footer has always said on its own side. Four corpus documents keeping a 36pt
  // header margin over a 1pt top margin and drawing no header had every line of
  // them 35pt below where Word drew it, and every one of the four was read as a
  // fault about the columns it also states.
  const bodyTopOf = (geometry: SectionGeometry, headerPart: string | null): number => {
    const marginTopPt = twipsToPoints(geometry.margin.topTwips);
    if (headerPart === null) return marginTopPt;
    const measured = heightOfHeader(headerPart, headerTopOf(geometry));
    if (measured.kind === "blocked") return marginTopPt;
    return Math.max(marginTopPt, headerTopOf(geometry) + measured.heightPt);
  };

  // Whether the page a paragraph opens is the first of its section, which is the
  // whole of what `w:titlePg` turns on. A page opened by a paragraph standing in
  // another section than the page above it opens that section, and the first page of
  // the document opens after nothing.
  const opensASection = (index: number | null): boolean =>
    index !== null && sections.opensAt(index) === index;

  const openingSection = sectionAt(blockParagraphs(bodyBlocks)[0]?.index ?? -1);
  const openingStories = storiesOf(openingSection, true);
  const bodyTopPt = bodyTopOf(openingPage, openingStories.header);

  // The footer hangs from the bottom edge, so it is measured at the origin and then
  // dropped to the height it turned out to need.
  const footers = new Map<string, StoryMeasurement>();
  const measureFooter = (part: string | null): StoryMeasurement => {
    const known = part === null ? undefined : footers.get(part);
    if (known !== undefined) return known;
    const measured = measureStory(pkg, part, frame, 0, bandsIn(floatFrame(part, 0, bodyTopPt)));
    if (part !== null) footers.set(part, measured);
    return measured;
  };

  // Where a section's own page hangs the footer, which is as far above the bottom
  // edge as that section keeps for one.
  const footerTopOf = (geometry: SectionGeometry, footerPart: string | null): number => {
    const measured = measureFooter(footerPart);
    return (
      twipsToPoints(geometry.heightTwips) -
      twipsToPoints(geometry.margin.footerTwips) -
      (measured.kind === "blocked" ? 0 : measured.heightPt)
    );
  };

  // What a page a section opens keeps for the body of the text. Page size and
  // margins are stated per section, so a document whose sections differ breaks each
  // page against the geometry of the one whose text opened it, and now against the
  // header and footer that page draws as well.
  const bodyOfPage = (section: BodySection | undefined, opensItsSection: boolean): PageBody => {
    const geometry = section?.geometry ?? page;
    const stories = storiesOf(section, opensItsSection);
    const marginBottomPt =
      twipsToPoints(geometry.heightTwips) - twipsToPoints(geometry.margin.bottomTwips);
    return {
      topPt: bodyTopOf(geometry, stories.header),
      bottomPt:
        stories.footer === null
          ? marginBottomPt
          : Math.min(marginBottomPt, footerTopOf(geometry, stories.footer)),
    };
  };

  const bodyBottomPt = bodyOfPage(openingSection, true).bottomPt;

  const bodyStack = measureStack({
    ...frame,
    blocks: bodyBlocks,
    part: MAIN_DOCUMENT_PART,
    originPt: bodyTopPt,
    bandsFor: bandsIn(floatFrame(MAIN_DOCUMENT_PART, bodyTopPt, bodyTopPt)),
    sectionsClosed: sectionsClosedIn(pkg, bodyBlocks),
    bodyHeightPt: bodyBottomPt - bodyTopPt,
    frameOf: (block) => {
      const section = bodySectionOf.get(block);
      return section === undefined ? undefined : frameOfSection(section);
    },
    columnsOf: (block) => {
      const section = bodySectionOf.get(block);
      if (section === undefined) return [];
      return columnsAcross(section.columns, frameOfSection(section));
    },
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
      if (pictureBullet !== null && relationshipId === pictureBullet.relationshipId) {
        return pictureBullet.part;
      }
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
          placeFloatIn(anchor, box.anchorTopPt, anchorLineFootPt(box), floats, resolvePart),
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

  const broken = breakStack({
    boxes: bodyStack.boxes,
    cells: bodyStack.cells,
    untornRows: bodyStack.untornRows,
    anchoredObjects: bodyStack.anchoredObjects,
    topPt: bodyTopPt,
    bottomPt: bodyBottomPt,
    bodyOf: (box) => bodyOfPage(sectionAt(box.index), opensASection(box.index)),
  });
  const bodyDrawings = pageBoxes(broken).map((boxOf) =>
    drawingsFor(
      bodyBlocks,
      boxOf,
      floatFrame(MAIN_DOCUMENT_PART, bodyTopPt, bodyTopPt, bodyBottomPt),
    ),
  );

  // The header and the footer each page draws, measured once for each part at each
  // height and then hung where that page keeps it: a document naming three headers
  // draws three, and a page names which of them it took.
  //
  // **Keyed by where the page hangs the story as well as by the part**, since the
  // room a section keeps above its header and below its footer is stated per
  // section: one part drawn under two such sections is two different drawings, and
  // keying by the part alone gave whichever section asked first to everybody.
  const storyKey = (part: string | null, topPt: number): string => `${String(topPt)}|${part ?? ""}`;

  const laidHeaders = new Map<string, LaidStory>();
  const headerOn = (part: string | null, geometry: SectionGeometry): LaidStory => {
    const topPt = headerTopOf(geometry);
    const known = laidHeaders.get(storyKey(part, topPt));
    if (known !== undefined) return known;
    const measured = measureStory(
      pkg,
      part,
      frame,
      topPt,
      bandsIn(floatFrame(part, topPt, bodyTopPt)),
    );
    const story: LaidStory =
      measured.kind === "blocked"
        ? NOTHING_DRAWN
        : { ...measured, topPt, ...drawingsIn(measured, topPt) };
    laidHeaders.set(storyKey(part, topPt), story);
    return story;
  };

  const laidFooters = new Map<string, LaidStory>();
  const footerOn = (part: string | null, geometry: SectionGeometry): LaidStory => {
    const topPt = footerTopOf(geometry, part);
    const known = laidFooters.get(storyKey(part, topPt));
    if (known !== undefined) return known;
    const measured = measureFooter(part);
    if (measured.kind === "blocked") {
      laidFooters.set(storyKey(part, topPt), NOTHING_DRAWN);
      return NOTHING_DRAWN;
    }
    const dropped: Story = {
      ...measured,
      boxes: shiftBoxes(measured.boxes, topPt),
      cells: shiftCells(measured.cells, topPt),
    };
    const story: LaidStory = { ...dropped, topPt, ...drawingsIn(dropped, topPt) };
    laidFooters.set(storyKey(part, topPt), story);
    return story;
  };

  // Which header and footer each page took, keyed the way the drawings below are
  // gathered: a story is filled once however many pages draw it.
  const drawnOn = broken.map((each) => {
    const section = sectionAt(each.openedBy ?? -1);
    const opensItsSection = opensASection(each.openedBy);
    const stories = storiesOf(section, opensItsSection);
    const geometry = section?.geometry ?? page;
    return {
      header: headerOn(stories.header, geometry),
      footer: footerOn(stories.footer, geometry),
      // The same answer the break was given for this page, kept rather than thrown
      // away: it is what says where the page's own body begins and ends.
      body: bodyOfPage(section, opensItsSection),
    };
  });

  const stories = [...new Set([...laidHeaders.values(), ...laidFooters.values()])];
  const filled = fillTextBoxes(
    [
      ...stories.map((story) => ({
        floats: story.floats,
        inlines: story.inlines,
        part: story.part ?? MAIN_DOCUMENT_PART,
      })),
      ...bodyDrawings.map((drawings) => ({
        floats: drawings.floats,
        inlines: drawings.inlines,
        part: MAIN_DOCUMENT_PART,
      })),
    ],
    styles,
    metricsFor,
    settings,
  );
  if (filled.kind === "blocked") return filled;
  const floatsOfStory = new Map(
    stories.map((story, at) => [story, filled.floats[at] ?? []] as const),
  );
  const inlinesOfStory = new Map(
    stories.map((story, at) => [story, filled.inlines[at] ?? []] as const),
  );
  const pageFloats = filled.floats.slice(stories.length);
  const pageInlines = filled.inlines.slice(stories.length);

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
    headerTopPt: headerTopOf(openingPage),
    bodyTopPt,
    bodyBottomPt,
    pages: broken.map((each) => {
      const drawn = drawnOn[each.index] ?? {
        header: NOTHING_DRAWN,
        footer: NOTHING_DRAWN,
        body: { topPt: bodyTopPt, bottomPt: bodyBottomPt },
      };
      return {
        index: each.index,
        geometry: geometryAt(each.openedBy),
        body: each.boxes,
        cells: each.cells,
        floats: pageFloats[each.index] ?? [],
        inlines: pageInlines[each.index] ?? [],
        headerTopPt: headerTopOf(geometryAt(each.openedBy)),
        headerHeightPt: drawn.header.heightPt,
        footerTopPt: drawn.footer.topPt,
        bodyTopPt: drawn.body.topPt,
        bodyBottomPt: drawn.body.bottomPt,
        header: drawn.header.boxes,
        footer: drawn.footer.boxes,
        headerCells: drawn.header.cells,
        footerCells: drawn.footer.cells,
        headerFloats: floatsOfStory.get(drawn.header) ?? [],
        footerFloats: floatsOfStory.get(drawn.footer) ?? [],
        headerInlines: inlinesOfStory.get(drawn.header) ?? [],
        footerInlines: inlinesOfStory.get(drawn.footer) ?? [],
      };
    }),
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
