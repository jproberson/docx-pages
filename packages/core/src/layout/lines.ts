import type { RunPiece, TextRun } from "../docx/runs.js";
import type { ParagraphMark } from "../docx/styles.js";
import {
  advanceWidthPt,
  ascentPt,
  lineHeightPt,
  type AdvancesUnavailable,
  type FaceRequest,
  type FontMetrics,
  type GlyphAdvances,
  type MetricsLookup,
} from "./font-metrics.js";
import { nextTabStopPt } from "./tab-stops.js";
import { emuToPoints } from "./units.js";

export type MetricsResolver = (request: FaceRequest) => MetricsLookup;

// Where the run sits along its line, measured from the line's own start. A tab
// opens a gap the runs after it never account for, so each one carries the place
// it reached rather than leaving it to be added up.
export type LineSegment =
  | {
      readonly kind: "text";
      readonly mark: ParagraphMark;
      readonly text: string;
      readonly widthPt: number;
      readonly offsetPt: number;
    }
  | {
      readonly kind: "drawing";
      readonly widthPt: number;
      readonly heightPt: number;
      readonly offsetPt: number;
    };

export type TextLine = {
  readonly segments: readonly LineSegment[];
  readonly widthPt: number;
  readonly heightPt: number;
  readonly ascentPt: number;
};

export type MeasureFailure =
  | { readonly kind: "unresolved-font" }
  | { readonly kind: "unknown-font-metrics"; readonly fontName: string }
  | {
      readonly kind: "unmeasurable-text";
      readonly fontName: string;
      readonly reason: AdvancesUnavailable;
    }
  | {
      readonly kind: "unmapped-character";
      readonly fontName: string;
      readonly codePoint: number;
    };

export type LineBreaking =
  | { readonly kind: "lines"; readonly lines: readonly TextLine[] }
  | { readonly kind: "unmeasurable"; readonly failure: MeasureFailure };

// Tab stops are measured from the left edge of the text area, so a line has to say
// how far its own start sits from that edge for a tab to land on the right one. A
// hanging first line starts outside that edge, and a tab on it reaches the stop at
// the indent that the lines below it start from.
export type LineTabs = {
  readonly stopsPt: readonly number[];
  readonly originPt: number;
  readonly firstLineOriginPt?: number;
};

export type BreakLinesInput = {
  readonly runs: readonly TextRun[];
  readonly widthPt: number;
  // What the first line alone has room for, which a hanging indent makes wider
  // than the lines under it and a first-line indent makes narrower.
  readonly firstLineWidthPt?: number;
  readonly metricsFor: MetricsResolver;
  readonly tabs?: LineTabs;
};

const NO_TABS: LineTabs = { stopsPt: [], originPt: 0 };

// Widths are compared, not accumulated into a coordinate, so this only has to
// absorb the last bits of a sum of exact ratios.
const EPSILON = 1e-9;

export const faceRequestFor = (mark: ParagraphMark): FaceRequest => ({
  name: mark.font.kind === "named" ? mark.font.name : "",
  bold: mark.bold,
  italic: mark.italic,
});

type Face = {
  readonly metrics: FontMetrics;
  readonly advanceFor: GlyphAdvances;
};

type Fragment = {
  readonly mark: ParagraphMark;
  readonly text: string;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly ascentPt: number;
};

type Unit =
  | { readonly kind: "word" | "space"; readonly fragments: readonly Fragment[] }
  | { readonly kind: "tab" }
  | { readonly kind: "break" }
  | { readonly kind: "drawing"; readonly widthPt: number; readonly heightPt: number };

type Measured<T> = { readonly kind: "ok"; readonly value: T } | { readonly kind: "failed" };

// Word advances one code point at a time, so splitting there is the right
// granularity for measuring even though it is not grapheme-aware.
const charactersOf = (text: string): readonly string[] => Array.from(text);

const widthOf = (fragments: readonly Fragment[]): number =>
  fragments.reduce((sum, fragment) => sum + fragment.widthPt, 0);

class Measurer {
  private readonly faces = new Map<ParagraphMark, Face>();
  failure: MeasureFailure | null = null;

  constructor(private readonly metricsFor: MetricsResolver) {}

  private faceFor(mark: ParagraphMark): Face | null {
    const cached = this.faces.get(mark);
    if (cached !== undefined) return cached;

    if (mark.font.kind === "unresolved") {
      this.failure ??= { kind: "unresolved-font" };
      return null;
    }

    const fontName = mark.font.name;
    const lookup = this.metricsFor(faceRequestFor(mark));
    if (lookup.kind === "missing") {
      this.failure ??= { kind: "unknown-font-metrics", fontName };
      return null;
    }
    if (lookup.advances.kind !== "advances") {
      this.failure ??= { kind: "unmeasurable-text", fontName, reason: lookup.advances.reason };
      return null;
    }

    const face = { metrics: lookup.metrics, advanceFor: lookup.advances.advanceFor };
    this.faces.set(mark, face);
    return face;
  }

  fragment(mark: ParagraphMark, text: string): Fragment | null {
    const face = this.faceFor(mark);
    if (face === null) return null;

    let widthPt = 0;
    for (const character of text) {
      const codePoint = character.codePointAt(0) ?? 0;
      const advance = face.advanceFor(codePoint);
      if (advance === null) {
        this.failure ??= {
          kind: "unmapped-character",
          fontName: mark.font.kind === "named" ? mark.font.name : "",
          codePoint,
        };
        return null;
      }
      widthPt += advanceWidthPt(advance, face.metrics, mark.fontSizePt);
    }

    return {
      mark,
      text,
      widthPt,
      heightPt: lineHeightPt(face.metrics, mark.fontSizePt),
      ascentPt: ascentPt(face.metrics, mark.fontSizePt),
    };
  }
}

// A no-break space is what its name says: text runs on through it, so it belongs
// to the word around it rather than opening a place the line can break. Every
// other run of whitespace is a gap between words.
const GAP = /([^\S\u00a0]+)/;
const IS_GAP = /^[^\S\u00a0]+$/;

// Word lets a line end on a hyphen inside a word, so each one closes the word it
// belongs to and the rest of it becomes a word of its own.
const AFTER_HYPHEN = /(?<=-)/;

const endsOnHyphen = (unit: Unit): boolean =>
  unit.kind === "word" && (unit.fragments.at(-1)?.text ?? "").endsWith("-");

function tokenize(runs: readonly TextRun[], measurer: Measurer): Measured<readonly Unit[]> {
  const units: Unit[] = [];

  const append = (kind: "word" | "space", fragment: Fragment): void => {
    const last = units.at(-1);
    if (last !== undefined && last.kind === kind && !endsOnHyphen(last)) {
      units[units.length - 1] = { kind, fragments: [...last.fragments, fragment] };
      return;
    }
    units.push({ kind, fragments: [fragment] });
  };

  for (const run of runs) {
    for (const piece of run.pieces) {
      if (!addPiece(piece, run.mark, units, append, measurer)) return { kind: "failed" };
    }
  }

  return { kind: "ok", value: units };
}

function addPiece(
  piece: RunPiece,
  mark: ParagraphMark,
  units: Unit[],
  append: (kind: "word" | "space", fragment: Fragment) => void,
  measurer: Measurer,
): boolean {
  if (piece.kind === "tab") {
    units.push({ kind: "tab" });
    return true;
  }
  if (piece.kind === "break") {
    units.push({ kind: "break" });
    return true;
  }
  if (piece.kind === "drawing") {
    units.push({
      kind: "drawing",
      widthPt: emuToPoints(piece.widthEmu),
      heightPt: emuToPoints(piece.heightEmu),
    });
    return true;
  }

  for (const token of piece.text.split(GAP).filter((each) => each !== "")) {
    const space = IS_GAP.test(token);
    for (const part of space ? [token] : token.split(AFTER_HYPHEN)) {
      const fragment = measurer.fragment(mark, part);
      if (fragment === null) return false;
      append(space ? "space" : "word", fragment);
    }
  }
  return true;
}

// Word breaks greedily: it fills a line until the next word will not fit, then
// starts a new one, and the spaces it broke at hang past the edge rather than
// opening the next line.
class Breaker {
  private readonly lines: TextLine[] = [];
  private segments: LineSegment[] = [];
  private committedPt = 0;
  private pending: LineSegment[] = [];
  private pendingPt = 0;
  private heightPt = 0;
  private ascentPt = 0;
  private wrapped = false;
  private tabbed = false;

  constructor(
    private readonly widthPt: number,
    private readonly firstLineWidthPt: number,
    private readonly tabs: LineTabs,
  ) {}

  private get room(): number {
    return this.lines.length === 0 ? this.firstLineWidthPt : this.widthPt;
  }

  private raise(heightPt: number, ascentPt: number): void {
    this.heightPt = Math.max(this.heightPt, heightPt);
    this.ascentPt = Math.max(this.ascentPt, ascentPt);
  }

  private commit(segments: readonly LineSegment[], widthPt: number): void {
    let offsetPt = this.committedPt + this.pendingPt;
    this.segments.push(...this.pending);
    for (const segment of segments) {
      this.segments.push(startingAt(segment, offsetPt));
      offsetPt += segment.widthPt;
    }
    this.committedPt += this.pendingPt + widthPt;
    this.pending = [];
    this.pendingPt = 0;
    this.wrapped = false;
    this.tabbed = false;
  }

  private get filled(): number {
    return this.committedPt + this.pendingPt;
  }

  private get empty(): boolean {
    return this.segments.length === 0;
  }

  // A trailing space hangs past the edge, but a trailing tab holds the line open
  // as far as the stop it reached: Word wraps that line around the width the tab
  // gave it, even when nothing is drawn there.
  flush(): void {
    if (!this.empty || this.tabbed) {
      this.lines.push({
        segments: this.segments,
        widthPt: this.committedPt + (this.tabbed ? this.pendingPt : 0),
        heightPt: this.heightPt,
        ascentPt: this.ascentPt,
      });
    }
    this.segments = [];
    this.committedPt = 0;
    this.pending = [];
    this.pendingPt = 0;
    this.heightPt = 0;
    this.ascentPt = 0;
    this.tabbed = false;
  }

  space(fragments: readonly Fragment[]): void {
    if (this.empty && this.wrapped) return;
    for (const fragment of fragments) {
      this.pending.push(startingAt(segmentOf(fragment), this.committedPt + this.pendingPt));
      this.pendingPt += fragment.widthPt;
      this.raise(fragment.heightPt, fragment.ascentPt);
    }
  }

  tab(): void {
    if (this.empty && this.wrapped) return;
    const { stopsPt } = this.tabs;
    const originPt =
      this.lines.length === 0
        ? (this.tabs.firstLineOriginPt ?? this.tabs.originPt)
        : this.tabs.originPt;
    this.pendingPt = nextTabStopPt(originPt + this.filled, stopsPt) - originPt - this.committedPt;
    this.tabbed = true;
  }

  drawing(widthPt: number, heightPt: number): void {
    if (!this.empty && this.filled + widthPt > this.room + EPSILON) this.wrap();
    this.raise(heightPt, heightPt);
    this.commit([{ kind: "drawing", widthPt, heightPt, offsetPt: 0 }], widthPt);
  }

  word(fragments: readonly Fragment[]): void {
    let rest = fragments;
    while (rest.length > 0) {
      const widthPt = widthOf(rest);
      if (this.filled + widthPt <= this.room + EPSILON) {
        this.take(rest);
        return;
      }
      if (!this.empty) {
        this.wrap();
        continue;
      }
      const [head, tail] = splitFragments(rest, this.room - this.filled);
      this.take(head);
      this.wrap();
      rest = tail;
    }
  }

  private take(fragments: readonly Fragment[]): void {
    for (const fragment of fragments) this.raise(fragment.heightPt, fragment.ascentPt);
    this.commit(fragments.map(segmentOf), widthOf(fragments));
  }

  private wrap(): void {
    this.flush();
    this.wrapped = true;
  }

  finish(): readonly TextLine[] {
    this.flush();
    return this.lines;
  }
}

const segmentOf = (fragment: Fragment): LineSegment => ({
  kind: "text",
  mark: fragment.mark,
  text: fragment.text,
  widthPt: fragment.widthPt,
  offsetPt: 0,
});

const startingAt = (segment: LineSegment, offsetPt: number): LineSegment =>
  segment.kind === "text" ? { ...segment, offsetPt } : { ...segment, offsetPt };

// A word with no line of its own to fit on is cut at the character that overflows,
// and never before the first one, so a narrow column still makes progress.
function splitFragments(
  fragments: readonly Fragment[],
  availablePt: number,
): readonly [readonly Fragment[], readonly Fragment[]] {
  const head: Fragment[] = [];
  const tail: Fragment[] = [];
  let taken = 0;
  let filled = 0;

  for (const fragment of fragments) {
    if (taken > 0 && filled + fragment.widthPt > availablePt + EPSILON) {
      tail.push(fragment);
      continue;
    }
    if (tail.length > 0) {
      tail.push(fragment);
      continue;
    }
    if (filled + fragment.widthPt <= availablePt + EPSILON) {
      head.push(fragment);
      filled += fragment.widthPt;
      taken += charactersOf(fragment.text).length;
      continue;
    }

    const [left, right] = splitFragment(fragment, availablePt - filled, taken === 0);
    if (left !== null) {
      head.push(left);
      filled += left.widthPt;
      taken += charactersOf(left.text).length;
    }
    if (right !== null) tail.push(right);
  }

  return [head, tail];
}

function splitFragment(
  fragment: Fragment,
  availablePt: number,
  atLeastOne: boolean,
): readonly [Fragment | null, Fragment | null] {
  const characters = charactersOf(fragment.text);
  const perCharacter = fragment.widthPt / characters.length;

  let count = 0;
  let filled = 0;
  while (count < characters.length && filled + perCharacter <= availablePt + EPSILON) {
    filled += perCharacter;
    count += 1;
  }
  if (count === 0 && atLeastOne) {
    count = 1;
    filled = perCharacter;
  }

  const head = characters.slice(0, count).join("");
  const tail = characters.slice(count).join("");
  return [
    head === "" ? null : { ...fragment, text: head, widthPt: filled },
    tail === "" ? null : { ...fragment, text: tail, widthPt: fragment.widthPt - filled },
  ];
}

// Word justifies a line by handing every space character on it an equal share of
// the room the line did not fill, whatever size that space is set in. A tab takes
// none, since it holds the stop it reached; a no-break space takes none either,
// being part of the word around it rather than a gap between words; and a space
// the line ended on has already hung past the edge and is not on the line at all.
export function justifyLine(line: TextLine, roomPt: number): TextLine {
  const slackPt = roomPt - line.widthPt;
  const spaces = line.segments.reduce((count, segment) => count + spaceCountOf(segment), 0);
  if (slackPt <= EPSILON || spaces === 0) return line;

  const sharePt = slackPt / spaces;
  let shiftPt = 0;

  const segments = line.segments.map((segment) => {
    const moved = { ...segment, offsetPt: segment.offsetPt + shiftPt };
    const grownPt = spaceCountOf(segment) * sharePt;
    if (grownPt === 0) return moved;

    shiftPt += grownPt;
    return { ...moved, widthPt: segment.widthPt + grownPt };
  });

  return { ...line, segments, widthPt: roomPt };
}

const spaceCountOf = (segment: LineSegment): number =>
  segment.kind === "text" && IS_GAP.test(segment.text) ? charactersOf(segment.text).length : 0;

export type TextMeasurement =
  | {
      readonly kind: "measured";
      readonly widthPt: number;
      readonly heightPt: number;
      readonly ascentPt: number;
    }
  | { readonly kind: "unmeasurable"; readonly failure: MeasureFailure };

// A run of text that never breaks, which is what a list number is.
export function measureText(
  text: string,
  mark: ParagraphMark,
  metricsFor: MetricsResolver,
): TextMeasurement {
  const measurer = new Measurer(metricsFor);
  const fragment = measurer.fragment(mark, text);
  if (fragment === null) {
    return { kind: "unmeasurable", failure: measurer.failure ?? { kind: "unresolved-font" } };
  }
  return {
    kind: "measured",
    widthPt: fragment.widthPt,
    heightPt: fragment.heightPt,
    ascentPt: fragment.ascentPt,
  };
}

export function breakLines(input: BreakLinesInput): LineBreaking {
  const measurer = new Measurer(input.metricsFor);
  const tokens = tokenize(input.runs, measurer);
  if (tokens.kind === "failed") {
    return { kind: "unmeasurable", failure: measurer.failure ?? { kind: "unresolved-font" } };
  }

  const breaker = new Breaker(
    input.widthPt,
    input.firstLineWidthPt ?? input.widthPt,
    input.tabs ?? NO_TABS,
  );
  for (const unit of tokens.value) {
    switch (unit.kind) {
      case "word":
        breaker.word(unit.fragments);
        break;
      case "space":
        breaker.space(unit.fragments);
        break;
      case "tab":
        breaker.tab();
        break;
      case "drawing":
        breaker.drawing(unit.widthPt, unit.heightPt);
        break;
      case "break":
        breaker.flush();
        break;
    }
  }

  return { kind: "lines", lines: breaker.finish() };
}
