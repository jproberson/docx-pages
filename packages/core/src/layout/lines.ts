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

export type LineSegment =
  | {
      readonly kind: "text";
      readonly mark: ParagraphMark;
      readonly text: string;
      readonly widthPt: number;
    }
  | { readonly kind: "drawing"; readonly widthPt: number; readonly heightPt: number };

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
// how far its own start sits from that edge for a tab to land on the right one.
export type LineTabs = {
  readonly stopsPt: readonly number[];
  readonly originPt: number;
};

export type BreakLinesInput = {
  readonly runs: readonly TextRun[];
  readonly widthPt: number;
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

const IS_SPACE = /^\s+$/;

function tokenize(runs: readonly TextRun[], measurer: Measurer): Measured<readonly Unit[]> {
  const units: Unit[] = [];

  const append = (kind: "word" | "space", fragment: Fragment): void => {
    const last = units.at(-1);
    if (last !== undefined && last.kind === kind) {
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

  for (const token of piece.text.split(/(\s+)/).filter((each) => each !== "")) {
    const fragment = measurer.fragment(mark, token);
    if (fragment === null) return false;
    append(IS_SPACE.test(token) ? "space" : "word", fragment);
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

  constructor(
    private readonly widthPt: number,
    private readonly tabs: LineTabs,
  ) {}

  private raise(heightPt: number, ascentPt: number): void {
    this.heightPt = Math.max(this.heightPt, heightPt);
    this.ascentPt = Math.max(this.ascentPt, ascentPt);
  }

  private commit(segments: readonly LineSegment[], widthPt: number): void {
    this.segments.push(...this.pending, ...segments);
    this.committedPt += this.pendingPt + widthPt;
    this.pending = [];
    this.pendingPt = 0;
    this.wrapped = false;
  }

  private get filled(): number {
    return this.committedPt + this.pendingPt;
  }

  private get empty(): boolean {
    return this.segments.length === 0;
  }

  flush(): void {
    if (!this.empty) {
      this.lines.push({
        segments: this.segments,
        widthPt: this.committedPt,
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
  }

  space(fragments: readonly Fragment[]): void {
    if (this.empty && this.wrapped) return;
    for (const fragment of fragments) {
      this.pending.push(segmentOf(fragment));
      this.pendingPt += fragment.widthPt;
      this.raise(fragment.heightPt, fragment.ascentPt);
    }
  }

  tab(): void {
    if (this.empty && this.wrapped) return;
    const { originPt, stopsPt } = this.tabs;
    this.pendingPt = nextTabStopPt(originPt + this.filled, stopsPt) - originPt - this.committedPt;
  }

  drawing(widthPt: number, heightPt: number): void {
    if (!this.empty && this.filled + widthPt > this.widthPt + EPSILON) this.wrap();
    this.raise(heightPt, heightPt);
    this.commit([{ kind: "drawing", widthPt, heightPt }], widthPt);
  }

  word(fragments: readonly Fragment[]): void {
    let rest = fragments;
    while (rest.length > 0) {
      const widthPt = widthOf(rest);
      if (this.filled + widthPt <= this.widthPt + EPSILON) {
        this.take(rest);
        return;
      }
      if (!this.empty) {
        this.wrap();
        continue;
      }
      const [head, tail] = splitFragments(rest, this.widthPt - this.filled);
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
});

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

  const breaker = new Breaker(input.widthPt, input.tabs ?? NO_TABS);
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
