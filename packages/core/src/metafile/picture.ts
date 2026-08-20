import type { FaceRequest, FontMetrics } from "../layout/font-metrics.js";
import type { MetricsResolver } from "../layout/lines.js";
import { EMR, readHeader, readRecords, type MetafileRecord } from "./records.js";

export type MetafileRect = {
  readonly leftUnits: number;
  readonly topUnits: number;
  readonly widthUnits: number;
  readonly heightUnits: number;
};

// Everything a played metafile turns out to draw, in the device units the picture
// is measured in. A shape carries the rectangle it was cut to rather than having
// been cut already, so text that overruns its clip is cut where Word cuts it
// instead of being dropped or drawn whole.
export type MetafileShape =
  | {
      readonly kind: "fill";
      readonly rect: MetafileRect;
      readonly color: string;
      readonly clipTo: MetafileRect | null;
    }
  | {
      readonly kind: "stroke";
      readonly fromXUnits: number;
      readonly fromYUnits: number;
      readonly toXUnits: number;
      readonly toYUnits: number;
      readonly widthUnits: number;
      readonly color: string;
      readonly clipTo: MetafileRect | null;
    }
  | {
      readonly kind: "text";
      readonly text: string;
      // Where each character of the text starts. A metafile states its own
      // advances, so the drawing does not have to measure the face to space the
      // text; one entry alone leaves the face to space it.
      readonly xUnits: readonly number[];
      readonly baselineUnits: number;
      readonly emUnits: number;
      readonly color: string;
      readonly face: FaceRequest;
      readonly clipTo: MetafileRect | null;
    };

export type MetafilePicture = {
  readonly widthUnits: number;
  readonly heightUnits: number;
  readonly shapes: readonly MetafileShape[];
};

type Pen = { readonly color: string; readonly widthUnits: number } | null;
type Brush = { readonly color: string } | null;
type Font = { readonly face: FaceRequest; readonly emUnits: number; readonly ascentUnits: number };

// Selecting an object the player could not make sense of is not itself a refusal:
// a metafile that selects the system font and never writes with it still plays.
// Drawing with one is what refuses the picture.
type MetafileObject =
  | { readonly kind: "pen"; readonly pen: Pen }
  | { readonly kind: "brush"; readonly brush: Brush }
  | { readonly kind: "font"; readonly font: Font }
  // Selecting the palette decides nothing about how a shape is drawn.
  | { readonly kind: "palette" }
  | { readonly kind: "unknown" };

const UNKNOWN: MetafileObject = { kind: "unknown" };

type DeviceContext = {
  readonly clip: MetafileRect | null;
  readonly pen: Pen | "unknown";
  readonly brush: Brush | "unknown";
  readonly font: Font | "unknown" | null;
  readonly textColor: string;
  readonly textAlign: number;
  readonly xUnits: number;
  readonly yUnits: number;
};

const INITIAL: DeviceContext = {
  clip: null,
  pen: { color: "#000000", widthUnits: 1 },
  brush: { color: "#ffffff" },
  font: null,
  textColor: "#000000",
  textAlign: 0,
  xUnits: 0,
  yUnits: 0,
};

const channel = (value: number): string => (value & 0xff).toString(16).padStart(2, "0");

// A colour is one word holding red, green and blue in the order the bytes fall.
const colorOf = (value: number): string =>
  `#${channel(value)}${channel(value >> 8)}${channel(value >> 16)}`;

const right = (rect: MetafileRect): number => rect.leftUnits + rect.widthUnits;
const bottom = (rect: MetafileRect): number => rect.topUnits + rect.heightUnits;

// A rectangle given as its two corners, holding the columns and rows from its left
// and top up to but not including its right and bottom.
function cornersRect(record: MetafileRecord, at: number): MetafileRect {
  const leftUnits = record.int(at);
  const topUnits = record.int(at + 4);
  return {
    leftUnits,
    topUnits,
    widthUnits: record.int(at + 8) - leftUnits,
    heightUnits: record.int(at + 12) - topUnits,
  };
}

function intersect(one: MetafileRect | null, other: MetafileRect): MetafileRect {
  if (one === null) return other;
  const leftUnits = Math.max(one.leftUnits, other.leftUnits);
  const topUnits = Math.max(one.topUnits, other.topUnits);
  return {
    leftUnits,
    topUnits,
    widthUnits: Math.min(right(one), right(other)) - leftUnits,
    heightUnits: Math.min(bottom(one), bottom(other)) - topUnits,
  };
}

const STOCK = 0x8000_0000;

const STOCK_OBJECTS: ReadonlyMap<number, MetafileObject> = new Map<number, MetafileObject>([
  [0x00, { kind: "brush", brush: { color: "#ffffff" } }],
  [0x01, { kind: "brush", brush: { color: "#c0c0c0" } }],
  [0x02, { kind: "brush", brush: { color: "#808080" } }],
  [0x03, { kind: "brush", brush: { color: "#404040" } }],
  [0x04, { kind: "brush", brush: { color: "#000000" } }],
  [0x05, { kind: "brush", brush: null }],
  [0x06, { kind: "pen", pen: { color: "#ffffff", widthUnits: 1 } }],
  [0x07, { kind: "pen", pen: { color: "#000000", widthUnits: 1 } }],
  [0x08, { kind: "pen", pen: null }],
  [0x0f, { kind: "palette" }],
  // The stock fonts, whose metrics belong to whatever machine recorded the file.
  [0x0a, UNKNOWN],
  [0x0b, UNKNOWN],
  [0x0c, UNKNOWN],
  [0x0d, UNKNOWN],
  [0x0e, UNKNOWN],
  [0x10, UNKNOWN],
  [0x11, UNKNOWN],
]);

const BRUSH_SOLID = 0;
const BRUSH_NULL = 1;

const PEN_SOLID = 0;
const PEN_NULL = 5;

function readPen(record: MetafileRecord): MetafileObject {
  const style = record.uint(12);
  if (style === PEN_NULL) return { kind: "pen", pen: null };
  if (style !== PEN_SOLID) return UNKNOWN;
  // A pen stored with no width still draws a line one unit wide.
  const widthUnits = Math.max(record.int(16), 1);
  return { kind: "pen", pen: { color: colorOf(record.uint(24)), widthUnits } };
}

function readBrush(record: MetafileRecord): MetafileObject {
  const style = record.uint(12);
  if (style === BRUSH_NULL) return { kind: "brush", brush: null };
  if (style !== BRUSH_SOLID) return UNKNOWN;
  return { kind: "brush", brush: { color: colorOf(record.uint(16)) } };
}

const LOG_FONT_AT = 12;
const FACE_NAME_AT = LOG_FONT_AT + 28;
const FACE_NAME_CHARACTERS = 32;

const BOLD_WEIGHT = 700;

// The height stored positive is the whole cell the characters sit in, which is the
// face's own ascent and descent together.
const cellEm = (height: number, metrics: FontMetrics): number =>
  (height * metrics.unitsPerEm) / (metrics.ascender - metrics.descender);

// A face is asked for by the height of its characters when the height is stored
// negative and by the height of its cell when it is positive, so either one gives
// the em it is drawn at; the ascent under the cell's top follows from the same
// metrics. A face nothing can supply metrics for is refused rather than drawn in a
// substitute, since the metafile's own advances would then space the wrong glyphs.
function readFont(record: MetafileRecord, metricsFor: MetricsResolver): MetafileObject {
  if (record.int(LOG_FONT_AT + 8) !== 0 || record.int(LOG_FONT_AT + 12) !== 0) return UNKNOWN;

  const name = record.text(FACE_NAME_AT, FACE_NAME_CHARACTERS).replace(/\0[\s\S]*$/, "");
  if (name === "") return UNKNOWN;

  const face: FaceRequest = {
    name,
    bold: record.int(LOG_FONT_AT + 16) >= BOLD_WEIGHT,
    italic: record.uint(LOG_FONT_AT + 20) % 0x100 !== 0,
  };

  const found = metricsFor(face);
  if (found.kind !== "found") return UNKNOWN;

  const height = record.int(LOG_FONT_AT);
  const emUnits = height < 0 ? -height : cellEm(height, found.metrics);
  if (emUnits <= 0) return UNKNOWN;

  const ascentUnits = (emUnits * found.metrics.ascender) / found.metrics.unitsPerEm;
  return { kind: "font", font: { face, emUnits, ascentUnits } };
}

// The whole of a raster operation that takes no source: the destination is painted
// in the brush that is selected, which is how this file draws every block of colour
// and every rule between them.
const PATCOPY = 0x00f0_0021;

const REGION_COPY = 5;
const MAP_MODE_TEXT = 1;
const ROP2_COPY_PEN = 13;

// Where the reference point of a run of text sits against the text itself. Only the
// vertical half is answered: a run aligned by its middle or its right edge would
// have to be measured to place, and is refused instead.
const ALIGN_HORIZONTAL = 0x0006;
const ALIGN_UPDATE_POSITION = 0x0001;
const ALIGN_BASELINE = 0x0018;
const ALIGN_BOTTOM = 0x0008;

// The text is glyph indexes rather than characters, which only the recording
// machine's own font could turn back into text.
const TEXT_GLYPH_INDEX = 0x0010;
const TEXT_CLIPPED = 0x0004;
const TEXT_OPAQUE = 0x0002;
// Each advance is a pair, of which only the first moves the text along.
const TEXT_PAIRED_ADVANCES = 0x2000;

const EMR_TEXT_AT = 36;

// The reference point names the top of the character cell unless the alignment says
// otherwise, so the baseline sits the face's own ascent below it. Measured against
// Word's own pdf: four runs recorded at an em of 22 units, in a face whose ascender
// is 1950 of 2048, landed on baselines 20.98 units below their reference points,
// and this puts them at 20.95.
function baselineOffset(align: number, font: Font): number {
  if ((align & ALIGN_BASELINE) === ALIGN_BASELINE) return 0;
  if ((align & ALIGN_BOTTOM) !== 0) return font.ascentUnits - font.emUnits;
  return font.ascentUnits;
}

const placesText = (align: number): boolean =>
  (align & ALIGN_HORIZONTAL) === 0 && (align & ALIGN_UPDATE_POSITION) === 0;

function readTextOut(
  record: MetafileRecord,
  dc: DeviceContext,
  font: Font,
): MetafileShape | undefined {
  const options = record.uint(EMR_TEXT_AT + 16);
  if ((options & (TEXT_GLYPH_INDEX | TEXT_OPAQUE)) !== 0) return undefined;

  const characters = record.uint(EMR_TEXT_AT + 8);
  const text = record.text(record.uint(EMR_TEXT_AT + 12), characters);
  if (text.length !== characters) return undefined;

  const leftUnits = record.int(EMR_TEXT_AT);
  const offDx = record.uint(EMR_TEXT_AT + 36);
  const stride = (options & TEXT_PAIRED_ADVANCES) === 0 ? 4 : 8;
  // A record answers nought for anything read past its own end, so an array stated
  // where the record does not reach would space every character at nought and stack
  // the whole run on its first. The array the format asks for is one advance a
  // character, which is the bound held here: a record short of that is refused
  // rather than drawn wrong.
  if (offDx !== 0 && offDx + characters * stride > record.length) return undefined;

  const xUnits = [leftUnits];
  if (offDx !== 0) {
    for (let at = 0; at + 1 < characters; at += 1) {
      xUnits.push((xUnits[at] ?? leftUnits) + record.int(offDx + at * stride));
    }
  }

  return {
    kind: "text",
    text,
    xUnits,
    baselineUnits: record.int(EMR_TEXT_AT + 4) + baselineOffset(dc.textAlign, font),
    emUnits: font.emUnits,
    color: dc.textColor,
    face: font.face,
    clipTo:
      (options & TEXT_CLIPPED) === 0
        ? dc.clip
        : intersect(dc.clip, cornersRect(record, EMR_TEXT_AT + 20)),
  };
}

// What selecting a handle changes about the context, or undefined for a handle the
// player cannot account for at all.
function selecting(
  handle: number,
  objects: ReadonlyMap<number, MetafileObject>,
): Partial<DeviceContext> | undefined {
  const object = (handle & STOCK) === 0 ? objects.get(handle) : STOCK_OBJECTS.get(handle & ~STOCK);
  if (object === undefined) return undefined;

  switch (object.kind) {
    case "pen":
      return { pen: object.pen };
    case "brush":
      return { brush: object.brush };
    case "font":
      return { font: object.font };
    case "palette":
      return {};
    case "unknown":
      return { pen: "unknown", brush: "unknown", font: "unknown" };
  }
}

/**
 * Plays an enhanced metafile into the shapes it draws, or answers null when it
 * holds anything this cannot draw faithfully. A picture nobody can draw is marked
 * where it sits rather than half drawn, so refusing is the honest answer to a
 * record whose meaning is not known; only records that certainly put no ink on the
 * page are passed over.
 */
export function readMetafilePicture(
  bytes: Uint8Array,
  metricsFor: MetricsResolver,
): MetafilePicture | null {
  const records = readRecords(bytes);
  const first = records?.[0];
  if (records === null || first === undefined) return null;

  const header = readHeader(first);
  if (header === null) return null;

  const shapes: MetafileShape[] = [];
  const objects = new Map<number, MetafileObject>();
  const saved: DeviceContext[] = [];
  let dc = INITIAL;

  for (const record of records.slice(1)) {
    switch (record.type) {
      case EMR.eof:
      case EMR.gdiComment:
      case EMR.setMapperFlags:
      case EMR.setPolyFillMode:
      case EMR.setStretchBltMode:
      case EMR.setArcDirection:
      case EMR.setMiterLimit:
      case EMR.setIcmMode:
      case EMR.setLayout:
      case EMR.setBkMode:
      case EMR.setBkColor:
        break;

      // Every coordinate in the file is already the device's own, and stays so for
      // as long as nothing shifts the mapping under it.
      case EMR.setMapMode:
        if (record.uint(8) !== MAP_MODE_TEXT) return null;
        break;
      case EMR.setWindowOrgEx:
        if (record.int(8) !== 0 || record.int(12) !== 0) return null;
        break;
      case EMR.setRop2:
        if (record.uint(8) !== ROP2_COPY_PEN) return null;
        break;

      case EMR.saveDc:
        saved.push(dc);
        break;
      case EMR.restoreDc: {
        // The count says how many of the saved contexts to come back past.
        const steps = Math.abs(record.int(8));
        const restored = steps === 0 ? undefined : saved.splice(saved.length - steps)[0];
        if (restored === undefined) return null;
        dc = restored;
        break;
      }

      case EMR.setTextColor:
        dc = { ...dc, textColor: colorOf(record.uint(8)) };
        break;
      case EMR.setTextAlign:
        dc = { ...dc, textAlign: record.uint(8) };
        break;

      case EMR.intersectClipRect:
        dc = { ...dc, clip: intersect(dc.clip, cornersRect(record, 8)) };
        break;
      // The only region this plays is the empty one that puts the clip back to the
      // whole picture; any other is a shape of its own to work out.
      case EMR.extSelectClipRgn:
        if (record.uint(8) !== 0 || record.uint(12) !== REGION_COPY) return null;
        dc = { ...dc, clip: null };
        break;

      case EMR.createPen:
        objects.set(record.uint(8), readPen(record));
        break;
      case EMR.createBrushIndirect:
        objects.set(record.uint(8), readBrush(record));
        break;
      case EMR.extCreateFontIndirectW:
        objects.set(record.uint(8), readFont(record, metricsFor));
        break;
      case EMR.deleteObject:
        objects.delete(record.uint(8));
        break;
      case EMR.selectObject: {
        const change = selecting(record.uint(8), objects);
        if (change === undefined) return null;
        dc = { ...dc, ...change };
        break;
      }

      case EMR.moveToEx:
        dc = { ...dc, xUnits: record.int(8), yUnits: record.int(12) };
        break;
      case EMR.lineTo: {
        if (dc.pen === "unknown") return null;
        const toXUnits = record.int(8);
        const toYUnits = record.int(12);
        if (dc.pen !== null) {
          shapes.push({
            kind: "stroke",
            fromXUnits: dc.xUnits,
            fromYUnits: dc.yUnits,
            toXUnits,
            toYUnits,
            widthUnits: dc.pen.widthUnits,
            color: dc.pen.color,
            clipTo: dc.clip,
          });
        }
        dc = { ...dc, xUnits: toXUnits, yUnits: toYUnits };
        break;
      }

      case EMR.bitBlt: {
        if (record.uint(40) !== PATCOPY) return null;
        // A source of any kind makes this a bitmap rather than a block of the
        // brush, and a bitmap is not something this can draw.
        if (record.uint(88) !== 0 || record.uint(96) !== 0) return null;
        if (dc.brush === "unknown") return null;
        const rect = {
          leftUnits: record.int(24),
          topUnits: record.int(28),
          widthUnits: record.int(32),
          heightUnits: record.int(36),
        };
        if (dc.brush !== null && rect.widthUnits > 0 && rect.heightUnits > 0) {
          shapes.push({ kind: "fill", rect, color: dc.brush.color, clipTo: dc.clip });
        }
        break;
      }

      case EMR.extTextOutW: {
        const { font } = dc;
        if (font === null || font === "unknown" || !placesText(dc.textAlign)) return null;
        const shape = readTextOut(record, dc, font);
        if (shape === undefined) return null;
        shapes.push(shape);
        break;
      }

      default:
        return null;
    }
  }

  // A file whose drawing is all in records this passed over has not been played at
  // all, whatever it looked like on the way through.
  return shapes.length === 0 ? null : { ...header, shapes };
}
