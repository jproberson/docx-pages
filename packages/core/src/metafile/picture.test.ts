import { describe, expect, it } from "vitest";

import {
  buildMetafile,
  metafileFont,
  metafileHeader,
  metafileRecord,
  metafileText,
} from "../testing/build-metafile.js";
import { lookupFontMetrics } from "../layout/font-metrics.js";
import { EMR } from "./records.js";
import { readMetafilePicture } from "./picture.js";

const metricsFor = (request: Parameters<typeof lookupFontMetrics>[0]) => lookupFontMetrics(request);

// Half a page across at 1920 pixels over 309mm, which is the resolution the sample
// this was measured against was recorded at.
const FRAME = { frameWidth: 6196, frameHeight: 4286 };

const PATCOPY = 0x00f00021;
const SOLID = 0;

const brush = (handle: number, color: number): Uint8Array =>
  metafileRecord(EMR.createBrushIndirect, [handle, SOLID, color, 0]);

const pen = (handle: number, color: number, width: number): Uint8Array =>
  metafileRecord(EMR.createPen, [handle, SOLID, width, 0, color]);

const select = (handle: number): Uint8Array => metafileRecord(EMR.selectObject, [handle]);

const blt = (
  leftUnits: number,
  topUnits: number,
  widthUnits: number,
  heightUnits: number,
  operation = PATCOPY,
): Uint8Array =>
  // prettier-ignore
  metafileRecord(EMR.bitBlt, [
    0, 0, 0, 0,
    leftUnits, topUnits, widthUnits, heightUnits,
    operation,
    0, 0,
    0, 0, 0, 0, 0, 0,
    0, 0, 0, 0, 0, 0,
  ]);

// Calibri, whose metrics the builtin table answers for, asked for by the height of
// its characters alone.
const CALIBRI_ASCENDER = 1950;
const CALIBRI_UNITS_PER_EM = 2048;
const CALIBRI_AT_22 = metafileFont({ handle: 7, name: "Calibri", heightUnits: -22 });

const played = (records: readonly Uint8Array[]) =>
  readMetafilePicture(buildMetafile([metafileHeader(FRAME), ...records]), metricsFor);

describe("readMetafilePicture", () => {
  // Measured against Word's own pdf, which drew this frame 386 units wide and 267
  // tall: the frame comes back to the recording device's units and the rectangle it
  // describes takes in both of its edges.
  it("measures the picture by the frame it was recorded at, edges included", () => {
    const picture = played([brush(1, 0x0000ff), select(1), blt(0, 0, 10, 10)]);
    expect(picture).toMatchObject({ widthUnits: 386, heightUnits: 267 });
  });

  it("paints a source-less block in the brush that is selected", () => {
    const picture = played([brush(1, 0x00b0f0), select(1), blt(3, 5, 40, 20)]);
    expect(picture?.shapes).toEqual([
      {
        kind: "fill",
        rect: { leftUnits: 3, topUnits: 5, widthUnits: 40, heightUnits: 20 },
        // The colour is stored with its channels in the order the bytes fall.
        color: "#f0b000",
        clipTo: null,
      },
    ]);
  });

  it("leaves a block the brush would not cover unpainted", () => {
    const hollow = metafileRecord(EMR.createBrushIndirect, [1, 1, 0, 0]);
    expect(played([hollow, select(1), blt(0, 0, 10, 10)])).toBeNull();
  });

  it("cuts a shape to the rectangles the clip has been narrowed by", () => {
    const picture = played([
      metafileRecord(EMR.intersectClipRect, [0, 0, 100, 100]),
      metafileRecord(EMR.intersectClipRect, [20, 0, 200, 40]),
      brush(1, 0),
      select(1),
      blt(0, 0, 300, 300),
    ]);
    expect(picture?.shapes[0]?.clipTo).toEqual({
      leftUnits: 20,
      topUnits: 0,
      widthUnits: 80,
      heightUnits: 40,
    });
  });

  it("puts back the clip a saved context was narrowed under", () => {
    const picture = played([
      metafileRecord(EMR.saveDc, []),
      metafileRecord(EMR.intersectClipRect, [0, 0, 10, 10]),
      metafileRecord(EMR.restoreDc, [-1]),
      brush(1, 0),
      select(1),
      blt(0, 0, 300, 300),
    ]);
    expect(picture?.shapes[0]?.clipTo).toBeNull();
  });

  it("draws a line in the pen that is selected, from where the position was moved", () => {
    const picture = played([
      pen(1, 0xd4d4d4, 2),
      select(1),
      metafileRecord(EMR.moveToEx, [10, 20]),
      metafileRecord(EMR.lineTo, [90, 20]),
    ]);
    expect(picture?.shapes).toEqual([
      {
        kind: "stroke",
        fromXUnits: 10,
        fromYUnits: 20,
        toXUnits: 90,
        toYUnits: 20,
        widthUnits: 2,
        color: "#d4d4d4",
        clipTo: null,
      },
    ]);
  });

  // Measured against Word's own pdf: four runs recorded at an em of 22 units in a
  // face whose ascender is 1950 of 2048 landed 20.98 units below their reference
  // points, which the face's own ascent puts at 20.95.
  it("hangs a run of text from its reference point by the face's own ascent", () => {
    const picture = played([
      CALIBRI_AT_22,
      select(7),
      metafileText({ xUnits: 28, yUnits: 88, text: "ab", advances: [11, 5] }),
    ]);
    expect(picture?.shapes[0]).toMatchObject({
      kind: "text",
      text: "ab",
      emUnits: 22,
      baselineUnits: 88 + (22 * CALIBRI_ASCENDER) / CALIBRI_UNITS_PER_EM,
      face: { name: "Calibri", bold: false, italic: false },
    });
  });

  it("starts each character where the metafile's own advances put it", () => {
    const picture = played([
      CALIBRI_AT_22,
      select(7),
      metafileText({ xUnits: 5, yUnits: 0, text: "abc", advances: [11, 7, 9] }),
    ]);
    expect(picture?.shapes[0]).toMatchObject({ xUnits: [5, 16, 23] });
  });

  it("leaves the face to space a run that states no advances of its own", () => {
    const picture = played([
      CALIBRI_AT_22,
      select(7),
      metafileText({ xUnits: 5, yUnits: 0, text: "abc" }),
    ]);
    expect(picture?.shapes[0]).toMatchObject({ xUnits: [5] });
  });

  it("refuses a run written in a font whose metrics nothing can answer for", () => {
    const unknown = metafileFont({ handle: 7, name: "Nothing Supplies This", heightUnits: -22 });
    expect(
      played([unknown, select(7), metafileText({ xUnits: 0, yUnits: 0, text: "a" })]),
    ).toBeNull();
  });

  it("refuses a run written in a stock font, whose metrics belong to the recorder", () => {
    const systemFont = 0x8000000d;
    expect(
      played([select(systemFont), metafileText({ xUnits: 0, yUnits: 0, text: "a" })]),
    ).toBeNull();
  });

  it("refuses a block that copies a bitmap rather than the brush", () => {
    const blackness = 0x00000042;
    expect(played([brush(1, 0), select(1), blt(0, 0, 10, 10, blackness)])).toBeNull();
  });

  it("refuses a record whose meaning it does not know, rather than passing over ink", () => {
    const polygon16 = 86;
    expect(
      played([brush(1, 0), select(1), blt(0, 0, 10, 10), metafileRecord(polygon16, [0, 0])]),
    ).toBeNull();
  });

  it("refuses a file that drew nothing it could play", () => {
    expect(played([metafileRecord(EMR.eof, [0, 0, 0])])).toBeNull();
    expect(readMetafilePicture(new Uint8Array(0), metricsFor)).toBeNull();
  });

  it("refuses bytes that do not divide into records", () => {
    const bytes = buildMetafile([metafileHeader(FRAME)]).slice(0, 60);
    expect(readMetafilePicture(bytes, metricsFor)).toBeNull();
  });

  // A header without the signature every enhanced metafile opens with says only
  // that the bytes happen to start with a record of the right shape.
  it("refuses a file that is not an enhanced metafile at all", () => {
    const record = metafileRecord(
      EMR.header,
      Array.from({ length: 20 }, () => 0),
    );
    expect(readMetafilePicture(record, metricsFor)).toBeNull();
  });
});
