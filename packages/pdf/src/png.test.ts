import { zlibSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  hasAlpha,
  readPng,
  samplesOf,
  splitAlpha,
  type PngColourType,
  type PngImage,
} from "./png.js";

// Pngs built here rather than shipped as files, so that a test can state exactly
// which filter each row was written with. That is the half of this worth pinning:
// the five filters are the only place a picture can come back subtly wrong rather
// than not at all.

const CRC_TABLE = Uint32Array.from({ length: 256 }, (_, byte) => {
  let value = byte;
  for (let bit = 0; bit < 8; bit += 1) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});

function crc32(bytes: Uint8Array): number {
  let value = 0xffffffff;
  for (const byte of bytes) value = (CRC_TABLE[(value ^ byte) & 0xff] ?? 0) ^ (value >>> 8);
  return (value ^ 0xffffffff) >>> 0;
}

function chunk(name: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(data.byteLength + 12);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.byteLength);
  for (const [at, character] of Array.from(name).entries()) out[4 + at] = character.charCodeAt(0);
  out.set(data, 8);
  view.setUint32(out.byteLength - 4, crc32(out.subarray(4, out.byteLength - 4)));
  return out;
}

const SIGNATURE = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

type PngFixture = {
  readonly width: number;
  readonly height: number;
  readonly colourType: PngColourType;
  readonly bitDepth?: number;
  readonly interlaced?: boolean;
  readonly palette?: Uint8Array;
  readonly transparency?: Uint8Array;
  // The rows exactly as the file carries them: a filter number then that row's
  // bytes, already written as whatever difference the filter means.
  readonly rows: readonly (readonly number[])[];
  // How many IDAT chunks to split the pixels across, since a real png splits them
  // over as many as it likes.
  readonly chunks?: number;
};

function buildPng(fixture: PngFixture): Uint8Array {
  const header = new Uint8Array(13);
  const view = new DataView(header.buffer);
  view.setUint32(0, fixture.width);
  view.setUint32(4, fixture.height);
  header[8] = fixture.bitDepth ?? 8;
  header[9] = fixture.colourType;
  header[12] = (fixture.interlaced ?? false) ? 1 : 0;

  const pixels = zlibSync(Uint8Array.from(fixture.rows.flat()));
  const across = fixture.chunks ?? 1;
  const per = Math.ceil(pixels.byteLength / across);
  const idat: Uint8Array[] = [];
  for (let at = 0; at < pixels.byteLength; at += per) {
    idat.push(chunk("IDAT", pixels.subarray(at, at + per)));
  }

  const parts = [
    SIGNATURE,
    chunk("IHDR", header),
    ...(fixture.palette === undefined ? [] : [chunk("PLTE", fixture.palette)]),
    ...(fixture.transparency === undefined ? [] : [chunk("tRNS", fixture.transparency)]),
    ...idat,
    chunk("IEND", new Uint8Array(0)),
  ];

  const total = parts.reduce((sum, part) => sum + part.byteLength, 0);
  const out = new Uint8Array(total);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.byteLength;
  }
  return out;
}

// Two rows of two pixels, red then green over blue then white, so that every
// channel differs and a row mixed up with another shows.
const PIXELS: readonly (readonly number[])[] = [
  [255, 0, 0, 255, 0, 255, 0, 128],
  [0, 0, 255, 64, 255, 255, 255, 255],
];

const NONE = 0;
const SUB = 1;
const UP = 2;
const AVERAGE = 3;
const PAETH = 4;

const unfiltered = (rows: readonly (readonly number[])[]): readonly (readonly number[])[] =>
  rows.map((row) => [NONE, ...row]);

// What a neighbour is worth to each filter, which is the definition the reader
// puts back by. Stated once here so the two sides cannot drift apart in wording
// while meaning different things.
function guess(filter: number, left: number, up: number, upLeft: number): number {
  if (filter === SUB) return left;
  if (filter === UP) return up;
  if (filter === AVERAGE) return (left + up) >> 1;
  if (filter !== PAETH) return 0;

  const estimate = left + up - upLeft;
  const toLeft = Math.abs(estimate - left);
  const toUp = Math.abs(estimate - up);
  const toUpLeft = Math.abs(estimate - upLeft);
  if (toLeft <= toUp && toLeft <= toUpLeft) return left;
  return toUp <= toUpLeft ? up : upLeft;
}

// `PIXELS` written out under one filter, which is what a png actually holds.
const BYTES_PER_PIXEL = 4;

function written(filter: number): readonly (readonly number[])[] {
  return PIXELS.map((row, at) => {
    const above = PIXELS[at - 1] ?? [];
    return [
      filter,
      ...row.map((value, index) => {
        const left = index >= BYTES_PER_PIXEL ? (row[index - BYTES_PER_PIXEL] ?? 0) : 0;
        const up = above[index] ?? 0;
        const upLeft = index >= BYTES_PER_PIXEL ? (above[index - BYTES_PER_PIXEL] ?? 0) : 0;
        return (value - guess(filter, left, up, upLeft)) & 0xff;
      }),
    ];
  });
}

const RGBA: PngFixture = { width: 2, height: 2, colourType: 6, rows: unfiltered(PIXELS) };

// Every kind of picture the format has, named rather than cast, so the list is
// checked against the type instead of being asserted past it.
const COLOUR_TYPES: readonly PngColourType[] = [0, 2, 3, 4, 6];

const read = (fixture: PngFixture): PngImage => {
  const png = readPng(buildPng(fixture));
  if (png === null) throw new Error("the fixture is not a png");
  return png;
};

describe("reading a png", () => {
  it("reads what the header states", () => {
    const png = readPng(buildPng(RGBA));

    expect(png).toMatchObject({
      widthPixels: 2,
      heightPixels: 2,
      bitDepth: 8,
      colourType: 6,
      interlaced: false,
    });
  });

  // A real png splits its pixels over as many chunks as it likes, and most of the
  // larger ones do. Taking the first alone reads a few files and truncates the
  // rest, which shows as a picture that fades to nothing part way down.
  it("joins every IDAT chunk, however many the file was split into", () => {
    const one = read(RGBA);
    const many = read({ ...RGBA, chunks: 4 });

    expect(many.deflated.byteLength).toBe(one.deflated.byteLength);
    expect(splitAlpha(many)).toStrictEqual(splitAlpha(one));
  });

  it("reads the palette and the transparency an indexed png states", () => {
    const png = readPng(
      buildPng({
        width: 2,
        height: 1,
        colourType: 3,
        palette: Uint8Array.from([255, 0, 0, 0, 255, 0]),
        transparency: Uint8Array.from([0, 255]),
        rows: [[NONE, 0, 1]],
      }),
    );

    expect([...(png?.palette ?? [])]).toStrictEqual([255, 0, 0, 0, 255, 0]);
    expect([...(png?.transparency ?? [])]).toStrictEqual([0, 255]);
  });

  it("says nothing at all for bytes that are not a png", () => {
    expect(readPng(Uint8Array.from([1, 2, 3, 4, 5, 6, 7, 8, 9]))).toBeNull();
    expect(readPng(new Uint8Array(0))).toBeNull();
  });

  it("reads an interlaced png as interlaced rather than refusing it", () => {
    expect(readPng(buildPng({ ...RGBA, interlaced: true }))?.interlaced).toBe(true);
  });
});

describe("which pictures carry alpha", () => {
  it("knows the two that do and the three that do not", () => {
    expect(COLOUR_TYPES.map(hasAlpha)).toStrictEqual([false, false, false, true, true]);
  });

  it("counts the samples a pixel of each kind carries", () => {
    expect(COLOUR_TYPES.map(samplesOf)).toStrictEqual([1, 3, 1, 2, 4]);
  });
});

describe("pulling the alpha out of a png", () => {
  const split = (fixture: PngFixture) => {
    const png = readPng(buildPng(fixture));
    if (png === null) throw new Error("the fixture is not a png");
    const out = splitAlpha(png);
    if (out === null) throw new Error("the pixels could not be split");
    return { colour: [...out.colour], alpha: [...out.alpha] };
  };

  it("keeps the colour in one run and what shows through in another", () => {
    expect(split(RGBA)).toStrictEqual({
      colour: [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255],
      alpha: [255, 128, 64, 255],
    });
  });

  // Each row of a png is written as a difference from its neighbours, and which
  // difference is the row's own to choose. All five have to be put back, since a
  // writer picks per row and a real file uses several.
  it("puts back a row written as a difference from the pixel to its left", () => {
    const rows = [
      [SUB, 255, 0, 0, 255, 1, 255, 0, 129],
      [NONE, 0, 0, 255, 64, 255, 255, 255, 255],
    ];
    expect(split({ ...RGBA, rows })).toStrictEqual(split(RGBA));
  });

  it("puts back a row written as a difference from the row above it", () => {
    const rows = [
      [NONE, 255, 0, 0, 255, 0, 255, 0, 128],
      [UP, 1, 0, 255, 65, 255, 0, 255, 127],
    ];
    expect(split({ ...RGBA, rows })).toStrictEqual(split(RGBA));
  });

  // The last two are written by the filter rather than by hand, because working
  // out an average or a Paeth guess by hand is how the first draft of this test
  // got them wrong. The two above are the check on that: they are written out
  // literally, and if the filters below were read the wrong way round those two
  // would have to be wrong as well.
  it("puts back a row written as a difference from the average of the two", () => {
    expect(split({ ...RGBA, rows: written(AVERAGE) })).toStrictEqual(split(RGBA));
  });

  it("puts back a row written as a difference from whichever neighbour is nearest", () => {
    expect(split({ ...RGBA, rows: written(PAETH) })).toStrictEqual(split(RGBA));
  });

  it("puts back a picture whose rows each chose a different filter", () => {
    const mixed = [written(SUB)[0], written(PAETH)[1]];
    expect(split({ ...RGBA, rows: mixed.map((row) => [...(row ?? [])]) })).toStrictEqual(
      split(RGBA),
    );
  });

  it("splits a greyscale picture that carries alpha as readily as a colour one", () => {
    const grey = read({ width: 2, height: 1, colourType: 4, rows: [[NONE, 10, 255, 20, 0]] });
    expect(splitAlpha(grey)).toStrictEqual({
      colour: Uint8Array.from([10, 20]),
      alpha: Uint8Array.from([255, 0]),
    });
  });

  // Bytes that stop early would otherwise be read past the end and come back as a
  // picture with a torn edge.
  it("answers nothing where the pixels do not add up to the size stated", () => {
    expect(splitAlpha(read({ ...RGBA, height: 4 }))).toBeNull();
  });
});
