// Enhanced metafiles for tests to play. Every record is a type, its own length and
// then a run of words, so a fixture states the words a record holds and the builder
// counts the bytes.

const HEADER_BYTES = 8;
const WORD_BYTES = 4;

const EMPTY = new Uint8Array(0);

export function metafileRecord(
  type: number,
  words: readonly number[],
  tail: Uint8Array = EMPTY,
): Uint8Array {
  const padding = (WORD_BYTES - (tail.byteLength % WORD_BYTES)) % WORD_BYTES;
  const length = HEADER_BYTES + words.length * WORD_BYTES + tail.byteLength + padding;

  const bytes = new Uint8Array(length);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, type, true);
  view.setUint32(4, length, true);
  words.forEach((word, at) => {
    view.setInt32(HEADER_BYTES + at * WORD_BYTES, word, true);
  });
  bytes.set(tail, HEADER_BYTES + words.length * WORD_BYTES);
  return bytes;
}

export function utf16(text: string): Uint8Array {
  const bytes = new Uint8Array(text.length * 2);
  const view = new DataView(bytes.buffer);
  for (let at = 0; at < text.length; at += 1) {
    view.setUint16(at * 2, text.charCodeAt(at), true);
  }
  return bytes;
}

export type MetafileFrame = {
  // The picture's own size, in hundredths of a millimetre.
  readonly frameWidth: number;
  readonly frameHeight: number;
  // What the machine that recorded it measured, which is what turns the frame back
  // into the units the records are written in.
  readonly deviceWidth?: number;
  readonly deviceHeight?: number;
  readonly millimetreWidth?: number;
  readonly millimetreHeight?: number;
};

const SIGNATURE = 0x464d4520;

export function metafileHeader(frame: MetafileFrame): Uint8Array {
  const {
    frameWidth,
    frameHeight,
    deviceWidth = 1920,
    deviceHeight = 1080,
    millimetreWidth = 309,
    millimetreHeight = 174,
  } = frame;

  // prettier-ignore
  return metafileRecord(1, [
    0, 0, 0, 0,
    0, 0, frameWidth, frameHeight,
    SIGNATURE, 0x10000, 0, 0,
    0, 0, 0, 0,
    deviceWidth, deviceHeight, millimetreWidth, millimetreHeight,
  ]);
}

export type MetafileFontFixture = {
  readonly handle: number;
  readonly name: string;
  // Negative for the height of the characters alone, positive for the whole cell
  // they sit in.
  readonly heightUnits: number;
  readonly weight?: number;
  readonly italic?: boolean;
  readonly escapement?: number;
};

const FACE_NAME_BYTES = 64;

export function metafileFont(fixture: MetafileFontFixture): Uint8Array {
  const { handle, name, heightUnits, weight = 400, italic = false, escapement = 0 } = fixture;

  const face = new Uint8Array(FACE_NAME_BYTES);
  face.set(utf16(name).subarray(0, FACE_NAME_BYTES));

  // prettier-ignore
  return metafileRecord(82, [
    handle, heightUnits, 0, escapement, 0, weight, italic ? 1 : 0, 0,
  ], face);
}

export type MetafileTextFixture = {
  readonly xUnits: number;
  readonly yUnits: number;
  readonly text: string;
  // How far each character moves the text along. Left out, the drawing spaces the
  // string with the face's own advances.
  readonly advances?: readonly number[];
  readonly options?: number;
};

const TEXT_WORDS = 17;

// A run of text keeps its characters and its advances past the fields that say
// where each of them starts, so the offsets follow from how long the string is.
export function metafileText(fixture: MetafileTextFixture): Uint8Array {
  const { xUnits, yUnits, text, advances, options = 0 } = fixture;

  const stringAt = HEADER_BYTES + TEXT_WORDS * WORD_BYTES;
  const advancesAt = stringAt + Math.ceil((text.length * 2) / WORD_BYTES) * WORD_BYTES;

  const tail = new Uint8Array(
    advancesAt - stringAt + (advances === undefined ? 0 : advances.length * WORD_BYTES),
  );
  tail.set(utf16(text));
  if (advances !== undefined) {
    const view = new DataView(tail.buffer);
    advances.forEach((advance, at) => {
      view.setInt32(advancesAt - stringAt + at * WORD_BYTES, advance, true);
    });
  }

  // prettier-ignore
  return metafileRecord(84, [
    0, 0, 0, 0,
    2, 0, 0,
    xUnits, yUnits, text.length, stringAt, options,
    0, 0, 0, 0,
    advances === undefined ? 0 : advancesAt,
  ], tail);
}

export function buildMetafile(records: readonly Uint8Array[]): Uint8Array {
  const length = records.reduce((total, record) => total + record.byteLength, 0);
  const bytes = new Uint8Array(length);
  let at = 0;
  for (const record of records) {
    bytes.set(record, at);
    at += record.byteLength;
  }
  return bytes;
}
