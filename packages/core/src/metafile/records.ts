// An enhanced metafile is a flat run of records, each stating its own type and its
// own length. Nothing here reads a record's contents: every field a record holds
// is stated at an offset from that record's own start, so the reader hands each one
// out as a window onto the file and the player reads only the fields it knows.

export type MetafileRecord = {
  readonly type: number;
  readonly length: number;
  int(offset: number): number;
  uint(offset: number): number;
  // The format stores text as UTF-16, in a run of characters whose length is
  // counted separately from where it starts.
  text(offset: number, characters: number): string;
};

// The record types this reader names. Everything else a metafile can hold is met
// by number and, if it might draw, refused.
export const EMR = {
  header: 1,
  setWindowOrgEx: 10,
  eof: 14,
  setMapperFlags: 16,
  setMapMode: 17,
  setBkMode: 18,
  setPolyFillMode: 19,
  setRop2: 20,
  setStretchBltMode: 21,
  setTextAlign: 22,
  setTextColor: 24,
  setBkColor: 25,
  moveToEx: 27,
  intersectClipRect: 30,
  saveDc: 33,
  restoreDc: 34,
  selectObject: 37,
  createPen: 38,
  createBrushIndirect: 39,
  deleteObject: 40,
  lineTo: 54,
  setArcDirection: 57,
  setMiterLimit: 58,
  gdiComment: 70,
  extSelectClipRgn: 75,
  bitBlt: 76,
  extCreateFontIndirectW: 82,
  extTextOutW: 84,
  setIcmMode: 98,
  setLayout: 115,
} as const;

const RECORD_HEADER_BYTES = 8;

function recordAt(view: DataView, start: number, length: number): MetafileRecord {
  const within = (offset: number, bytes: number): boolean =>
    offset >= 0 && offset + bytes <= length;

  return {
    type: view.getUint32(start, true),
    length,
    int: (offset) => (within(offset, 4) ? view.getInt32(start + offset, true) : 0),
    uint: (offset) => (within(offset, 4) ? view.getUint32(start + offset, true) : 0),
    text: (offset, characters) => {
      if (!within(offset, characters * 2)) return "";
      let value = "";
      for (let at = 0; at < characters; at += 1) {
        value += String.fromCharCode(view.getUint16(start + offset + at * 2, true));
      }
      return value;
    },
  };
}

// A run that does not divide cleanly into records is not a metafile this can play,
// which is a fact about the bytes rather than a fault to raise.
export function readRecords(bytes: Uint8Array): readonly MetafileRecord[] | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const records: MetafileRecord[] = [];

  let at = 0;
  while (at + RECORD_HEADER_BYTES <= bytes.byteLength) {
    const length = view.getUint32(at + 4, true);
    if (length < RECORD_HEADER_BYTES || length % 4 !== 0) return null;
    if (at + length > bytes.byteLength) return null;
    records.push(recordAt(view, at, length));
    at += length;
  }

  return at === bytes.byteLength ? records : null;
}

export type MetafileHeader = {
  // The extent the whole picture is drawn into, in the device units every record
  // states its coordinates in.
  readonly widthUnits: number;
  readonly heightUnits: number;
};

// " EMF", little-endian.
const SIGNATURE = 0x464d4520;

const HEADER_MINIMUM_BYTES = 88;

// Word plays a metafile into the rectangle the document gave the picture, and what
// goes in that rectangle is the frame rather than the bounds around the ink: this
// file's bounds stop two units short of its own outermost marks.
//
// The frame is stored in hundredths of a millimetre and the header says what the
// recording device measured, so the frame comes back to device units through that
// device's own resolution. The rectangle it describes is inclusive of both of its
// edges, which is the unit the extent is one wider than the conversion.
//
// Measured against Word's own pdf for a 6196 x 4286 frame recorded at 1920 x 1080
// over 309 x 174mm: Word drew it 386 units wide and 267 tall.
export function readHeader(record: MetafileRecord): MetafileHeader | null {
  if (record.type !== EMR.header || record.length < HEADER_MINIMUM_BYTES) return null;
  if (record.uint(40) !== SIGNATURE) return null;

  const frameWidth = record.int(32) - record.int(24);
  const frameHeight = record.int(36) - record.int(28);
  const deviceWidth = record.int(72);
  const deviceHeight = record.int(76);
  const millimetreWidth = record.int(80);
  const millimetreHeight = record.int(84);
  if (millimetreWidth <= 0 || millimetreHeight <= 0) return null;

  const widthUnits = Math.round((frameWidth * deviceWidth) / (millimetreWidth * 100)) + 1;
  const heightUnits = Math.round((frameHeight * deviceHeight) / (millimetreHeight * 100)) + 1;
  if (widthUnits <= 0 || heightUnits <= 0) return null;

  return { widthUnits, heightUnits };
}
