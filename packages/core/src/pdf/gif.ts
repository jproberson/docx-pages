// A gif opened far enough to draw it, which a pdf cannot be handed as it stands.
//
// **A pdf has no gif filter.** A jpeg goes across as `DCTDecode` and a png's rows
// are already what `FlateDecode` with predictor 15 means, so both are written
// without ever being opened. A gif is neither: its LZW is not the LZW a pdf reads
// (a gif packs codes low bit first and resets on a clear code of its own), so the
// bytes cannot be passed through and the pixels have to be produced here.
//
// **Word does exactly this.** Asked on 2026-08-24 of Word's own pdf of a corpus
// document holding one gif beside two pngs: the gif comes out as 8 bit RGB in an
// icc space with a soft mask of its own, which is the same shape Word gives a png
// carrying alpha. So decoding to pixels is not a shortcut around Word, it is what
// Word did.
//
// The first frame alone is read. A gif holding several is an animation, and a
// document draws the first of them.

export type GifImage = {
  readonly widthPixels: number;
  readonly heightPixels: number;
  // Three bytes a pixel, left to right and top to bottom.
  readonly colour: Uint8Array;
  // One byte a pixel, nought where the page shows through, or null where the file
  // marks nothing transparent and the picture is drawn without a mask.
  readonly alpha: Uint8Array | null;
};

const HEADER = 6;
const TRAILER = 0x3b;
const EXTENSION = 0x21;
const GRAPHIC_CONTROL = 0xf9;
const IMAGE = 0x2c;

// The rows of an interlaced gif arrive in four passes, each starting at its own row
// and stepping by its own stride. Woven back together here rather than refused: a
// gif written for the web is often interlaced, and reading one straight draws the
// picture in stripes.
const PASSES: readonly (readonly [number, number])[] = [
  [0, 8],
  [4, 8],
  [2, 4],
  [1, 2],
];

const isGif = (bytes: Uint8Array): boolean =>
  bytes.length > HEADER &&
  bytes[0] === 0x47 &&
  bytes[1] === 0x49 &&
  bytes[2] === 0x46 &&
  bytes[3] === 0x38 &&
  (bytes[4] === 0x37 || bytes[4] === 0x39) &&
  bytes[5] === 0x61;

const at = (bytes: Uint8Array, index: number): number => bytes[index] ?? 0;

const wordAt = (bytes: Uint8Array, index: number): number =>
  at(bytes, index) | (at(bytes, index + 1) << 8);

// A run of length-prefixed blocks, ending at a length of nought. Gathered into one
// buffer before anything is decoded, so the decoder never straddles a block edge.
function gather(bytes: Uint8Array, from: number): { data: Uint8Array; next: number } | null {
  const parts: Uint8Array[] = [];
  let cursor = from;
  let total = 0;
  for (;;) {
    if (cursor >= bytes.length) return null;
    const size = at(bytes, cursor);
    cursor += 1;
    if (size === 0) break;
    if (cursor + size > bytes.length) return null;
    parts.push(bytes.subarray(cursor, cursor + size));
    total += size;
    cursor += size;
  }
  const data = new Uint8Array(total);
  let put = 0;
  for (const part of parts) {
    data.set(part, put);
    put += part.length;
  }
  return { data, next: cursor };
}

/**
 * The indices of one frame, out of the LZW a gif holds them in.
 *
 * The dictionary starts as the palette itself plus a clear code and an end code,
 * and grows by one entry a code until it is full, whereupon the file either sends
 * a clear and starts again or goes on using the dictionary as it stands. **A code
 * one past the end of the dictionary is the legal one that means "what I last
 * emitted, and its own first byte again"**, which is the case a decoder that only
 * looks entries up gets wrong.
 */
function decode(data: Uint8Array, minCodeSize: number, pixels: number): Uint8Array | null {
  if (minCodeSize < 2 || minCodeSize > 8) return null;

  const clear = 1 << minCodeSize;
  const end = clear + 1;
  const prefix = new Int32Array(4096);
  const suffix = new Uint8Array(4096);
  const first = new Uint8Array(4096);
  const out = new Uint8Array(pixels);

  let put = 0;
  let codeSize = minCodeSize + 1;
  let next = end + 1;
  let previous = -1;
  let bits = 0;
  let held = 0;
  let read = 0;
  const stack = new Uint8Array(4096);

  for (let i = 0; i < clear; i += 1) {
    prefix[i] = -1;
    suffix[i] = i;
    first[i] = i;
  }

  while (put < pixels) {
    while (bits < codeSize) {
      if (read >= data.length) return put === 0 ? null : out;
      held |= at(data, read) << bits;
      bits += 8;
      read += 1;
    }
    const code = held & ((1 << codeSize) - 1);
    held >>>= codeSize;
    bits -= codeSize;

    if (code === clear) {
      codeSize = minCodeSize + 1;
      next = end + 1;
      previous = -1;
      continue;
    }
    if (code === end) break;

    let current = code;
    let top = 0;
    if (code >= next) {
      if (previous === -1) return null;
      stack[top] = first[previous] ?? 0;
      top += 1;
      current = previous;
    }
    while (current >= clear) {
      stack[top] = suffix[current] ?? 0;
      top += 1;
      current = prefix[current] ?? -1;
      if (current < 0) return null;
    }
    stack[top] = suffix[current] ?? 0;
    top += 1;

    for (let i = top - 1; i >= 0 && put < pixels; i -= 1) {
      out[put] = stack[i] ?? 0;
      put += 1;
    }

    if (previous !== -1 && next < 4096) {
      prefix[next] = previous;
      suffix[next] = first[current] ?? 0;
      first[next] = first[previous] ?? 0;
      next += 1;
      if ((next & (next - 1)) === 0 && next < 4096 && codeSize < 12) codeSize += 1;
    }
    if (previous === -1) first[code] = first[current] ?? 0;
    previous = code;
  }

  return out;
}

/**
 * A gif's first frame, as pixels and what shows through them.
 *
 * Answers null for anything it cannot read whole rather than drawing part of a
 * picture, which is the same bargain the metafile player makes: a frame drawn from
 * half a file is a plausible-looking page, and this project would rather report a
 * picture it did not draw.
 */
export function readGif(bytes: Uint8Array): GifImage | null {
  if (!isGif(bytes)) return null;

  const screenWidth = wordAt(bytes, 6);
  const screenHeight = wordAt(bytes, 8);
  if (screenWidth <= 0 || screenHeight <= 0) return null;

  const packed = at(bytes, 10);
  let cursor = 13;
  let global: Uint8Array | null = null;
  if ((packed & 0x80) !== 0) {
    const entries = 1 << ((packed & 0x07) + 1);
    if (cursor + entries * 3 > bytes.length) return null;
    global = bytes.subarray(cursor, cursor + entries * 3);
    cursor += entries * 3;
  }

  let transparent = -1;

  for (;;) {
    if (cursor >= bytes.length) return null;
    const marker = at(bytes, cursor);
    cursor += 1;

    if (marker === TRAILER) return null;

    if (marker === EXTENSION) {
      const label = at(bytes, cursor);
      cursor += 1;
      if (label === GRAPHIC_CONTROL) {
        // Size, flags, delay, the index that shows through, terminator.
        const flags = at(bytes, cursor + 1);
        transparent = (flags & 0x01) === 0 ? -1 : at(bytes, cursor + 4);
      }
      const skipped = gather(bytes, cursor);
      if (skipped === null) return null;
      cursor = skipped.next;
      continue;
    }

    if (marker !== IMAGE) return null;

    const left = wordAt(bytes, cursor);
    const top = wordAt(bytes, cursor + 2);
    const width = wordAt(bytes, cursor + 4);
    const height = wordAt(bytes, cursor + 6);
    const frameFlags = at(bytes, cursor + 8);
    cursor += 9;

    let palette = global;
    if ((frameFlags & 0x80) !== 0) {
      const entries = 1 << ((frameFlags & 0x07) + 1);
      if (cursor + entries * 3 > bytes.length) return null;
      palette = bytes.subarray(cursor, cursor + entries * 3);
      cursor += entries * 3;
    }
    if (palette === null || width <= 0 || height <= 0) return null;

    const minCodeSize = at(bytes, cursor);
    cursor += 1;
    const gathered = gather(bytes, cursor);
    if (gathered === null) return null;

    const indices = decode(gathered.data, minCodeSize, width * height);
    if (indices === null) return null;

    const interlaced = (frameFlags & 0x40) !== 0;
    const pixels = screenWidth * screenHeight;
    const colour = new Uint8Array(pixels * 3);
    // Everything the frame does not cover shows the page through, which is what an
    // uncovered pixel of a gif standing smaller than its own screen means.
    const covers = left === 0 && top === 0 && width === screenWidth && height === screenHeight;
    const alpha = transparent >= 0 || !covers ? new Uint8Array(pixels) : null;
    if (alpha !== null && covers) alpha.fill(255);

    for (let row = 0; row < height; row += 1) {
      const source = interlaced ? interlacedRow(row, height) : row;
      if (source < 0) return null;
      for (let column = 0; column < width; column += 1) {
        const index = indices[source * width + column] ?? 0;
        const x = left + column;
        const y = top + row;
        if (x >= screenWidth || y >= screenHeight) continue;
        const pixel = y * screenWidth + x;
        const entry = index * 3;
        colour[pixel * 3] = palette[entry] ?? 0;
        colour[pixel * 3 + 1] = palette[entry + 1] ?? 0;
        colour[pixel * 3 + 2] = palette[entry + 2] ?? 0;
        if (alpha !== null) alpha[pixel] = index === transparent ? 0 : 255;
      }
    }

    return { widthPixels: screenWidth, heightPixels: screenHeight, colour, alpha };
  }
}

// Which row of the picture the nth row of an interlaced frame is.
function interlacedRow(row: number, height: number): number {
  let seen = 0;
  for (const pass of PASSES) {
    const [start, step] = pass;
    const rows = start >= height ? 0 : Math.ceil((height - start) / step);
    if (row < seen + rows) return start + (row - seen) * step;
    seen += rows;
  }
  return -1;
}
