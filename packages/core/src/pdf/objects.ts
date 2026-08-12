import { strToU8, zlibSync } from "fflate";

import { DocxPagesError } from "../errors.js";

// The half of a pdf that is not the drawing: objects, the table that says where
// each one starts, and the trailer that says which object the document begins at.
// Nothing here knows what a page is.
//
// A pdf is bytes rather than text. Its syntax is latin1 and its streams are
// arbitrary, so everything below is assembled as bytes and only the syntax is
// written as characters.

export type PdfReference = {
  readonly kind: "reference";
  readonly number: number;
};

export type PdfEntries = Readonly<Record<string, PdfValue | undefined>>;

export type PdfValue =
  | { readonly kind: "null" }
  | { readonly kind: "boolean"; readonly value: boolean }
  | { readonly kind: "number"; readonly value: number }
  // Written `/Name`. The characters a name may not hold are escaped rather than
  // refused, since a face's own name reaches this and a document may call a face
  // anything at all.
  | { readonly kind: "name"; readonly name: string }
  // Written `(text)`, one byte to a character, which is what the syntax reads a
  // string as until a document says otherwise.
  | { readonly kind: "literal-string"; readonly text: string }
  // Written `<hex>`. What carries a string of anything but latin1, and what the
  // text operators take: a font under Identity-H is addressed by glyph, and a
  // glyph number is two bytes rather than a character.
  | { readonly kind: "hex-string"; readonly bytes: Uint8Array }
  | { readonly kind: "array"; readonly items: readonly PdfValue[] }
  | { readonly kind: "dictionary"; readonly entries: PdfEntries }
  | {
      readonly kind: "stream";
      readonly entries: PdfEntries;
      readonly bytes: Uint8Array;
      // Whether the bytes are deflated on the way out. A stream that is already
      // compressed, a jpeg above all, states its own filter and is passed through
      // untouched.
      readonly deflate: boolean;
    }
  | PdfReference;

const AT = "pdf/objects.writePdfObjects";

export const pdfNull: PdfValue = { kind: "null" };

export const pdfBoolean = (value: boolean): PdfValue => ({ kind: "boolean", value });

export const pdfNumber = (value: number): PdfValue => ({ kind: "number", value });

export const pdfName = (name: string): PdfValue => ({ kind: "name", name });

export const pdfString = (text: string): PdfValue => ({ kind: "literal-string", text });

export const pdfHexString = (bytes: Uint8Array): PdfValue => ({ kind: "hex-string", bytes });

export const pdfArray = (items: readonly PdfValue[]): PdfValue => ({ kind: "array", items });

export const pdfDictionary = (entries: PdfEntries): PdfValue => ({ kind: "dictionary", entries });

export const pdfStream = (entries: PdfEntries, bytes: Uint8Array, deflate = true): PdfValue => ({
  kind: "stream",
  entries,
  bytes,
  deflate,
});

// A pdf real carries no exponent, so a number near zero has to be written out in
// full: `1e-7` is not a number to a reader, it is a syntax error. Six places is
// finer than any coordinate here means, page geometry being twips and text being
// measured to a thousandth of a point.
const PLACES = 6;

// A pdf's own limit on a whole number, and past it JavaScript writes an exponent
// too. Nothing on a page reaches this: the widest page Word will make is a little
// over 15000 points, so a coordinate out here is a fault upstream rather than a
// number worth writing.
const LARGEST = 2147483647;

export function formatPdfNumber(value: number): string {
  if (!Number.isFinite(value) || Math.abs(value) > LARGEST) {
    throw new DocxPagesError({
      code: "pdf-number-unwritable",
      message: "a pdf holds no exponent, no infinity and no not-a-number",
      at: AT,
      context: { value: String(value) },
    });
  }
  if (Number.isInteger(value)) return String(value);

  const fixed = value.toFixed(PLACES).replace(/0+$/, "").replace(/\.$/, "");
  // `-0.0000001` rounds to `-0`, which is a number a reader accepts and a sign
  // nothing here means.
  return fixed === "-0" ? "0" : fixed;
}

// The characters a name may not carry: whitespace, the delimiters, and `#`
// itself, which is the escape. Everything else goes through as written, which is
// what keeps `/Type` and `/FontFile2` readable in the output.
const ESCAPED_IN_NAME = /[^!-~]|[#()<>[\]{}/%]/g;

const escapeName = (name: string): string =>
  name.replace(ESCAPED_IN_NAME, (character) => {
    const code = character.charCodeAt(0);
    if (code > 0xff) {
      throw new DocxPagesError({
        code: "pdf-name-unwritable",
        message: "a pdf name is bytes, and this one holds a character outside them",
        at: AT,
        context: { name },
      });
    }
    return `#${code.toString(16).padStart(2, "0")}`;
  });

// A literal string ends at its own closing bracket, so the brackets inside one
// and the backslash that escapes them are escaped in turn.
const escapeLiteral = (text: string): string =>
  text.replace(/[\\()\r]/g, (character) => (character === "\r" ? "\\r" : `\\${character}`));

const HEX = "0123456789abcdef";

function hexOf(bytes: Uint8Array): string {
  let out = "";
  for (const byte of bytes) out += `${HEX[byte >> 4] ?? "0"}${HEX[byte & 0xf] ?? "0"}`;
  return out;
}

// Every string here is pdf syntax, which is bytes rather than characters: latin1
// is the encoding that writes one for one.
const latin1 = (text: string): Uint8Array => strToU8(text, true);

// The bytes of a document, gathered rather than concatenated as they go, since a
// pdf holds whole font and image files and joining those repeatedly would copy
// them once per object written.
type Buffer = {
  readonly push: (bytes: Uint8Array) => void;
  readonly write: (text: string) => void;
  readonly length: () => number;
  readonly bytes: () => Uint8Array;
};

function buffer(): Buffer {
  const chunks: Uint8Array[] = [];
  let length = 0;

  const push = (bytes: Uint8Array): void => {
    chunks.push(bytes);
    length += bytes.byteLength;
  };

  return {
    push,
    write: (text) => {
      push(latin1(text));
    },
    length: () => length,
    bytes: () => {
      const out = new Uint8Array(length);
      let at = 0;
      for (const chunk of chunks) {
        out.set(chunk, at);
        at += chunk.byteLength;
      }
      return out;
    },
  };
}

// An entry stating no value is left out rather than written as null, which is how
// a dictionary carries the parts of itself a document did not ask for.
const statedEntries = (entries: PdfEntries): readonly (readonly [string, PdfValue])[] =>
  Object.entries(entries).flatMap(([key, value]) =>
    value === undefined ? [] : [[key, value] as const],
  );

function writeValue(out: Buffer, value: PdfValue): void {
  switch (value.kind) {
    case "null":
      out.write("null");
      return;
    case "boolean":
      out.write(value.value ? "true" : "false");
      return;
    case "number":
      out.write(formatPdfNumber(value.value));
      return;
    case "name":
      out.write(`/${escapeName(value.name)}`);
      return;
    case "literal-string":
      out.write(`(${escapeLiteral(value.text)})`);
      return;
    case "hex-string":
      out.write(`<${hexOf(value.bytes)}>`);
      return;
    case "reference":
      out.write(`${String(value.number)} 0 R`);
      return;
    case "array":
      out.write("[");
      value.items.forEach((item, at) => {
        if (at > 0) out.write(" ");
        writeValue(out, item);
      });
      out.write("]");
      return;
    case "dictionary":
      writeDictionary(out, value.entries);
      return;
    case "stream": {
      const body = value.deflate ? zlibSync(value.bytes) : value.bytes;
      writeDictionary(out, {
        ...value.entries,
        ...(value.deflate ? { Filter: pdfName("FlateDecode") } : {}),
        Length: pdfNumber(body.byteLength),
      });
      out.write("\nstream\n");
      out.push(body);
      out.write("\nendstream");
      return;
    }
  }
}

function writeDictionary(out: Buffer, entries: PdfEntries): void {
  out.write("<<");
  for (const [key, value] of statedEntries(entries)) {
    out.write(`/${escapeName(key)} `);
    writeValue(out, value);
  }
  out.write(">>");
}

/**
 * A document being written: objects are numbered as they are reserved and filled
 * in whenever their contents are known.
 *
 * Reserving and filling are separate because a pdf's own shape is circular. A page
 * names the page tree it hangs off and the tree names every page in it, so one of
 * the two has to be referred to before it can be written.
 */
export type PdfObjects = {
  // A number for an object not yet written, which anything may refer to meanwhile.
  readonly reserve: () => PdfReference;
  readonly put: (at: PdfReference, value: PdfValue) => void;
  // Reserve and fill in one go, for an object nothing has to name before it exists.
  readonly add: (value: PdfValue) => PdfReference;
  // The whole file: header, every object, the table saying where each one starts,
  // and the trailer.
  readonly bytes: (trailer: PdfTrailer) => Uint8Array;
};

export type PdfTrailer = {
  readonly root: PdfReference;
  readonly info?: PdfReference;
};

// Marks the file as binary for anything that would otherwise take it for text and
// mangle its line endings. Four bytes above 127, which is what the specification
// asks for and every writer does.
const BINARY_COMMENT = Uint8Array.from([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]);

const HEADER = "%PDF-1.7\n";

// Every entry in the cross-reference table is exactly twenty bytes, offset and
// generation and kind, so that a reader can seek straight to the object it wants.
const xrefEntry = (offset: number, kind: "n" | "f", generation = 0): string =>
  `${String(offset).padStart(10, "0")} ${String(generation).padStart(5, "0")} ${kind} \n`;

export function pdfObjects(): PdfObjects {
  const values = new Map<number, PdfValue>();
  let count = 0;

  const reserve = (): PdfReference => {
    count += 1;
    return { kind: "reference", number: count };
  };

  const put = (at: PdfReference, value: PdfValue): void => {
    if (values.has(at.number)) {
      throw new DocxPagesError({
        code: "pdf-object-written-twice",
        message: "an object was filled in twice, and a pdf holds one of each",
        at: AT,
        context: { object: at.number },
      });
    }
    values.set(at.number, value);
  };

  return {
    reserve,
    put,
    add: (value) => {
      const at = reserve();
      put(at, value);
      return at;
    },
    bytes: (trailer) => {
      const out = buffer();
      out.write(HEADER);
      out.push(BINARY_COMMENT);

      const offsets: number[] = [];
      for (let number = 1; number <= count; number += 1) {
        const value = values.get(number);
        if (value === undefined) {
          throw new DocxPagesError({
            code: "pdf-object-unwritten",
            message: "an object was reserved and never filled in",
            at: AT,
            context: { object: number },
          });
        }
        offsets.push(out.length());
        out.write(`${String(number)} 0 obj\n`);
        writeValue(out, value);
        out.write("\nendobj\n");
      }

      // Object zero is the head of the list of free ones and is always there,
      // which is why the size is one more than the objects written.
      const startedAt = out.length();
      out.write(`xref\n0 ${String(count + 1)}\n`);
      out.write(xrefEntry(0, "f", 65535));
      for (const offset of offsets) out.write(xrefEntry(offset, "n"));

      out.write("trailer\n");
      writeDictionary(out, {
        Size: pdfNumber(count + 1),
        Root: trailer.root,
        ...(trailer.info === undefined ? {} : { Info: trailer.info }),
      });
      out.write(`\nstartxref\n${String(startedAt)}\n%%EOF\n`);

      return out.bytes();
    },
  };
}
