const encoder = new TextEncoder();

const concat = (chunks: readonly Uint8Array[]): Uint8Array => {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const out = new Uint8Array(total);
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }
  return out;
};

export type PdfFixture = {
  readonly contents: string;
  readonly widthPt?: number;
  readonly heightPt?: number;
};

export function buildPdf(fixture: PdfFixture): Uint8Array {
  const width = fixture.widthPt ?? 612;
  const height = fixture.heightPt ?? 792;
  const stream = encoder.encode(fixture.contents);

  const objects: Uint8Array[] = [
    encoder.encode("<</Type/Catalog/Pages 2 0 R>>"),
    encoder.encode("<</Type/Pages/Kids[3 0 R]/Count 1>>"),
    encoder.encode(
      `<</Type/Page/Parent 2 0 R/MediaBox[0 0 ${String(width)} ${String(height)}]` +
        `/Resources<</XObject<</Im0 5 0 R>>>>/Contents 4 0 R>>`,
    ),
    concat([
      encoder.encode(`<</Length ${String(stream.length)}>>\nstream\n`),
      stream,
      encoder.encode("\nendstream"),
    ]),
    concat([
      encoder.encode(
        "<</Type/XObject/Subtype/Image/Width 1/Height 1/ColorSpace/DeviceGray" +
          "/BitsPerComponent 8/Length 1>>\nstream\n",
      ),
      new Uint8Array([0]),
      encoder.encode("\nendstream"),
    ]),
  ];

  const chunks: Uint8Array[] = [encoder.encode("%PDF-1.7\n")];
  const offsets: number[] = [];
  let position = chunks[0]?.length ?? 0;

  objects.forEach((body, index) => {
    offsets.push(position);
    const piece = concat([
      encoder.encode(`${String(index + 1)} 0 obj\n`),
      body,
      encoder.encode("\nendobj\n"),
    ]);
    chunks.push(piece);
    position += piece.length;
  });

  const count = objects.length + 1;
  const rows = offsets
    .map((offset) => `${offset.toString().padStart(10, "0")} 00000 n \n`)
    .join("");
  chunks.push(
    encoder.encode(
      `xref\n0 ${String(count)}\n0000000000 65535 f \n${rows}` +
        `trailer\n<</Size ${String(count)}/Root 1 0 R>>\nstartxref\n${String(position)}\n%%EOF\n`,
    ),
  );

  return concat(chunks);
}
