// Faces already handed over, so that a document redrawn does not hand the same
// bytes to the browser again.
const offered = new Set<string>();

const keyOf = (name: string, bold: boolean, italic: boolean): string =>
  `${name.toLowerCase()}|${bold ? "b" : ""}${italic ? "i" : ""}`;

/**
 * Lets the browser draw a face under the name the layout measured it as: a
 * supplied face under its own name, and a stood-in face under the name the
 * document asked for, so what is painted is the very bytes that were measured.
 * A runtime without the FontFace API paints whatever its styles find, at the
 * measured widths; that is the one way a page here is right in its geometry and
 * wrong on the screen, and it is the runtime's limit rather than a quiet choice.
 *
 * **The browser keeps one set of faces for the whole page and matches them by
 * family name.** Two documents drawn on one page, each supplying its own bytes
 * under one name, are both painted in whichever face got there first, though each
 * was measured in its own. Nothing here can part them: the name is all the browser
 * matches on, so parting them would mean giving each document a family name of its
 * own and drawing every run of it under that.
 */
export function offerToBrowser(
  name: string,
  bold: boolean,
  italic: boolean,
  bytes: Uint8Array,
): void {
  if (typeof FontFace === "undefined" || typeof document === "undefined") return;
  const key = keyOf(name, bold, italic);
  if (offered.has(key)) return;
  offered.add(key);

  const face = new FontFace(name, bytes.slice().buffer, {
    weight: bold ? "bold" : "normal",
    style: italic ? "italic" : "normal",
  });
  document.fonts.add(face);
  face.load().catch(() => {
    // A face the browser refuses stays measured and unpainted, as above, and is
    // taken back out rather than left standing as the answer for its name: bytes
    // offered under it later are then offered rather than passed over.
    document.fonts.delete(face);
    offered.delete(key);
  });
}
