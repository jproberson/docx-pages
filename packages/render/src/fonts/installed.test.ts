import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { buildCollection, buildSfnt, type FontFixture } from "@docx-pages/core/testing";

import { installedFaces } from "./installed.js";

const FACE: FontFixture = { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 };

// A directory of fonts written for the test, since what a machine happens to have
// installed is not a thing to pin a suite against.
function directoryOf(files: Readonly<Record<string, Uint8Array>>): string {
  const directory = mkdtempSync(join(tmpdir(), "docx-pages-fonts-"));
  for (const [name, bytes] of Object.entries(files)) writeFileSync(join(directory, name), bytes);
  return directory;
}

const named = (
  faces: readonly { name: string; bold: boolean; italic: boolean }[],
): readonly string[] =>
  faces
    .map((each) => `${each.name}${each.bold ? " bold" : ""}${each.italic ? " italic" : ""}`)
    .sort();

describe("installedFaces", () => {
  // A file on disk is named for whoever shipped it: `seguisb.ttf` is Segoe UI
  // Semibold and `calibril.ttf` is Calibri Light, so the name to offer a document
  // has to come out of the font and never off the file.
  it("names a face by what it calls itself and not by its file", () => {
    const directory = directoryOf({
      "xyzzy1.ttf": buildSfnt({ ...FACE, faceName: "Meridian" }),
      "xyzzy2.ttf": buildSfnt({
        ...FACE,
        faceName: "Meridian Bold",
        familyName: "Meridian",
        bold: true,
      }),
    });

    expect(named(installedFaces([directory]))).toStrictEqual([
      "Meridian",
      "Meridian Bold",
      "Meridian bold",
    ]);
  });

  it("offers a face under its whole name as well, which is how a document may ask", () => {
    const directory = directoryOf({
      "a.ttf": buildSfnt({ ...FACE, faceName: "Meridian Light", familyName: "Meridian" }),
    });
    const faces = installedFaces([directory]);

    expect(named(faces)).toStrictEqual(["Meridian", "Meridian Light"]);
    // Asked for by its whole name, a light cut is a regular face of its own rather
    // than the family emboldened.
    expect(faces.find((each) => each.name === "Meridian Light")?.bold).toBe(false);
  });

  it("reads every face of a collection, which holds several behind one file", () => {
    const directory = directoryOf({
      "both.ttc": buildCollection([
        { ...FACE, faceName: "Meridian", advances: { A: 660 } },
        { ...FACE, faceName: "Meridian Maths", advances: { A: 480 } },
      ]),
    });
    const faces = installedFaces([directory]);
    const maths = faces.find((each) => each.name === "Meridian Maths");

    expect(named(faces)).toStrictEqual(["Meridian", "Meridian Maths"]);
    // The advances have to come out of the face asked for rather than the first in
    // the file, which is the whole reason a collection is read by name.
    expect(maths?.advances.kind === "advances" ? maths.advances.advanceFor(0x41) : null).toBe(480);
  });

  it("keeps the first face to claim a name, so an earlier directory wins", () => {
    const first = directoryOf({
      "a.ttf": buildSfnt({ ...FACE, faceName: "Meridian", advances: { A: 660 } }),
    });
    const second = directoryOf({
      "b.ttf": buildSfnt({ ...FACE, faceName: "Meridian", advances: { A: 100 } }),
    });
    const face = installedFaces([first, second])[0];

    expect(face?.advances.kind === "advances" ? face.advances.advanceFor(0x41) : null).toBe(660);
  });

  // A directory of fonts is not a thing to refuse a whole sweep over: what cannot
  // be read is a face this machine cannot offer, which is what it was before it was
  // looked for.
  it("passes over what it cannot read rather than giving up on the directory", () => {
    const directory = directoryOf({
      "broken.ttf": new Uint8Array([1, 2, 3, 4]),
      "notafont.txt": new Uint8Array([5, 6, 7, 8]),
      "good.ttf": buildSfnt({ ...FACE, faceName: "Meridian" }),
    });

    expect(named(installedFaces([directory]))).toStrictEqual(["Meridian"]);
  });

  it("answers with nothing for a directory that is not there", () => {
    expect(installedFaces([join(tmpdir(), "docx-pages-no-such-directory")])).toStrictEqual([]);
  });
});
