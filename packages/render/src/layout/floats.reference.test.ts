import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

import { layOutDocument, lookupFontMetrics, openDocx, type FontMetrics } from "@onepager/core";

const SAMPLES = resolve(process.env["ONEPAGER_SAMPLE_DIR"] ?? "samples");
const MERIDIAN_SANS: FontMetrics = { unitsPerEm: 1000, ascender: 971, descender: -242, lineGap: 0 };
const sample = resolve(SAMPLES, "Reference.docx");

const laidOut = () => {
  const result = layOutDocument(openDocx(new Uint8Array(readFileSync(sample))), (name) =>
    lookupFontMetrics(name, new Map([["Meridian Sans Medium", MERIDIAN_SANS]])),
  );
  if (result.kind !== "laid-out") throw new Error(`blocked: ${result.blocker.kind}`);
  return result;
};

const named = (name: string) => {
  const { headerFloats, bodyFloats } = laidOut();
  const found = [...headerFloats, ...bodyFloats].find((placed) => placed.anchor.name === name);
  if (found === undefined) throw new Error(`no float named ${name}`);
  return found;
};

// Left and top as Word drew them, read out of the reference pdf content stream.
const WORD = {
  "Picture 12": { leftPt: 445.35, topPt: 8.85 },
  "Picture 9": { leftPt: 25.5, topPt: 21.75 },
  "Picture 1": { leftPt: 303.75, topPt: 259.73 },
  "Picture 6": { leftPt: 303.75, topPt: 494.97 },
} as const;

describe.skipIf(!existsSync(sample))("float placement against Word", () => {
  it.each(Object.entries(WORD))("places %s horizontally where Word did", (name, expected) => {
    expect(named(name).leftPt).toBeCloseTo(expected.leftPt, 1);
  });

  it.each(Object.entries(WORD))(
    "places %s vertically within half a point of Word",
    (name, expected) => {
      expect(Math.abs(named(name).topPt - expected.topPt)).toBeLessThan(0.5);
    },
  );

  it("keeps the header logo above the header paragraph it is anchored to", () => {
    const logo = named("Picture 12");
    expect(logo.topPt).toBeLessThan(laidOut().headerTopPt);
    expect(logo.anchor.vertical).toStrictEqual({
      kind: "offset",
      from: "paragraph",
      offsetEmu: -162119,
    });
  });

  it("finds every floating object in the document", () => {
    const { headerFloats, bodyFloats } = laidOut();
    expect(headerFloats).toHaveLength(3);
    expect(bodyFloats.length).toBeGreaterThanOrEqual(10);
  });

  it("separates the tractor photo from the MEASUREMENTS box instead of overlapping them", () => {
    const photo = named("Picture 6");
    const box = named("Text Box 13");
    expect(photo.leftPt).toBeGreaterThan(box.leftPt + box.widthPt);
  });
});
