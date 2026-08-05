import { describe, expect, it } from "vitest";

import { lookupFontMetrics, openDocx, type FaceRequest } from "@docx-pages/core";
import {
  buildDocx,
  buildMetafile,
  metafileHeader,
  metafileRecord,
  wordDocument,
} from "@docx-pages/core/testing";

import { imageDataUrl, imageResolver } from "./images.js";

const PIXEL = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const metricsFor = (request: FaceRequest) => lookupFontMetrics(request);

const packageWith = (parts: Readonly<Record<string, Uint8Array>>) =>
  openDocx(buildDocx({ "word/document.xml": wordDocument("<w:p/>"), ...parts }));

// A metafile whose whole drawing is one block of colour: a brush, the choice of it
// and a source-less copy of it into a rectangle.
const PATCOPY = 0x00f00021;

const METAFILE = buildMetafile([
  metafileHeader({ frameWidth: 2540, frameHeight: 2540 }),
  metafileRecord(39, [1, 0, 0x0000ff, 0]),
  metafileRecord(37, [1]),
  // prettier-ignore
  metafileRecord(76, [0, 0, 0, 0, 0, 0, 20, 20, PATCOPY, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
]);

describe("imageDataUrl", () => {
  it("encodes a part with the media type its extension names", () => {
    const url = imageDataUrl(
      packageWith({ "word/media/image1.png": PIXEL }),
      "word/media/image1.png",
    );
    expect(url).toBe("data:image/png;base64,iVBORw0KGgo=");
  });

  it("reads jpeg from either spelling of the extension", () => {
    const pkg = packageWith({ "a.jpg": PIXEL, "b.jpeg": PIXEL });
    expect(imageDataUrl(pkg, "a.jpg")?.startsWith("data:image/jpeg;")).toBe(true);
    expect(imageDataUrl(pkg, "b.jpeg")?.startsWith("data:image/jpeg;")).toBe(true);
  });

  it("declines a part it cannot name a media type for rather than guessing one", () => {
    expect(imageDataUrl(packageWith({ "a.emf": METAFILE }), "a.emf")).toBeUndefined();
  });

  it("declines a part the package does not have", () => {
    expect(imageDataUrl(packageWith({}), "word/media/absent.png")).toBeUndefined();
  });

  it("encodes an image larger than one chunk of the argument limit", () => {
    const large = new Uint8Array(0x8000 * 2 + 5).fill(0x41);
    const url = imageDataUrl(packageWith({ "big.png": large }), "big.png");
    expect(url?.length).toBeGreaterThan(0x8000);
  });
});

describe("imageResolver", () => {
  it("encodes each part once", () => {
    const resolve = imageResolver(packageWith({ "a.png": PIXEL }), metricsFor);
    expect(resolve("a.png")).toBe(resolve("a.png"));
  });

  it("remembers that a part could not be encoded", () => {
    const resolve = imageResolver(packageWith({}), metricsFor);
    expect(resolve("absent.png")).toBeUndefined();
    expect(resolve("absent.png")).toBeUndefined();
  });

  // A browser has no decoder for a metafile, so it is played into the shapes it
  // draws instead of being handed over as bytes nothing can show.
  it("plays a metafile into its shapes rather than encoding it", () => {
    const resolve = imageResolver(packageWith({ "a.emf": METAFILE }), metricsFor);
    expect(resolve("a.emf")).toMatchObject({
      kind: "metafile",
      picture: { shapes: [{ kind: "fill", color: "#ff0000" }] },
    });
  });

  it("declines a metafile it cannot play, which is marked rather than half drawn", () => {
    const resolve = imageResolver(packageWith({ "a.emf": PIXEL }), metricsFor);
    expect(resolve("a.emf")).toBeUndefined();
  });
});
