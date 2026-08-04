import { describe, expect, it } from "vitest";

import { openDocx } from "@onepager/core";
import { buildDocx, wordDocument } from "@onepager/core/testing";

import { imageDataUrl, imageResolver } from "./images.js";

const PIXEL = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const packageWith = (parts: Readonly<Record<string, Uint8Array>>) =>
  openDocx(buildDocx({ "word/document.xml": wordDocument("<w:p/>"), ...parts }));

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
    expect(imageDataUrl(packageWith({ "a.emf": PIXEL }), "a.emf")).toBeUndefined();
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
    const resolve = imageResolver(packageWith({ "a.png": PIXEL }));
    expect(resolve("a.png")).toBe(resolve("a.png"));
  });

  it("remembers that a part could not be encoded", () => {
    const resolve = imageResolver(packageWith({}));
    expect(resolve("absent.png")).toBeUndefined();
    expect(resolve("absent.png")).toBeUndefined();
  });
});
