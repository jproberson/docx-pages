import { afterEach, describe, expect, it, vi } from "vitest";

import { offerToBrowser } from "./offer-face.js";

// The browser's own font set, stood in for. What is asked of it is what matters:
// which faces it was handed, and which were taken away again.
type Offered = { readonly family: string; readonly load: () => Promise<unknown> };

const browserWhere = (loadFails: boolean) => {
  const held: Offered[] = [];
  const taken: Offered[] = [];

  class Face {
    constructor(readonly family: string) {}
    load(): Promise<unknown> {
      return loadFails ? Promise.reject(new Error("refused")) : Promise.resolve(this);
    }
  }

  vi.stubGlobal("FontFace", Face);
  vi.stubGlobal("document", {
    fonts: {
      add: (face: Offered) => held.push(face),
      delete: (face: Offered) => {
        taken.push(face);
        return true;
      },
    },
  });

  return { held, taken };
};

// The catch that forgets a refused face runs a turn later than the offer itself.
const settled = (): Promise<void> => new Promise((keep) => setTimeout(keep, 0));

afterEach(() => {
  vi.unstubAllGlobals();
});

const BYTES = Uint8Array.from([1, 2, 3]);

describe("offerToBrowser", () => {
  it("hands one face over once, however many documents name it", async () => {
    const browser = browserWhere(false);

    offerToBrowser("Twice Named", false, false, BYTES);
    offerToBrowser("Twice Named", false, false, BYTES);
    await settled();

    expect(browser.held.map((face) => face.family)).toStrictEqual(["Twice Named"]);
  });

  it("tells one cut of a face from another", () => {
    const browser = browserWhere(false);

    offerToBrowser("Cut Both Ways", false, false, BYTES);
    offerToBrowser("Cut Both Ways", true, false, BYTES);

    expect(browser.held).toHaveLength(2);
  });

  // A face the browser would not take is no answer for its name. Left standing it
  // would pass over the next bytes offered under that name, which is a document
  // painted in nothing over one lost load.
  it("takes a refused face back out, and offers the name again", async () => {
    const refused = browserWhere(true);

    offerToBrowser("Refused Once", false, false, BYTES);
    await settled();

    expect(refused.taken).toHaveLength(1);

    const again = browserWhere(false);
    offerToBrowser("Refused Once", false, false, BYTES);
    await settled();

    expect(again.held.map((face) => face.family)).toStrictEqual(["Refused Once"]);
  });

  // Neither is there where a page is drawn to a string rather than to a screen,
  // and a page drawn there is measured right and painted by whatever the styles
  // find, which is the runtime's limit rather than a fault to throw over.
  it("does nothing at all where there is no browser to offer a face to", () => {
    expect(() => {
      offerToBrowser("Unseen", false, false, BYTES);
    }).not.toThrow();
  });
});
