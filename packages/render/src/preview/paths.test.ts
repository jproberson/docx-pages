import { join, resolve } from "node:path";
import { describe, expect, it } from "vitest";

import { pathOf } from "./paths.js";

const ROOT = resolve("/where", "samples");
const START = "/preview/index.html";
const asked = (url: string): string => pathOf(url, ROOT, START);

describe("pathOf", () => {
  it("lands the bare root on the page the server was started for", () => {
    expect(asked("/")).toBe(join(ROOT, "preview", "index.html"));
  });

  it("takes a document under the root", () => {
    expect(asked("/check/one.pdf")).toBe(join(ROOT, "check", "one.pdf"));
  });

  it("reads a name back out of its encoding", () => {
    expect(asked("/check/one%20two.pdf")).toBe(join(ROOT, "check", "one two.pdf"));
  });

  it("refuses a climb out of the root that survives the url as an encoded slash", () => {
    expect(asked("/..%2fsecrets/one.pdf")).toBe(ROOT);
  });

  it("refuses a sibling directory whose name begins with the root's", () => {
    expect(asked("/..%2fsamples-backup/one.pdf")).toBe(ROOT);
  });
});
