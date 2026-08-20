import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { LayoutBlocked, looksOf, type Workspace } from "./compare.js";

const workspace: Workspace = {
  directory: tmpdir(),
  stylesheet: "fonts.css",
  profile: join(tmpdir(), "profile"),
  keep: false,
};

// Word's side is never read here: each of these gives up before it is asked for.
const drawnPath = ((): string => {
  const path = join(mkdtempSync(join(tmpdir(), "docx-pages-compare-")), "word.pdf");
  writeFileSync(path, "not a pdf");
  return path;
})();

const looksWhenDrawingThrows = (thrown: Error): Promise<{ readonly outcome: string }> =>
  looksOf(new Uint8Array(), "aaaaaaaaaaaa", drawnPath, workspace, undefined, () =>
    Promise.reject(thrown),
  );

// A gap this project knows about and one it does not are counted apart, and telling
// them apart is the whole worth of the two outcomes.
describe("looksOf over a document it could not draw", () => {
  it("counts a document it refused as blocked", async () => {
    const looks = await looksWhenDrawingThrows(new LayoutBlocked("rotated-text"));
    expect(looks.outcome).toBe("blocked");
  });

  it("counts anything else as threw, whatever the message says", async () => {
    const looks = await looksWhenDrawingThrows(new Error("blocked: not really"));
    expect(looks.outcome).toBe("threw");
  });
});
