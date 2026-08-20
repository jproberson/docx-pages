import { tmpdir } from "node:os";
import { isAbsolute } from "node:path";

import { afterEach, describe, expect, it } from "vitest";

import { MEASURED_PATH, readMeasured } from "./measured.js";

const startedIn = process.cwd();
afterEach(() => {
  process.chdir(startedIn);
});

// Word's answers about the authored documents are committed, so the suite that
// reads them goes quiet over a machine without Word's Calibri and over nothing
// else. A path read against the working directory used to empty it instead, from
// anywhere but the repository root, and say nothing about having done so.
describe("readMeasured", () => {
  it("finds the measurements from wherever the run was started", () => {
    expect(isAbsolute(MEASURED_PATH)).toBe(true);
    process.chdir(tmpdir());
    expect(Object.keys(readMeasured().documents).length).toBeGreaterThan(0);
  });

  it("says so rather than answering nothing when there are none to read", () => {
    expect(() => readMeasured("/nowhere/measured.json")).toThrow(/no measurements/);
  });
});
