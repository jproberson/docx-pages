import { describe, expect, it } from "vitest";

import { readMeasured } from "./authored/measured.js";
import { authoredFace } from "./authored/faces.js";
import { referenceCases } from "./testing/cases.js";

/**
 * What a run has to have collected before a green tick means anything.
 *
 * **The three sets of documents go quiet separately and say so nowhere.** The
 * reference suites need a manifest that is never committed, the authored suite
 * needs a face this repository may not ship, and the corpus needs a directory of
 * other people's work. Each is written to report nothing to run rather than to
 * fail, which is right for a stranger's clone and wrong for the machine a release
 * is cut on: deleting the manifest drops 311 tests and still exits zero.
 *
 * So this is the one test that fails for finding nothing. It is off unless
 * `DOCX_PAGES_FULL_EVIDENCE` is set, because CI and a stranger both have every
 * right to run without any of it. Set it wherever a release is cut, and a run
 * that quietly proved almost nothing stops being a run that passed.
 *
 * The two lists are floors rather than counts, since evidence is added often and
 * lost rarely and it is losing it silently that this exists to catch. Word's own
 * answers are held to their exact number instead: `pnpm measure` rewrites that
 * file whole, and a run that drops an answer on the way is the way this file has
 * gone wrong before.
 */
const ASKED_FOR = process.env["DOCX_PAGES_FULL_EVIDENCE"] === undefined;

describe.skipIf(ASKED_FOR)("the evidence a release is cut on", () => {
  it("found the reference documents Word's answers are held against", () => {
    expect(referenceCases().length).toBeGreaterThanOrEqual(8);
  });

  it("found the face the authored suite is measured in", () => {
    expect(authoredFace()).not.toBeNull();
  });

  it("found Word's own answers about the authored documents", () => {
    expect(Object.keys(readMeasured().documents)).toHaveLength(56);
  });
});
