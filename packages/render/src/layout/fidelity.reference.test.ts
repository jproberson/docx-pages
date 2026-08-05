import { describe, expect, it } from "vitest";

import { readUnhonoured } from "@docx-pages/core";

import { authoredCases } from "../authored/cases.js";
import { referenceCases, type ReferenceCase } from "../testing/cases.js";
import { readReferenceDocument } from "../testing/documents.js";

// What each document says this project passed over. Every reference document
// matched Word line by line, so this is not a list of things known to be drawn
// wrong: it is what a document asks for that nothing here answers, whether or not
// it showed.
//
// The seven one-pagers report nothing at all, and that is the assertion worth
// having: a feature that stops being read has to be named here before it can move
// anything, so a gap cannot be introduced quietly.
const CASES: readonly ReferenceCase[] = [...referenceCases(), ...authoredCases()];

describe.skipIf(CASES.length === 0)("the fidelity report against every document", () => {
  for (const each of CASES) {
    it(`says what ${each.id} asks for and does not get`, () => {
      const report = readUnhonoured(readReferenceDocument(each));
      expect(report.map((entry) => entry.kind)).toStrictEqual(each.unhonoured);
    });
  }
});
