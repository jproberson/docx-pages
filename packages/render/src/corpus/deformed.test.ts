import { describe, expect, it } from "vitest";

import { rankedBy, reportOf, type Deformed } from "./deformed.js";

// The ranking, over rows written by hand. What a page's shape means is settled in
// `pdf/agreement.test.ts`; what is settled here is that the order a reader is handed
// is the order of pages a preview cannot be shown for, and that a document laid out in
// a face this machine has not got is not in it.

// Word drew as many pages as we did unless the case says otherwise, since a count
// that disagrees is a fault of its own and no case here is about that one.
function row(id: string, of: Partial<Deformed>): Deformed {
  const made: Deformed = {
    id,
    outcome: "compared",
    facesStoodIn: 0,
    asks: [],
    pagesOurs: 1,
    pagesWord: 1,
    agrees: 0,
    shifted: 0,
    drifting: 0,
    deformed: 0,
    missing: 0,
    lines: 10,
    placed: 10,
    pages: [],
    detail: "",
    ...of,
  };
  return of.pagesWord === undefined ? { ...made, pagesWord: made.pagesOurs } : made;
}

describe("the order documents are ranked in", () => {
  it("puts the document with the most unshowable pages first, whatever else it does", () => {
    // The shifted document is wrong about far more lines: every line of all nine of
    // its pages is out of place, and every one of them is out of place the same way,
    // so the pages read as Word's own moved down. The other has one page nothing
    // explains, and that is the one a preview cannot show.
    const shifted = row("aaaaaaaaaaaa", { pagesOurs: 9, shifted: 9, placed: 0, lines: 300 });
    const deformed = row("bbbbbbbbbbbb", { pagesOurs: 9, agrees: 8, deformed: 1, placed: 290 });

    expect(rankedBy([shifted, deformed]).map((each) => each.id)).toStrictEqual([
      "bbbbbbbbbbbb",
      "aaaaaaaaaaaa",
    ]);
  });

  // Two pages of three is a document nobody can read; two pages of forty is a
  // document with two bad pages in it.
  it("breaks a tie by the share of itself a document cannot show", () => {
    const mostly = row("cccccccccccc", { pagesOurs: 3, deformed: 2, agrees: 1 });
    const partly = row("dddddddddddd", { pagesOurs: 40, deformed: 1, missing: 1, agrees: 38 });

    expect(rankedBy([partly, mostly]).map((each) => each.id)).toStrictEqual([
      "cccccccccccc",
      "dddddddddddd",
    ]);
  });

  it("counts a page missing content beside a page nothing explains", () => {
    const missing = row("eeeeeeeeeeee", { pagesOurs: 4, missing: 3, agrees: 1 });
    const deformed = row("ffffffffffff", { pagesOurs: 4, deformed: 2, agrees: 2 });

    expect(rankedBy([missing, deformed]).map((each) => each.id)).toStrictEqual([
      "eeeeeeeeeeee",
      "ffffffffffff",
    ]);
  });
});

describe("what the report says", () => {
  const rows = [
    row("aaaaaaaaaaaa", { pagesOurs: 2, agrees: 2 }),
    row("bbbbbbbbbbbb", { pagesOurs: 2, deformed: 1, agrees: 1 }),
    // A document laid out in a face this machine has not got has every line in the
    // wrong place before a rule is consulted, so it is counted and not ranked.
    row("cccccccccccc", { pagesOurs: 2, deformed: 2, facesStoodIn: 3 }),
    row("dddddddddddd", { outcome: "blocked", pagesOurs: 0 }),
  ];

  const report = reportOf(rows);

  it("ranks only the documents needing no face stood in", () => {
    const ranked = report.slice(report.indexOf("worst first:"));

    expect(ranked).toContain("bbbbbbbbbbbb");
    expect(ranked).not.toContain("cccccccccccc");
  });

  it("says how many documents were compared and how many were not", () => {
    expect(report).toContain("compared  3");
    expect(report).toContain("blocked   1");
  });

  it("counts the pages of every shape, and the documents with nothing wrong with them", () => {
    // Once over the three compared, where the two with a deformed page count against
    // it, and once over the two needing no face stood in.
    expect(report).toContain("every page showable in 1 documents (33.3%)");
    expect(report).toContain("every page showable in 1 documents (50.0%)");
    expect(report).toContain("deformed       3 pages");
  });

  it("says when a document makes the wrong number of pages", () => {
    const miscounted = reportOf([row("aaaaaaaaaaaa", { pagesOurs: 4, pagesWord: 5, agrees: 4 })]);

    expect(miscounted).toContain("1 of the clean make the wrong number of pages");
  });
});
