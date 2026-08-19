import { describe, expect, it } from "vitest";

import { buildAuthoredDocx } from "../authored/package.js";
import { censusOf, coveringDocuments, featuresIn, profileOf } from "./census.js";
import { rankGaps } from "./gaps.js";

const countsOf = (body: string): Readonly<Record<string, number>> =>
  censusOf(buildAuthoredDocx({ body })).counts;

const census = (id: string, features: readonly string[]) => ({
  id,
  bytes: 10,
  counts: Object.fromEntries(features.map((feature) => [feature, 1])),
});

describe("censusOf", () => {
  it("counts what a document holds, not what this project cannot honour", () => {
    // Nothing here is a gap, so the fidelity report says nothing at all about it
    // and the census still has to.
    const counts = countsOf(
      `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    );
    expect(counts["tables"]).toBe(1);
    expect(counts["table-rows"]).toBe(1);
    expect(counts["table-cells"]).toBe(1);
  });

  it("tells a nested table apart from one in the flow", () => {
    const inner = `<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`;
    const counts = countsOf(`<w:tbl><w:tr><w:tc>${inner}</w:tc></w:tr></w:tbl>`);
    expect(counts["tables"]).toBe(1);
    expect(counts["nested-tables"]).toBe(1);
  });

  // An equation is written in its own namespace, so the walk has to reach it before any
  // of it can be built. Read off the deformed ranking on 2026-08-12: the two documents
  // it named lose 16pt a page to one, and nothing in the corpus report could see them.
  it("counts an equation, which nothing here reads", () => {
    const fraction =
      `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">` +
      `<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>` +
      `</m:oMath>`;

    expect(countsOf(`<w:p>${fraction}</w:p>`)["equations"]).toBe(1);
  });

  it("counts how many faces a document asks for and never which", () => {
    const fonts = (face: string) =>
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="${face}"/></w:rPr><w:t>a</w:t></w:r></w:p>`;
    const counts = countsOf(
      fonts("Meridian Sans") + fonts("Granulation Serif") + fonts("Meridian Sans"),
    );
    expect(counts["distinct-faces"]).toBe(2);
    expect(JSON.stringify(counts)).not.toContain("Meridian");
  });

  it("says nothing at all about a package it cannot open", () => {
    const each = censusOf(new Uint8Array([1, 2, 3]));
    expect(each.counts).toStrictEqual({});
    expect(each.id).toHaveLength(12);
  });

  it("names the side an object wraps on, which is a thing to sample for", () => {
    const anchor = `<w:p><w:r><w:drawing><wp:anchor><wp:wrapSquare wrapText="largest"/></wp:anchor></w:drawing></w:r></w:p>`;
    const counts = countsOf(anchor);
    expect(counts["wrap-square"]).toBe(1);
    expect(counts["wrap-side-largest"]).toBe(1);
  });
});

describe("coveringDocuments", () => {
  it("takes the fewest documents that cover every feature the corpus has", () => {
    const corpus = [
      census("aaa", ["tables", "shading"]),
      census("bbb", ["tables"]),
      census("ccc", ["shading"]),
      census("ddd", ["notes"]),
    ];
    const chosen = coveringDocuments(corpus, 1);
    expect(chosen.map((each) => each.id).sort()).toStrictEqual(["aaa", "ddd"]);
  });

  it("takes a feature as often as it is asked for, where the corpus has it", () => {
    const corpus = [
      census("aaa", ["tables"]),
      census("bbb", ["tables"]),
      census("ccc", ["tables"]),
      census("ddd", ["tables"]),
    ];
    expect(coveringDocuments(corpus, 3)).toHaveLength(3);
  });

  // A corpus with one document holding a feature has one document's worth of it,
  // and asking for three cannot invent the other two.
  it("settles for what the corpus has where it has less than was asked for", () => {
    expect(coveringDocuments([census("aaa", ["notes"])], 3)).toHaveLength(1);
  });

  it("does not let a second copy of a document cover a feature twice over", () => {
    const twice = [census("aaa", ["notes"]), census("aaa", ["notes"])];
    expect(coveringDocuments(twice, 2).map((each) => each.id)).toStrictEqual(["aaa"]);
  });

  it("chooses the same sample twice from the same corpus", () => {
    const corpus = [
      census("bbb", ["tables", "notes"]),
      census("aaa", ["tables", "notes"]),
      census("ccc", ["shading"]),
    ];
    expect(coveringDocuments(corpus, 1).map((each) => each.id)).toStrictEqual(
      coveringDocuments([...corpus].reverse(), 1).map((each) => each.id),
    );
  });
});

describe("featuresIn", () => {
  it("counts a feature by documents as well as by how often it was met", () => {
    const tallies = featuresIn([
      { id: "aaa", bytes: 1, counts: { tables: 4, shading: 1 } },
      { id: "bbb", bytes: 1, counts: { tables: 1 } },
    ]);
    expect(tallies).toStrictEqual([
      { feature: "tables", documents: 2, occurrences: 5 },
      { feature: "shading", documents: 1, occurrences: 1 },
    ]);
  });

  // A corpus holds the same bytes under two names, and the sweep reads the second
  // copy no further, so a census that counts both puts the ranking's denominator
  // over its numerator.
  it("counts a document held twice in the corpus once", () => {
    const tallies = featuresIn([
      { id: "aaa", bytes: 1, counts: { tables: 4 } },
      { id: "aaa", bytes: 1, counts: { tables: 4 } },
      { id: "bbb", bytes: 1, counts: { tables: 1 } },
    ]);
    expect(tallies).toStrictEqual([{ feature: "tables", documents: 2, occurrences: 5 }]);
  });
});

describe("profileOf", () => {
  // How many faces a document asks for is a fact about the document, but it is a
  // number rather than a feature, and sampling on it would ask for one document
  // per distinct count.
  it("leaves the face count out of what a sample has to cover", () => {
    expect(
      profileOf({ id: "a", bytes: 1, counts: { tables: 1, "distinct-faces": 9 } }),
    ).toStrictEqual(["tables"]);
  });
});

describe("rankGaps", () => {
  const census = (id: string, counts: Readonly<Record<string, number>>) => ({
    id,
    bytes: 1,
    counts,
  });

  it("ranks a gap against the documents that could have met it, not the whole corpus", () => {
    const ranked = rankGaps(
      [
        { id: "aaa", asks: ["merged-cells"] },
        { id: "bbb", asks: [] },
        { id: "ccc", asks: [] },
      ],
      [
        census("aaa", { "table-cells": 4 }),
        census("bbb", { "table-cells": 2 }),
        census("ccc", { paragraphs: 9 }),
      ],
    );
    // Two documents have cells, one of them states a merge: the denominator is the
    // two, not the three.
    expect(ranked).toStrictEqual([{ kind: "merged-cells", met: 1, could: 2 }]);
  });

  it("ranks a gap no feature answers for against the whole corpus", () => {
    const ranked = rankGaps(
      [
        { id: "aaa", asks: ["keep-lines-together"] },
        { id: "bbb", asks: [] },
      ],
      [census("aaa", {}), census("bbb", {})],
    );
    expect(ranked[0]).toStrictEqual({ kind: "keep-lines-together", met: 1, could: 2 });
  });

  // The sweep reads a repeated document no further, so its numerator is over
  // distinct documents; a census counting every file would divide the one by the
  // other and state every share too low.
  it("ranks against distinct documents where the corpus holds one twice", () => {
    const ranked = rankGaps(
      [
        { id: "aaa", asks: ["merged-cells"] },
        { id: "bbb", asks: [] },
      ],
      [
        census("aaa", { "table-cells": 4 }),
        census("aaa", { "table-cells": 4 }),
        census("bbb", { "table-cells": 2 }),
      ],
    );
    expect(ranked).toStrictEqual([{ kind: "merged-cells", met: 1, could: 2 }]);
  });

  it("ranks a shape's own gap against the documents holding a shape", () => {
    const ranked = rankGaps(
      [
        { id: "aaa", asks: ["custom-geometry"] },
        { id: "bbb", asks: [] },
        { id: "ccc", asks: [] },
      ],
      [census("aaa", { shape: 2 }), census("bbb", { shape: 1 }), census("ccc", { paragraphs: 3 })],
    );
    expect(ranked).toStrictEqual([{ kind: "custom-geometry", met: 1, could: 2 }]);
  });

  it("counts a document once however many places it met a gap in", () => {
    const ranked = rankGaps(
      [{ id: "aaa", asks: ["highlighting", "highlighting"] }],
      [census("aaa", {})],
    );
    expect(ranked[0]?.met).toBe(1);
  });
});

describe("page geometry across sections", () => {
  // Every authored package ends with a section of its own, so a document here
  // states one more in the paragraph above it and the two are what is compared.
  const LETTER = `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>`;
  const breakingTo = (properties: string) =>
    countsOf(`<w:p><w:pPr><w:sectPr>${properties}</w:sectPr></w:pPr></w:p>`);

  // Only the last section's geometry is read, so a document whose sections agree
  // about the page asks nothing of that rule: counting sections alone would call
  // it a risk when it is not one.
  it("passes over a second section that makes the same page as the last", () => {
    const counts = breakingTo(LETTER);
    expect(counts["sections"]).toBe(2);
    expect(counts["distinct-page-geometries"]).toBe(1);
    expect(counts["more-than-one-page-geometry"]).toBeUndefined();
  });

  it("names a document whose sections make different pages", () => {
    const landscape = `<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>`;
    expect(breakingTo(landscape)["more-than-one-page-geometry"]).toBe(1);
  });

  // A section break that changes only the columns leaves the page alone.
  it("takes two sections differing only in their columns as the same page", () => {
    expect(breakingTo(`${LETTER}<w:cols w:num="2"/>`)["distinct-page-geometries"]).toBe(1);
  });

  it("counts a margin that differs as a different page", () => {
    const wider = LETTER.replace(`w:right="720"`, `w:right="1440"`);
    expect(breakingTo(wider)["more-than-one-page-geometry"]).toBe(1);
  });
});
