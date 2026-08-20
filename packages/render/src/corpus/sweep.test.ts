import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { buildAuthoredDocx } from "../authored/package.js";
import { corpusFaces } from "./faces.js";
import {
  changesBetween,
  described,
  summaryOf,
  sweepCorpus,
  sweepDocument,
  type SweptDocument,
} from "./sweep.js";

const FACES = corpusFaces();

const document = (body: string): Uint8Array => buildAuthoredDocx({ body });

const swept = (each: Partial<SweptDocument>): SweptDocument => ({
  id: "aaaaaaaaaaaa",
  bytes: 100,
  asks: [],
  facesStoodIn: 0,
  outcome: { kind: "laid-out" },
  pages: 1,
  millis: 1,
  ...each,
});

describe("sweepDocument", () => {
  it("reads what a document asks for without needing to lay it out", () => {
    const hidden = document(`<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>a</w:t></w:r></w:p>`);
    expect(sweepDocument(hidden, []).asks).toStrictEqual(["hidden-text"]);
  });

  it("names a document it cannot open rather than throwing out of the sweep", () => {
    const nonsense = new Uint8Array([1, 2, 3, 4, 5]);
    const each = sweepDocument(nonsense, FACES);
    expect(each.outcome.kind).not.toBe("laid-out");
    expect(each.asks).toStrictEqual([]);
  });

  // Two runs of a sweep have to agree about which document is which, and the only
  // thing they share is the bytes.
  it("gives the same document the same identity every time", () => {
    const bytes = document(`<w:p><w:r><w:t>a</w:t></w:r></w:p>`);
    expect(sweepDocument(bytes, FACES).id).toBe(sweepDocument(bytes, FACES).id);
    expect(sweepDocument(bytes, FACES).id).not.toBe(
      sweepDocument(document(`<w:p><w:r><w:t>b</w:t></w:r></w:p>`), FACES).id,
    );
  });

  it("says nothing about a document but the hash of its bytes", () => {
    const each = sweepDocument(document(`<w:p><w:r><w:t>a</w:t></w:r></w:p>`), FACES);
    expect(Object.keys(each).sort()).toStrictEqual([
      "asks",
      "bytes",
      "facesStoodIn",
      "id",
      "millis",
      "outcome",
      "pages",
    ]);
  });
});

describe("summaryOf", () => {
  it("counts a kind by the documents it was met in as well as by how often", () => {
    const summary = summaryOf([
      swept({ id: "one", asks: ["highlighting", "highlighting", "footnote"] }),
      swept({ id: "two", asks: ["footnote"] }),
    ]);
    expect(summary.kinds).toStrictEqual([
      { kind: "footnote", documents: 2, occurrences: 2 },
      { kind: "highlighting", documents: 1, occurrences: 2 },
    ]);
  });

  it("keeps what a document was refused for apart from what it asked for", () => {
    const summary = summaryOf([
      swept({ id: "one", outcome: { kind: "threw", detail: "boom" }, pages: null }),
      swept({ id: "two" }),
    ]);
    expect(summary.laidOut).toBe(1);
    expect(summary.threw).toBe(1);
    expect(summary.failures).toStrictEqual([{ kind: "threw: boom", documents: 1, occurrences: 1 }]);
  });
});

describe("changesBetween", () => {
  it("says which documents came out differently", () => {
    const changes = changesBetween(
      [swept({ id: "one", pages: 2 }), swept({ id: "two", pages: 3 })],
      [swept({ id: "one", pages: 2 }), swept({ id: "two", pages: 4 })],
    );
    expect(changes.map((each) => each.id)).toStrictEqual(["two"]);
    expect(changes[0]?.was).toContain("3 pages");
    expect(changes[0]?.now).toContain("4 pages");
  });

  // A corpus grows. A document nobody had before is not a change to anything.
  it("passes over a document that was not in the earlier run", () => {
    expect(changesBetween([swept({ id: "one" })], [swept({ id: "new", pages: 9 })])).toStrictEqual(
      [],
    );
  });

  it("does not read a difference in timing as a change", () => {
    const changes = changesBetween(
      [swept({ id: "one", millis: 5 })],
      [swept({ id: "one", millis: 5000 })],
    );
    expect(changes).toStrictEqual([]);
  });
});

// A corpus gathered from the wild holds the same document under several names, and
// counting one of those as several documents is what weighted every share the sweep
// ever reported.
describe("sweepCorpus over a corpus holding the same document twice", () => {
  it("sweeps a document once however many names it is filed under", () => {
    const directory = mkdtempSync(join(tmpdir(), "docx-pages-corpus-"));
    const bytes = document(`<w:p><w:r><w:t>one</w:t></w:r></w:p>`);
    writeFileSync(join(directory, "one.docx"), bytes);
    writeFileSync(join(directory, "again.docx"), bytes);
    writeFileSync(join(directory, "other.docx"), document(`<w:p><w:r><w:t>two</w:t></w:r></w:p>`));

    const swept = sweepCorpus(directory);
    expect(swept).toHaveLength(2);
    expect(new Set(swept.map((each) => each.id)).size).toBe(2);
  });
});

// A report is meant to be safe to keep and to quote, and the only thing standing
// between a real document's name and one is what a failure is allowed to say.
describe("described", () => {
  it("says only the errno of a file it could not read", () => {
    const thrown = ((): unknown => {
      try {
        readFileSync(join(mkdtempSync(join(tmpdir(), "docx-pages-corpus-")), "Some Document.docx"));
        return null;
      } catch (error) {
        return error;
      }
    })();

    expect(described(thrown)).toBe("ENOENT");
  });

  it("takes the path and the quoted text out of anything else thrown", () => {
    expect(described(new Error("cannot parse '/home/someone/Their Report.docx' at byte 12"))).toBe(
      "cannot parse '...' at byte 12",
    );
    expect(described(new Error("unexpected token in /home/someone/Their Report.docx"))).toBe(
      "unexpected token in ...",
    );
  });
});
