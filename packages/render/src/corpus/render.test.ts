import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { buildAuthoredDocx } from "../authored/package.js";
import { documentsWanted, parseRenderings } from "./render.js";

// What Word answered, one line per document. The script names each document by the
// path it was handed, which is the staged copy named for the document's own id, so
// an answer says which document it is about without anything here naming one.
describe("parseRenderings", () => {
  it("reads a document Word drew, and how long it took over it", () => {
    expect(parseRenderings("ok|3|/somewhere/samples/corpus/pdf/aaaaaaaaaaaa.docx\n")).toStrictEqual(
      [{ id: "aaaaaaaaaaaa", outcome: "ok", detail: "", seconds: 3 }],
    );
  });

  it("keeps the error Word gave for one it would not draw", () => {
    expect(parseRenderings("fail|-1712|/somewhere/bbbbbbbbbbbb.docx\n")).toStrictEqual([
      { id: "bbbbbbbbbbbb", outcome: "fail", detail: "-1712", seconds: 0 },
    ]);
  });

  // Word answers for a whole batch at once, so the run has to survive a batch that
  // was cut off partway through its own output.
  it("passes over anything that is not an answer", () => {
    const answered = parseRenderings(
      [
        "ok|1|/x/aaaaaaaaaaaa.docx",
        "",
        "Microsoft Word got an error",
        "ok|2|/x/cccccccccccc.docx",
      ].join("\n"),
    );
    expect(answered.map((each) => each.id)).toStrictEqual(["aaaaaaaaaaaa", "cccccccccccc"]);
  });
});

// A quarter of the files in a corpus gathered from the wild are copies of a
// document already in it. Two copies in one batch stage to the one path and ask
// Word to open a document it is already holding.
describe("documentsWanted", () => {
  it("asks for a document once however many names it is filed under", () => {
    const directory = mkdtempSync(join(tmpdir(), "docx-pages-render-"));
    const bytes = buildAuthoredDocx({ body: `<w:p><w:r><w:t>one</w:t></w:r></w:p>` });
    writeFileSync(join(directory, "one.docx"), bytes);
    writeFileSync(join(directory, "again.docx"), bytes);
    writeFileSync(
      join(directory, "other.docx"),
      buildAuthoredDocx({ body: `<w:p><w:r><w:t>two</w:t></w:r></w:p>` }),
    );

    const wanted = documentsWanted(directory);
    expect(wanted).toHaveLength(2);
    expect(new Set(wanted.map((each) => each.id)).size).toBe(2);
  });
});
