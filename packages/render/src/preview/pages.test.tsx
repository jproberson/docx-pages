import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { buildFace } from "@docx-pages/core/testing";

import { buildAuthoredDocx, FACE } from "../authored/package.js";
import { caseOf } from "./documents.js";
import { writeBrowser, writeFonts, writePreview, type FaceSet } from "./pages.js";

// The previewer draws what a reader judges by eye, so nothing here checks a
// position: that is the suites' work. What it checks is that the page is written
// at all, that it is named where the browser looks for it, and that Word's own pdf
// is wired to the pane beside it. Each of those breaks silently.

const METRICS = { unitsPerEm: 1000, ascender: 800, descender: -200, lineGap: 0 };

// The face the authored documents state, supplied rather than found, so this runs
// on a machine with no Calibri and no manifest.
const faceSet = (files: FaceSet["files"] = []): FaceSet => ({
  id: "word",
  label: "As Word rendered it",
  faces: [buildFace({ name: FACE, metrics: METRICS })],
  files,
});

const paragraph = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

let root = "";
const at = (...parts: string[]): string => join(root, ...parts);

function documentAt(directory: string, id: string, body: string): string {
  mkdirSync(directory, { recursive: true });
  const path = resolve(directory, `${id}.docx`);
  writeFileSync(path, buildAuthoredDocx({ body }));
  return path;
}

beforeAll(() => {
  root = mkdtempSync(join(tmpdir(), "docx-pages-pages-"));
});

afterAll(() => {
  rmSync(root, { recursive: true, force: true });
});

describe("writePreview", () => {
  it("draws a page the browser can find by the name it asks for", () => {
    const into = at("drawn");
    const path = documentAt(into, "asked", paragraph("A paragraph to draw."));

    const written = writePreview(into, caseOf("asked", path), faceSet(), "hidden");

    expect(written).toBe(resolve(into, "asked.word.html"));
    const html = readFileSync(written, "utf8");
    expect(html).toContain(`data-docx-page="0"`);
    // Every run is drawn into a tspan of its own at a width of its own, so a
    // sentence never appears in the markup as one string.
    expect(html).toContain(">paragraph</tspan>");
    expect(html).toContain(`href="fonts.word.css"`);
  });

  it("names the outlined drawing apart from the plain one", () => {
    const into = at("outlined");
    const path = documentAt(into, "asked", paragraph("Outlined."));
    const set = faceSet();

    expect(writePreview(into, caseOf("asked", path), set, "hidden")).toBe(
      resolve(into, "asked.word.html"),
    );
    expect(writePreview(into, caseOf("asked", path), set, "outlined")).toBe(
      resolve(into, "asked.word.frames.html"),
    );
  });

  it("draws a page for every page the document makes", () => {
    const into = at("paged");
    const broken = `${paragraph("First page.")}<w:p><w:r><w:br w:type="page"/></w:r></w:p>${paragraph("Second page.")}`;
    const path = documentAt(into, "asked", broken);

    const html = readFileSync(
      writePreview(into, caseOf("asked", path), faceSet(), "hidden"),
      "utf8",
    );

    // The bare name is in the stylesheet and in the script that fits the page, so
    // only the attribute itself counts the pages.
    expect(html.split(`data-docx-page="`).length - 1).toBe(2);
  });
});

describe("writeFonts", () => {
  it("carries each face's own file in beside the page that needs it", () => {
    const into = at("fonts");
    const file = resolve(at("faces"), "Twin Sans.ttf");
    mkdirSync(at("faces"), { recursive: true });
    writeFileSync(file, "the face's bytes");

    const written = writeFonts(
      into,
      faceSet([{ name: "Twin Sans", bold: false, italic: false, filePath: file }]),
    );
    const css = readFileSync(written, "utf8");

    expect(written).toBe(resolve(into, "fonts.word.css"));
    expect(css).toContain(`font-family: "Twin Sans"`);
    // A name with a space in it would break the url, so the copy is renamed.
    expect(css).toContain(`url("fonts/Twin-Sans.ttf")`);
    expect(readdirSync(resolve(into, "fonts"))).toContain("Twin-Sans.ttf");
  });

  it("passes over a face whose file this machine has not got", () => {
    const into = at("missingFace");
    const set = faceSet([
      { name: "Absent", bold: false, italic: false, filePath: resolve(at("faces"), "Absent.ttf") },
    ]);

    expect(readFileSync(writeFonts(into, set), "utf8").trim()).toBe("");
  });
});

describe("writeBrowser", () => {
  // The pane beside the rendering is the whole point of the previewer, and it is
  // wired by a path worked out relative to the directory the page is written to.
  it("wires Word's own pdf to the pane beside the rendering", () => {
    const into = at("beside");
    const path = documentAt(into, "asked", paragraph("Beside Word's."));
    writeFileSync(resolve(into, "asked.pdf"), "a pdf");

    const html = readFileSync(
      writeBrowser(into, [caseOf("asked", path)], [faceSet()], (each) => each.id),
      "utf8",
    );

    expect(html).toContain(`data-word="asked.pdf"`);
  });

  it("leaves the pane empty where Word rendered nothing", () => {
    const into = at("alone");
    const path = documentAt(into, "asked", paragraph("Alone."));

    const html = readFileSync(
      writeBrowser(into, [caseOf("asked", path)], [faceSet()], (each) => each.id),
      "utf8",
    );

    expect(html).toContain(`data-word=""`);
  });

  it("offers every document and every face set to choose between", () => {
    const into = at("chosen");
    const first = caseOf("first", documentAt(into, "first", paragraph("One.")));
    const second = caseOf("second", documentAt(into, "second", paragraph("Two.")));
    const authored = { ...faceSet(), id: "authored", label: "In the authored face" };

    const html = readFileSync(
      writeBrowser(into, [first, second], [faceSet(), authored], (each) => `the ${each.id}`),
      "utf8",
    );

    expect(html).toContain(`>the first<`);
    expect(html).toContain(`>the second<`);
    expect(html).toContain(`value="authored"`);
  });
});
