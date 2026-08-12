import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";

import {
  drawablesOf,
  layOutDocument,
  lookupFontMetrics,
  writePdf,
  type LaidOutDocument,
  type MetricsResolver,
  type PdfFont,
} from "@docx-pages/core";

import { readTextPlacements, type TextPlacement } from "./text.js";
import { referenceCases, referenceFonts, suppliedFaces } from "../testing/cases.js";
import { readReferenceDocument } from "../testing/documents.js";

// The writer over the real documents, which is a different question from the one
// the unit suite asks over documents built in memory.
//
// What is held here is the writer's own claim and not the layout's: **that a page
// written out draws its text where layout put it**. Whether layout is right is
// Word's business, and `text.reference.test.ts` is where Word is asked.
//
// These documents are what make it worth pinning. They are built out of text
// boxes, name faces by the dozen, write in symbol faces and number their lists;
// a writer that agrees with layout on `first line` and loses a text box on one of
// these has a fault the unit suite cannot see.

const CASES = referenceCases();

// The faces themselves, off this machine, since a pdf carries what it draws in.
const pdfFonts = (): readonly PdfFont[] =>
  referenceFonts().flatMap((font) =>
    font.filePath === null
      ? []
      : [
          {
            name: font.name,
            bold: font.bold,
            italic: font.italic,
            bytes: new Uint8Array(readFileSync(font.filePath)),
          },
        ],
  );

const round = (value: number): number => Math.round(value * 100) / 100;

const lineKey = (pageIndex: number, baselinePt: number): string =>
  `${String(pageIndex)}@${String(round(baselinePt))}`;

/**
 * Where layout drew text, walked through `drawablesOf` rather than through the
 * layout's own fields.
 *
 * The writer walks that one list, so anything reached another way would be a
 * second answer to the question of what a page holds: a document built out of text
 * boxes keeps most of its text inside them, and the header, the body and the
 * footer between them do not name a word of it.
 */
function linesLayoutDrew(layout: LaidOutDocument): ReadonlyMap<string, number> {
  const leftmost = new Map<string, number>();

  const at = (pageIndex: number, baselinePt: number, leftPt: number): void => {
    const key = lineKey(pageIndex, baselinePt);
    leftmost.set(key, Math.min(leftmost.get(key) ?? leftPt, leftPt));
  };

  for (const page of layout.pages) {
    for (const drawable of drawablesOf(layout, page)) {
      if (drawable.kind !== "text") continue;
      for (const box of drawable.boxes) {
        if (box.marker !== null && box.marker.text !== "") {
          at(page.index, box.marker.baselinePt, box.marker.leftPt);
        }
        for (const placed of box.lines) {
          for (const segment of placed.line.segments) {
            if (segment.kind !== "text" || segment.text === "") continue;
            at(
              page.index,
              placed.baselinePt - segment.mark.raisePt,
              placed.leftPt + segment.offsetPt,
            );
          }
        }
      }
    }
  }

  return leftmost;
}

// The leftmost thing the reader found on each baseline. It joins runs that carry
// on from one another and splits a run at a space of its own accord, so how many
// items come back says nothing; where each line begins survives all of it.
function linesRead(placements: readonly TextPlacement[]): ReadonlyMap<string, number> {
  const leftmost = new Map<string, number>();
  for (const placement of placements) {
    if (placement.text.trim() === "") continue;
    const key = lineKey(placement.pageIndex, placement.baselinePt);
    leftmost.set(key, Math.min(leftmost.get(key) ?? placement.leftPt, placement.leftPt));
  }
  return leftmost;
}

describe.skipIf(CASES.length === 0)("a reference document written out", () => {
  const faces = CASES.length === 0 ? [] : suppliedFaces();
  const metricsFor: MetricsResolver = (request) => lookupFontMetrics(request, faces);

  for (const each of CASES) {
    describe(each.id, () => {
      it("draws text on every baseline layout drew one, and starts each where layout did", async () => {
        const pkg = readReferenceDocument(each);
        const layout = layOutDocument(pkg, metricsFor);
        if (layout.kind !== "laid-out") throw new Error(`blocked: ${layout.blocker.kind}`);

        const bytes = writePdf(layout, {
          fonts: pdfFonts(),
          imageBytes: (part) => pkg.parts.get(part),
          metricsFor,
        });

        const drew = linesLayoutDrew(layout);
        const read = linesRead(await readTextPlacements(bytes));

        expect(read.size).toBeGreaterThan(0);

        // Every baseline layout drew on can be read back, which is the half that
        // says nothing was lost on the way out.
        //
        // Not the other half. A metafile records text of its own and is played
        // back as text rather than as a picture of it, so a document holding one
        // reads back baselines the layout's own boxes never named: the one case
        // here that carries a metafile reads four more, every one of them inside
        // it. Which is the writer working rather than failing.
        expect([...drew.keys()].filter((key) => !read.has(key))).toStrictEqual([]);
        expect(
          [...drew.entries()]
            .filter(([key, leftPt]) => round(read.get(key) ?? Number.NaN) !== round(leftPt))
            .map(([key, leftPt]) => ({ key, drew: round(leftPt), read: read.get(key) })),
        ).toStrictEqual([]);
      });
    });
  }
});
