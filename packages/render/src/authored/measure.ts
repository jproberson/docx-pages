import { execFileSync } from "node:child_process";
import { writeFileSync } from "node:fs";
import { basename, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { authoredDocuments } from "./documents.js";
import {
  MEASURED_PATH,
  type Measured,
  type MeasuredDocument,
  type MeasuredCharacter,
  type MeasuredParagraph,
  type MeasuredShape,
} from "./measured.js";
import { authoredPath, writeAuthoredDocuments } from "./write.js";

const SCRIPT = fileURLToPath(new URL("measure.applescript", import.meta.url));

// Word answers one document at a time and leaves each one open unless it is asked
// to close, so the whole run goes through a single script that closes behind
// itself. Splitting it up is what piles eighty documents into one Word.
function ask(paths: readonly string[]): string {
  return execFileSync("osascript", [SCRIPT, ...paths], {
    encoding: "utf8",
    maxBuffer: 64 * 1024 * 1024,
  });
}

const number = (value: string | undefined): number => {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) throw new Error(`expected a number, got ${String(value)}`);
  return parsed;
};

// One line per thing measured, each naming what it is: `DOC` opens a document and
// every `P` and `S` after it belongs to that one.
export function parseMeasurements(output: string): Measured {
  const documents: Record<string, MeasuredDocument> = {};
  let paragraphs: MeasuredParagraph[] = [];
  let characters: MeasuredCharacter[] = [];
  let shapes: MeasuredShape[] = [];

  for (const line of output.split("\n")) {
    const [kind, ...fields] = line.trim().split("|");

    if (kind === "DOC") {
      paragraphs = [];
      characters = [];
      shapes = [];
      documents[basename(fields[0] ?? "", ".docx")] = { paragraphs, characters, shapes };
    } else if (kind === "P") {
      paragraphs.push({
        index: number(fields[0]),
        page: number(fields[1]),
        topPt: number(fields[2]),
        leftPt: number(fields[3]),
      });
    } else if (kind === "C") {
      characters.push({
        paragraph: number(fields[0]),
        index: number(fields[1]),
        leftPt: number(fields[2]),
      });
    } else if (kind === "S") {
      shapes.push({
        name: fields[0] ?? "",
        widthPt: number(fields[1]),
        heightPt: number(fields[2]),
      });
    }
  }

  return { documents };
}

function main(): void {
  writeAuthoredDocuments();
  const paths = authoredDocuments().map(
    (each) => `${each.measuresCharacters === true ? "C" : "P"}:${authoredPath(each.id)}`,
  );
  const measured = parseMeasurements(ask(paths));

  const documents = Object.keys(measured.documents).length;
  if (documents === 0) throw new Error("Word answered nothing; is it holding a dialog open?");

  writeFileSync(MEASURED_PATH, `${JSON.stringify(measured, null, 2)}\n`);
  process.stdout.write(`${MEASURED_PATH}: ${String(documents)} documents\n`);
}

// Named `measure.js`, this guard never fired under tsx and the run did nothing at
// all, quietly. The entry is compared against this module's own path instead.
if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main();
}
