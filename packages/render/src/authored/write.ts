import { mkdirSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { authoredDocuments } from "./documents.js";

// The documents themselves are built from the source beside them, so they are
// written out rather than kept: what is worth committing is what Word says about
// them, not the zip.
//
// Walked to from this module rather than taken off the working directory, since a
// run started elsewhere would write them into a directory Word has not been granted
// and a grant is per directory.
export const AUTHORED_DIRECTORY = fileURLToPath(
  new URL("../../../../samples/authored", import.meta.url),
);

export const authoredPath = (id: string): string => resolve(AUTHORED_DIRECTORY, `${id}.docx`);

export function writeAuthoredDocuments(): readonly string[] {
  mkdirSync(AUTHORED_DIRECTORY, { recursive: true });
  return authoredDocuments().map((each) => {
    const path = authoredPath(each.id);
    writeFileSync(path, each.bytes);
    return path;
  });
}

// Compared against this module's own path, which holds under tsx and from a build.
if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  for (const path of writeAuthoredDocuments()) process.stdout.write(`${path}\n`);
}
