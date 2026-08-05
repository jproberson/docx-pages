import { mkdirSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

import { authoredDocuments } from "./documents.js";

// The documents themselves are built from the source beside them, so they are
// written out rather than kept: what is worth committing is what Word says about
// them, not the zip.
export const AUTHORED_DIRECTORY = "samples/authored";

export const authoredPath = (id: string): string => resolve(AUTHORED_DIRECTORY, `${id}.docx`);

export function writeAuthoredDocuments(): readonly string[] {
  mkdirSync(AUTHORED_DIRECTORY, { recursive: true });
  return authoredDocuments().map((each) => {
    const path = authoredPath(each.id);
    writeFileSync(path, each.bytes);
    return path;
  });
}

if (process.argv[1]?.endsWith("write.js") === true) {
  for (const path of writeAuthoredDocuments()) process.stdout.write(`${path}\n`);
}
