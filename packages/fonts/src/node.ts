import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import { defaultFaces, type ReadBytes } from "./index.js";
import type { FaceDefaults } from "@docx-pages/core";

// Node's fetch will not read a file: url, so the pack's own files are read off
// the disk instead. A separate entry keeps the import out of browser bundles.
export const readFromDisk: ReadBytes = async (url) =>
  new Uint8Array(await readFile(fileURLToPath(url)));

export const defaultFacesFromDisk = (): Promise<FaceDefaults> => defaultFaces(readFromDisk);
