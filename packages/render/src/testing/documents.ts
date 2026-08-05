import { readFileSync } from "node:fs";

import { openDocx, type DocxPackage } from "@docx-pages/core";

import type { ReferenceCase } from "./cases.js";

export const readReferenceDocument = (each: ReferenceCase): DocxPackage =>
  openDocx(new Uint8Array(readFileSync(each.documentPath)));

export const readRenderedPages = (each: ReferenceCase): Uint8Array => {
  if (each.renderedPath === null) throw new Error(`case ${each.id} has no rendered pages`);
  return new Uint8Array(readFileSync(each.renderedPath));
};
