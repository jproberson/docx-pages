import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";

const REFERENCE_DIR = resolve(process.env["ONEPAGER_REFERENCE_DIR"] ?? "samples/reference");

export const referencePath = (name: string): string => resolve(REFERENCE_DIR, name);

export const hasReference = (name: string): boolean => existsSync(referencePath(name));

export const readReference = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(referencePath(name)));
