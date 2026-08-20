import { fileURLToPath } from "node:url";

import { defineConfig } from "vitest/config";

const resolveSrc = (path) => fileURLToPath(new URL(`./packages/${path}`, import.meta.url));

export const config = defineConfig({
  resolve: {
    // Tests run against source, so `pnpm test` never depends on a prior build.
    // Anchored patterns keep the bare package name from swallowing its subpaths.
    alias: [
      {
        find: /^@docx-pages\/core\/testing$/,
        replacement: resolveSrc("core/src/testing/index.ts"),
      },
      // The plumbing under the library's own promise. The package's `exports`
      // does not offer it, so only this repository can reach it at all.
      {
        find: /^@docx-pages\/core\/internal$/,
        replacement: resolveSrc("core/src/internal.ts"),
      },
      { find: /^@docx-pages\/core$/, replacement: resolveSrc("core/src/index.ts") },
      { find: /^@docx-pages\/fonts\/node$/, replacement: resolveSrc("fonts/src/node.ts") },
      { find: /^@docx-pages\/fonts$/, replacement: resolveSrc("fonts/src/index.ts") },
      { find: /^@docx-pages\/render$/, replacement: resolveSrc("render/src/index.ts") },
      { find: /^@docx-pages\/viewer$/, replacement: resolveSrc("viewer/src/index.ts") },
    ],
  },
  test: {
    include: ["packages/*/src/**/*.test.ts", "packages/*/src/**/*.test.tsx"],
    // Reading a pdf back loads pdfjs first, and the file that happens to do it
    // while every other file is running waits far longer than a unit test's own
    // patience allows for.
    testTimeout: 60_000,
    // A machine without the reference documents, without Word, or without Calibri
    // has whole suites report nothing to run rather than fail, so a green tick
    // means less there than here. Naming every test is how a run says which of
    // them it actually was.
    reporters: process.env["CI"] === undefined ? ["default"] : ["verbose"],
  },
});

export default config;
