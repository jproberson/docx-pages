import { fileURLToPath } from "node:url";

import { defineConfig } from "vitest/config";

const resolveSrc = (path) => fileURLToPath(new URL(`./packages/${path}`, import.meta.url));

export const config = defineConfig({
  resolve: {
    // Tests run against source, so `pnpm test` never depends on a prior build.
    // Anchored patterns keep the bare package name from swallowing its subpaths.
    alias: [
      { find: /^@onepager\/core\/testing$/, replacement: resolveSrc("core/src/testing/index.ts") },
      { find: /^@onepager\/core$/, replacement: resolveSrc("core/src/index.ts") },
      { find: /^@onepager\/render$/, replacement: resolveSrc("render/src/index.ts") },
      { find: /^@onepager\/viewer$/, replacement: resolveSrc("viewer/src/index.ts") },
    ],
  },
  test: {
    include: ["packages/*/src/**/*.test.ts", "packages/*/src/**/*.test.tsx"],
    // Reading a pdf back loads pdfjs first, and the file that happens to do it
    // while every other file is running waits far longer than a unit test's own
    // patience allows for.
    testTimeout: 60_000,
  },
});

export default config;
