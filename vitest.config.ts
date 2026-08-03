import { fileURLToPath } from "node:url";

import { defineConfig } from "vitest/config";

const resolveSrc = (packageName) =>
  fileURLToPath(new URL(`./packages/${packageName}/src/index.ts`, import.meta.url));

export const config = defineConfig({
  resolve: {
    // Tests run against source, so `pnpm test` never depends on a prior build.
    alias: {
      "@onepager/core": resolveSrc("core"),
      "@onepager/render": resolveSrc("render"),
      "@onepager/viewer": resolveSrc("viewer"),
    },
  },
  test: {
    include: ["packages/*/src/**/*.test.ts", "packages/*/src/**/*.test.tsx"],
  },
});

export default config;
