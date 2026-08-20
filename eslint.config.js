import js from "@eslint/js";
import globals from "globals";
import tseslint from "typescript-eslint";

const NODE_BUILTINS = [
  "fs",
  "node:fs",
  "fs/promises",
  "node:fs/promises",
  "path",
  "node:path",
  "child_process",
  "node:child_process",
  "os",
  "node:os",
  "crypto",
  "node:crypto",
  "url",
  "node:url",
];

/** Package source. Type-aware linting and the house rules apply here and only here. */
const SOURCE = ["packages/*/src/**/*.ts", "packages/*/src/**/*.tsx"];

/** Build tooling. Not part of any package, and legitimately needs default exports. */
const TOOLING = ["*.js", "*.config.ts"];

export default tseslint.config(
  { ignores: ["**/dist/**", "**/node_modules/**", "**/coverage/**", "probes/**"] },

  { files: [...SOURCE, ...TOOLING], ...js.configs.recommended },

  { files: TOOLING, languageOptions: { globals: globals.nodeBuiltin } },

  ...tseslint.configs.strictTypeChecked.map((entry) => ({ ...entry, files: SOURCE })),

  {
    files: SOURCE,
    languageOptions: {
      parserOptions: { projectService: true, tsconfigRootDir: import.meta.dirname },
    },
    rules: {
      // The house rules, enforced mechanically rather than by convention.
      "@typescript-eslint/no-explicit-any": "error",
      "@typescript-eslint/consistent-type-assertions": ["error", { assertionStyle: "never" }],
      "no-restricted-exports": ["error", { restrictDefaultExports: { direct: true, named: true } }],

      "@typescript-eslint/switch-exhaustiveness-check": "error",
      "@typescript-eslint/no-unnecessary-condition": "error",
    },
  },

  {
    // The viewer ships to a browser bundle. Node built-ins must never reach it.
    files: ["packages/viewer/src/**/*.ts", "packages/viewer/src/**/*.tsx"],
    rules: {
      "no-restricted-imports": [
        "error",
        {
          paths: NODE_BUILTINS.map((name) => ({
            name,
            message: "@docx-pages/viewer runs in the browser and cannot import Node built-ins.",
          })),
        },
      ],
    },
  },
);
