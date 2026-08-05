export type ContextValue = string | number | boolean | null | readonly ContextValue[];

export type ErrorContext = Readonly<Record<string, ContextValue>>;

export type DocxPagesErrorInit<Code extends string> = {
  readonly code: Code;
  readonly message: string;
  readonly at: string;
  readonly context: ErrorContext;
  readonly cause?: unknown;
};

const BRAND = "@docx-pages/core.DocxPagesError";

export class DocxPagesError<Code extends string = string> extends Error {
  readonly brand: typeof BRAND = BRAND;
  readonly code: Code;
  readonly at: string;
  readonly context: ErrorContext;

  constructor(init: DocxPagesErrorInit<Code>) {
    super(`[${init.at}] ${init.message}`, init.cause === undefined ? {} : { cause: init.cause });
    this.name = "DocxPagesError";
    this.code = init.code;
    this.at = init.at;
    this.context = init.context;
  }
}

// Branded rather than `instanceof`, which fails when a bundler or workspace link
// produces two copies of this module.
export function isDocxPagesError(value: unknown): value is DocxPagesError {
  return (
    typeof value === "object" &&
    value !== null &&
    "brand" in value &&
    value.brand === BRAND &&
    value instanceof Error
  );
}
