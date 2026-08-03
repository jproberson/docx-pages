export type ContextValue = string | number | boolean | null | readonly ContextValue[];

export type ErrorContext = Readonly<Record<string, ContextValue>>;

export type OnePagerErrorInit<Code extends string> = {
  readonly code: Code;
  readonly message: string;
  readonly at: string;
  readonly context: ErrorContext;
  readonly cause?: unknown;
};

const BRAND = "@onepager/core.OnePagerError";

export class OnePagerError<Code extends string = string> extends Error {
  readonly brand: typeof BRAND = BRAND;
  readonly code: Code;
  readonly at: string;
  readonly context: ErrorContext;

  constructor(init: OnePagerErrorInit<Code>) {
    super(`[${init.at}] ${init.message}`, init.cause === undefined ? {} : { cause: init.cause });
    this.name = "OnePagerError";
    this.code = init.code;
    this.at = init.at;
    this.context = init.context;
  }
}

// Branded rather than `instanceof`, which fails when a bundler or workspace link
// produces two copies of this module.
export function isOnePagerError(value: unknown): value is OnePagerError {
  return (
    typeof value === "object" &&
    value !== null &&
    "brand" in value &&
    value.brand === BRAND &&
    value instanceof Error
  );
}
