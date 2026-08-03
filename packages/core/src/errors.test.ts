import { describe, expect, it } from "vitest";

import { OnePagerError, isOnePagerError } from "./errors.js";

describe("OnePagerError", () => {
  it("carries the location and context on the instance", () => {
    const error = new OnePagerError({
      code: "backend-unavailable",
      message: "Microsoft Word is not installed",
      at: "render/backend/word-desktop.locateWord",
      context: { platform: "macos", searchedPath: "/Applications/Microsoft Word.app" },
    });

    expect(error.code).toBe("backend-unavailable");
    expect(error.at).toBe("render/backend/word-desktop.locateWord");
    expect(error.context).toStrictEqual({
      platform: "macos",
      searchedPath: "/Applications/Microsoft Word.app",
    });
  });

  it("puts the location in the message so a bare stack trace is still diagnosable", () => {
    const error = new OnePagerError({
      code: "manifest-invalid",
      message: "pages must not be empty",
      at: "core/manifest.parseManifest",
      context: { sourceFileName: "Reference.docx" },
    });

    expect(error.message).toBe("[core/manifest.parseManifest] pages must not be empty");
  });

  it("preserves the underlying cause", () => {
    const cause = new Error("ENOENT");
    const error = new OnePagerError({
      code: "render-failed",
      message: "could not read the converted PDF",
      at: "render/pipeline.readOutput",
      context: { outputPath: "/tmp/out.pdf" },
      cause,
    });

    expect(error.cause).toBe(cause);
  });

  it("is recognisable across module boundaries without instanceof", () => {
    const error = new OnePagerError({
      code: "render-failed",
      message: "boom",
      at: "render/pipeline.run",
      context: {},
    });

    expect(isOnePagerError(error)).toBe(true);
    expect(isOnePagerError(new Error("boom"))).toBe(false);
    expect(isOnePagerError(null)).toBe(false);
    expect(isOnePagerError({ code: "render-failed" })).toBe(false);
  });

  it("narrows to the specific code it was constructed with", () => {
    const error = new OnePagerError({
      code: "backend-unavailable",
      message: "no Word",
      at: "render/backend.select",
      context: {},
    });

    // The generic parameter is inferred from `code`, so this is a compile-time check
    // as much as a runtime one.
    const code: "backend-unavailable" = error.code;
    expect(code).toBe("backend-unavailable");
  });
});
