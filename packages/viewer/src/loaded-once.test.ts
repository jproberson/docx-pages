import { describe, expect, it } from "vitest";

import { loadedOnce } from "./loaded-once.js";

describe("loadedOnce", () => {
  it("runs what it was given once, however many callers ask", async () => {
    let asked = 0;
    const load = loadedOnce(() => {
      asked += 1;
      return Promise.resolve(asked);
    });

    expect(await Promise.all([load(), load(), load()])).toStrictEqual([1, 1, 1]);
    expect(asked).toBe(1);
  });

  // The whole reason it is written out rather than being a `??=`: a component that
  // failed to fetch once would be handed the same rejection for the life of the
  // page, remount after remount, with no way to ask again.
  it("forgets a failure, so the next caller tries again", async () => {
    const answers = [Promise.reject(new Error("no network")), Promise.resolve("read")];
    const load = loadedOnce(() => answers.shift() ?? Promise.reject(new Error("asked too often")));

    await expect(load()).rejects.toThrow("no network");
    await expect(load()).resolves.toBe("read");
  });

  it("hands the same failure to everyone who was already waiting on it", async () => {
    let asked = 0;
    const load = loadedOnce(() => {
      asked += 1;
      return Promise.reject(new Error("no network"));
    });

    const both = [load(), load()];
    await expect(Promise.allSettled(both)).resolves.toMatchObject([
      { status: "rejected" },
      { status: "rejected" },
    ]);
    expect(asked).toBe(1);
  });
});
