import { describe, expect, it } from "vitest";

import { spelledAsMath } from "./math-letters.js";

// Every case is what Word drew, read straight out of its own pdf: one string spelled
// six times over, once for each style, and the pdf's own mapping back to characters.
function points(text: string): readonly string[] {
  const found: string[] = [];
  for (const character of text) {
    found.push((character.codePointAt(0) ?? 0).toString(16).padStart(4, "0"));
  }
  return found;
}

describe("spelledAsMath", () => {
  it("leaves a plain run as it is written", () => {
    expect(spelledAsMath("bandril h QRS 1234 αβΓ", "plain")).toBe("bandril h QRS 1234 αβΓ");
  });

  it("leaves a run stating m:nor as it is written", () => {
    expect(spelledAsMath("bandril", null)).toBe("bandril");
  });

  it("spells an italic run in the italic alphabet", () => {
    expect(points(spelledAsMath("bandril", "italic"))).toEqual([
      "1d44f",
      "1d44e",
      "1d45b",
      "1d451",
      "1d45f",
      "1d456",
      "1d459",
    ]);
  });

  // U+1D455 is reserved, and Word draws the Planck constant in its place.
  it("spells an italic h as U+210E", () => {
    expect(points(spelledAsMath("h", "italic"))).toEqual(["210e"]);
    expect(points(spelledAsMath("h", "bold"))).toEqual(["1d421"]);
    expect(points(spelledAsMath("h", "bold-italic"))).toEqual(["1d489"]);
  });

  it("leaves an italic run's digits alone and moves a bold run's", () => {
    expect(points(spelledAsMath("1234", "italic"))).toEqual(["0031", "0032", "0033", "0034"]);
    expect(points(spelledAsMath("1234", "bold"))).toEqual(["1d7cf", "1d7d0", "1d7d1", "1d7d2"]);
  });

  // Word drew a bold italic run's digits in the bold alphabet, there being no bold
  // italic one to draw them in.
  it("spells a bold italic run's digits as the bold ones", () => {
    expect(points(spelledAsMath("1234", "bold-italic"))).toEqual([
      "1d7cf",
      "1d7d0",
      "1d7d1",
      "1d7d2",
    ]);
  });

  it("spells capitals in each alphabet", () => {
    expect(points(spelledAsMath("QRS", "italic"))).toEqual(["1d444", "1d445", "1d446"]);
    expect(points(spelledAsMath("QRS", "bold"))).toEqual(["1d410", "1d411", "1d412"]);
    expect(points(spelledAsMath("QRS", "bold-italic"))).toEqual(["1d478", "1d479", "1d47a"]);
  });

  it("spells Greek in each alphabet, capitals and letters in one sequence", () => {
    expect(points(spelledAsMath("αβΓ", "italic"))).toEqual(["1d6fc", "1d6fd", "1d6e4"]);
    expect(points(spelledAsMath("αβΓ", "bold"))).toEqual(["1d6c2", "1d6c3", "1d6aa"]);
    expect(points(spelledAsMath("αβΓ", "bold-italic"))).toEqual(["1d736", "1d737", "1d71e"]);
  });

  it("leaves alone a character no alphabet holds", () => {
    expect(spelledAsMath(" +=(", "italic")).toBe(" +=(");
  });
});
