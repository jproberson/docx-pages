import { describe, expect, it } from "vitest";

import {
  aliasedSymbolCharacter,
  aliasedSymbolText,
  isAliasedSymbolFace,
} from "./symbol-aliases.js";

describe("symbol aliases", () => {
  it("answers for a position bare or lifted onto the private-use page alike", () => {
    expect(aliasedSymbolCharacter("Wingdings", 0xa7)).toBe("▪");
    expect(aliasedSymbolCharacter("Wingdings", 0xf0a7)).toBe("▪");
    expect(aliasedSymbolCharacter("Symbol", 0x61)).toBe("α");
    expect(aliasedSymbolCharacter("Symbol", 0xf061)).toBe("α");
  });

  it("says what a whole run means, and boxes what the tables do not carry", () => {
    expect(aliasedSymbolText("Wingdings", "ü ý")).toBe("✓ ☒");
    // 0x71 is a position the table does not carry: bare, it would be painted as
    // the letter q, so it is lifted to the private-use page and drawn as a box.
    expect(aliasedSymbolText("Wingdings", "q")).toBe(String.fromCodePoint(0xf071));
  });

  it("answers only for the faces it knows", () => {
    expect(isAliasedSymbolFace("Wingdings")).toBe(true);
    expect(isAliasedSymbolFace("wingdings ")).toBe(true);
    expect(isAliasedSymbolFace("Webdings")).toBe(false);
    expect(aliasedSymbolText("Calibri", "l")).toBeNull();
  });
});
