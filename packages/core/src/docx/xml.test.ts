import { describe, expect, it } from "vitest";

import { attribute, parseXml } from "./xml.js";

const parse = (source: string) => {
  const root = parseXml(source);
  if (root === null) throw new Error("expected a root element");
  return root;
};

describe("parseXml", () => {
  it("decodes a numeric character reference in text", () => {
    expect(parse(`<t>a&#233;b</t>`).text).toBe("a\u00e9b");
  });

  it("decodes a hexadecimal character reference in an attribute", () => {
    expect(attribute(parse(`<lvl val="&#xF0A7;"/>`), "", "val")).toBe("\uF0A7");
  });

  it("decodes the named entities xml defines", () => {
    expect(parse(`<t>a&amp;b&lt;c&gt;d&quot;e&apos;f</t>`).text).toBe(`a&b<c>d"e'f`);
  });

  it("leaves an ampersand that opens no entity alone", () => {
    expect(parse(`<t>a &#zz; b</t>`).text).toBe("a &#zz; b");
  });
});
