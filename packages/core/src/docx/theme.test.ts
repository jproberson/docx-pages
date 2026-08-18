import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument, WORDPROCESSING_NS } from "../testing/build-docx.js";
import { openDocx } from "./package.js";
import { A_NS } from "./styles.js";
import { readColorReference, readTheme, themeColor, NO_THEME } from "./theme.js";
import { parseXml, type XmlElement } from "./xml.js";

const themeXml = (scheme: string) => `<?xml version="1.0"?>
<a:theme xmlns:a="${A_NS}"><a:themeElements><a:clrScheme name="Test">${scheme}</a:clrScheme>
</a:themeElements></a:theme>`;

const SCHEME = `<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
  <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
  <a:accent1><a:srgbClr val="4472C4"/></a:accent1>`;

const MAPPING = `<w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:accent1="accent1"/>`;

const settings = (inner: string) =>
  `<?xml version="1.0"?><w:settings xmlns:w="${WORDPROCESSING_NS}">${inner}</w:settings>`;

const themeOf = (scheme: string = SCHEME, mapping: string = MAPPING) =>
  readTheme(
    openDocx(
      buildDocx({
        "word/document.xml": wordDocument("<w:p/>"),
        "word/theme/theme1.xml": themeXml(scheme),
        "word/settings.xml": settings(mapping),
      }),
    ),
  );

function fill(inner: string): XmlElement {
  const root = parseXml(`<a:solidFill xmlns:a="${A_NS}">${inner}</a:solidFill>`);
  if (root === null) throw new Error("expected a fill");
  return root;
}

const colorOf = (inner: string) => {
  const reference = readColorReference(fill(inner));
  if (reference === null) throw new Error("expected a colour");
  return themeColor(themeOf(), reference);
};

describe("readColorReference", () => {
  it("reads a literal colour and leaves its luminance alone", () => {
    expect(readColorReference(fill(`<a:srgbClr val="4472C4"/>`))).toStrictEqual({
      base: { kind: "literal", hex: "4472C4" },
      luminanceScale: 1,
      luminanceOffset: 0,
    });
  });

  it("reads the luminance transform written under the colour", () => {
    expect(
      readColorReference(fill(`<a:schemeClr val="bg1"><a:lumMod val="95000"/></a:schemeClr>`)),
    ).toStrictEqual({
      base: { kind: "scheme", slot: "bg1" },
      luminanceScale: 0.95,
      luminanceOffset: 0,
    });
  });

  // A system colour is whatever the machine that wrote the file resolved it to,
  // which is the only reading available anywhere else.
  it("takes a system colour at the value the producer last saw", () => {
    expect(
      readColorReference(fill(`<a:sysClr val="window" lastClr="FFFFFF"/>`))?.base,
    ).toStrictEqual({ kind: "literal", hex: "FFFFFF" });
  });

  it("reads nothing from a fill that names no colour it can resolve", () => {
    expect(readColorReference(fill(`<a:gradFill/>`))).toBeNull();
  });

  // A colour stated at no opacity at all is one Word puts no ink down for, and a
  // corpus document writes its full-width rules exactly this way.
  it("reads nothing from a colour stated fully transparent", () => {
    expect(
      readColorReference(fill(`<a:srgbClr val="000000"><a:alpha val="0"/></a:srgbClr>`)),
    ).toBeNull();
    expect(
      readColorReference(fill(`<a:schemeClr val="bg1"><a:alpha val="0"/></a:schemeClr>`)),
    ).toBeNull();
  });

  // Only nought is answered for. Two colours in the whole corpus stand part way,
  // and drawing them opaque is what this did before and keeps doing.
  it("draws every other alpha opaque, the full one included", () => {
    expect(
      readColorReference(fill(`<a:srgbClr val="4472C4"><a:alpha val="100000"/></a:srgbClr>`)),
    ).toStrictEqual({
      base: { kind: "literal", hex: "4472C4" },
      luminanceScale: 1,
      luminanceOffset: 0,
    });
    expect(
      readColorReference(fill(`<a:srgbClr val="4472C4"><a:alpha val="40000"/></a:srgbClr>`))?.base,
    ).toStrictEqual({ kind: "literal", hex: "4472C4" });
  });
});

describe("themeColor", () => {
  // Word draws both of these greys, measured off its own pdf: bg1 maps to the
  // theme's white, and lumMod scales that white's luminance.
  it("scales a mapped scheme colour's luminance", () => {
    expect(colorOf(`<a:schemeClr val="bg1"><a:lumMod val="95000"/></a:schemeClr>`)).toBe("#F2F2F2");
    expect(colorOf(`<a:schemeClr val="bg1"><a:lumMod val="75000"/></a:schemeClr>`)).toBe("#BFBFBF");
  });

  it("scales a literal colour's luminance too", () => {
    expect(colorOf(`<a:srgbClr val="FFFFFF"><a:lumMod val="50000"/></a:srgbClr>`)).toBe("#808080");
  });

  // The transforms Word offers on a scheme colour, against the swatches its own
  // colour picker shows for this accent: lighter by 40, 60 and 80 per cent, then
  // darker by 25 and 50.
  it("keeps a colour's hue while its luminance moves", () => {
    const tinted = (transform: string) =>
      colorOf(`<a:schemeClr val="accent1">${transform}</a:schemeClr>`);

    expect(colorOf(`<a:schemeClr val="accent1"/>`)).toBe("#4472C4");
    expect(tinted(`<a:lumMod val="60000"/><a:lumOff val="40000"/>`)).toBe("#8FAADC");
    expect(tinted(`<a:lumMod val="40000"/><a:lumOff val="60000"/>`)).toBe("#B4C7E7");
    expect(tinted(`<a:lumMod val="20000"/><a:lumOff val="80000"/>`)).toBe("#DAE3F3");
    expect(tinted(`<a:lumMod val="75000"/>`)).toBe("#2F5597");
    expect(tinted(`<a:lumMod val="50000"/>`)).toBe("#203864");
  });

  it("resolves a slot the theme names directly, without the mapping", () => {
    expect(colorOf(`<a:schemeClr val="lt1"/>`)).toBe("#FFFFFF");
  });

  it("answers for nothing when the theme has no such slot", () => {
    const reference = readColorReference(fill(`<a:schemeClr val="bg1"/>`));
    if (reference === null) throw new Error("expected a colour");
    expect(themeColor(NO_THEME, reference)).toBeNull();
  });
});
