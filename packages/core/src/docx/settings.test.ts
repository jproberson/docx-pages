import { describe, expect, it } from "vitest";

import { buildDocx, wordDocument, WORDPROCESSING_NS } from "../testing/build-docx.js";
import { openDocx } from "./package.js";
import { readDocumentSettings, roundsAnchorsToTwips, DEFAULT_SETTINGS } from "./settings.js";

const settingsOf = (inner: string | null) =>
  readDocumentSettings(
    openDocx(
      buildDocx({
        "word/document.xml": wordDocument("<w:p/>"),
        ...(inner === null
          ? {}
          : {
              "word/settings.xml": `<?xml version="1.0"?><w:settings xmlns:w="${WORDPROCESSING_NS}">${inner}</w:settings>`,
            }),
      }),
    ),
  );

const compat = (value: string) =>
  `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${value}"/></w:compat>`;

describe("readDocumentSettings", () => {
  it("reads the default tab stop", () => {
    expect(settingsOf(`<w:defaultTabStop w:val="709"/>`).defaultTabStopTwips).toBe(709);
  });

  it("falls back to Word's own default where the document states none", () => {
    expect(settingsOf("").defaultTabStopTwips).toBe(720);
  });

  it("reads the compatibility mode out of the settings that declare one", () => {
    expect(settingsOf(`<w:zoom w:percent="100"/>${compat("15")}`).compatibilityMode).toBe(15);
  });

  it("answers with nothing for a document whose compat states other things only", () => {
    const other = `<w:compat><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="x" w:val="1"/></w:compat>`;
    expect(settingsOf(other).compatibilityMode).toBeNull();
  });

  it("takes the whole of a package with no settings part as the defaults", () => {
    expect(settingsOf(null)).toStrictEqual(DEFAULT_SETTINGS);
  });
});

describe("roundsAnchorsToTwips", () => {
  it("rounds where the document declares no compatibility mode", () => {
    expect(roundsAnchorsToTwips(settingsOf(""))).toBe(true);
  });

  it("rounds where the document declares one Word laid out differently", () => {
    expect(roundsAnchorsToTwips(settingsOf(compat("14")))).toBe(true);
  });

  it("leaves an object where the flow put it for a document declaring 15", () => {
    expect(roundsAnchorsToTwips(settingsOf(compat("15")))).toBe(false);
  });
});
