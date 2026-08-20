import { describe, expect, it } from "vitest";

import { anchoredTextBox, bodySectionProperties, exactLine, exactLineClosedUp } from "./markup.js";

describe("anchoredTextBox", () => {
  // Word reads the size off the drawing and the shape reads it off the frame, so a
  // box whose two sizes differ is a box drawn at neither of them.
  it("states one size for the drawing and for the shape it holds", () => {
    const box = anchoredTextBox({
      id: 1,
      name: "a",
      widthPt: 300,
      heightPt: 100,
      wrap: `<wp:wrapNone/>`,
      content: "",
    });
    expect(box).toContain(`<wp:extent cx="3810000" cy="1270000"/>`);
    expect(box).toContain(`<a:ext cx="3810000" cy="1270000"/>`);
  });

  it("stands where it is put, and at the anchor where it is not", () => {
    const put = anchoredTextBox({
      id: 2,
      name: "b",
      widthPt: 10,
      heightPt: 10,
      wrap: `<wp:wrapNone/>`,
      leftPt: 36,
      offsetPt: 72,
      content: "",
    });
    expect(put).toContain(
      `<wp:positionH relativeFrom="column"><wp:posOffset>457200</wp:posOffset></wp:positionH>`,
    );
    expect(put).toContain(
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>914400</wp:posOffset></wp:positionV>`,
    );
  });
});

// Every authored document is written on the one page, so a section that states
// nothing states that page: a section stating anything else would be asking a
// second question.
describe("bodySectionProperties", () => {
  it("writes the page the whole suite uses where nothing is asked for", () => {
    expect(bodySectionProperties()).toBe(
      `<w:sectPr>` +
        `<w:pgSz w:w="12240" w:h="15840"/>` +
        `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>` +
        `</w:sectPr>`,
    );
  });

  it("changes only what it is asked to change", () => {
    expect(bodySectionProperties({ type: "continuous", leftTwips: 2880 })).toContain(
      `<w:type w:val="continuous"/>`,
    );
    expect(bodySectionProperties({ leftTwips: 2880 })).toContain(`w:left="2880"`);
    expect(bodySectionProperties({ leftTwips: 2880 })).toContain(`w:top="720"`);
  });
});

describe("exactLine", () => {
  it("asks for the height in twips, which is what Word states a line in", () => {
    expect(exactLine(24)).toBe(`<w:spacing w:line="480" w:lineRule="exact"/>`);
  });

  it("closes up what a style would add above and below where it is asked to", () => {
    expect(exactLineClosedUp(24)).toBe(
      `<w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/>`,
    );
  });
});
