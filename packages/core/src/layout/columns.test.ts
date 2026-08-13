import { describe, expect, it } from "vitest";

import type { SectionColumns } from "../docx/section.js";
import { columnsAcross } from "./columns.js";

// Word's own answers, measured by the authored `columns` document. The frame there
// runs 540pt from 36, which is what every case below is read against.
const FRAME = { leftPt: 36, widthPt: 540 };

const asked = (columns: Partial<SectionColumns>): SectionColumns => ({
  count: 1,
  widthsTwips: [],
  gapsTwips: [],
  spaceTwips: 0,
  ...columns,
});

describe("columnsAcross", () => {
  it("gives a section of one column the whole frame", () => {
    expect(columnsAcross(asked({ count: 1 }), FRAME)).toStrictEqual([FRAME]);
  });

  it("takes the gaps off the frame and splits what is left", () => {
    expect(columnsAcross(asked({ count: 2, spaceTwips: 720 }), FRAME)).toStrictEqual([
      { leftPt: 36, widthPt: 252 },
      { leftPt: 324, widthPt: 252 },
    ]);
  });

  it("ends the last of them on the frame's own right edge", () => {
    const across = columnsAcross(asked({ count: 3, spaceTwips: 360 }), FRAME);
    expect(across.map((column) => column.leftPt)).toStrictEqual([36, 222, 408]);
    expect(across.map((column) => column.widthPt)).toStrictEqual([168, 168, 168]);
    const last = across[2];
    expect((last?.leftPt ?? 0) + (last?.widthPt ?? 0)).toBeCloseTo(576, 9);
  });

  it("takes a section stating its own widths at its word", () => {
    const across = columnsAcross(
      asked({ count: 2, widthsTwips: [3600, 6480], gapsTwips: [720, 0] }),
      FRAME,
    );
    expect(across).toStrictEqual([
      { leftPt: 36, widthPt: 180 },
      { leftPt: 252, widthPt: 324 },
    ]);
  });

  it("leaves the slack of a narrow set at the frame's right edge", () => {
    expect(
      columnsAcross(asked({ count: 2, widthsTwips: [2160, 2880], gapsTwips: [720, 0] }), FRAME),
    ).toStrictEqual([
      { leftPt: 36, widthPt: 108 },
      { leftPt: 180, widthPt: 144 },
    ]);
  });

  it("packs three narrow columns against the frame's left", () => {
    expect(
      columnsAcross(
        asked({ count: 3, widthsTwips: [1440, 2880, 1440], gapsTwips: [360, 360, 0] }),
        FRAME,
      ),
    ).toStrictEqual([
      { leftPt: 36, widthPt: 72 },
      { leftPt: 126, widthPt: 144 },
      { leftPt: 288, widthPt: 72 },
    ]);
  });

  it("draws a set wider than its frame past the frame's edge rather than clamping it", () => {
    expect(
      columnsAcross(asked({ count: 2, widthsTwips: [5760, 5760], gapsTwips: [720, 0] }), FRAME),
    ).toStrictEqual([
      { leftPt: 36, widthPt: 288 },
      { leftPt: 360, widthPt: 288 },
    ]);
  });
});
