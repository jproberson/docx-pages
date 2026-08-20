export { DocxPagesError, isDocxPagesError } from "@docx-pages/core";
export type { ErrorContext } from "@docx-pages/core";

export { readFillPlacements } from "./pdf/fills.js";
export type { FillPlacement } from "./pdf/fills.js";

export { readImagePlacements } from "./pdf/placements.js";
export type { ImagePlacement, PlacedRect } from "./pdf/placements.js";

// How many pages a pdf holds is `readDrawnText`'s to answer and cannot be counted
// off the placements, since a page drawing nothing but a picture holds no text at
// all.
export { readDrawnText, readTextPlacements } from "./pdf/text.js";
export type { DrawnText, TextPlacement } from "./pdf/text.js";
