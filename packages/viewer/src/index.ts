export { Document, Page } from "./page.js";
export type { FrameStyle, DocumentProps, PageProps } from "./page.js";

export { DocxDocument } from "./docx-document.js";
export type { DocxDocumentProps, DocxFont, DocxRenderReport } from "./docx-document.js";

export { imageDataUrl, imageResolver } from "./images.js";
export type { DrawableImage, ImageResolver } from "./images.js";

// The traversal is core's, since the pdf backend walks the very same one. Offered
// again here because a caller drawing a page by hand already has the viewer and
// should not have to reach past it for the order to draw things in.
export { drawablesOf } from "@docx-pages/core";
export type { Drawable } from "@docx-pages/core";
