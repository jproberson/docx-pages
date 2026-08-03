export { OnePagerError, isOnePagerError } from "./errors.js";
export type { ContextValue, ErrorContext, OnePagerErrorInit } from "./errors.js";

export { openDocx, partText, partXml, MAIN_DOCUMENT_PART } from "./docx/package.js";
export type { DocxPackage } from "./docx/package.js";

export { readSectionGeometry, W_NS } from "./docx/section.js";
export type { PageMargin, SectionGeometry } from "./docx/section.js";

export { ascentPt, lineHeightPt, lookupFontMetrics } from "./layout/font-metrics.js";
export type { FontMetrics, MetricsLookup } from "./layout/font-metrics.js";

export { readFontMetrics } from "./layout/font-file.js";
export type { FontFileFormat, ReadFontMetricsResult } from "./layout/font-file.js";
