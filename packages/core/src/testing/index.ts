export { buildDocx, wordDocument, WORDPROCESSING_NS } from "./build-docx.js";
export type { DocxPart } from "./build-docx.js";

export { buildFace, buildSfnt, buildWoff, buildWoff2 } from "./build-font.js";
export type { FaceFixture, FontFixture } from "./build-font.js";

export {
  buildMetafile,
  metafileFont,
  metafileHeader,
  metafileRecord,
  metafileText,
  utf16,
} from "./build-metafile.js";
export type { MetafileFontFixture, MetafileFrame, MetafileTextFixture } from "./build-metafile.js";
