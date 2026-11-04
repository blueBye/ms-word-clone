export type * from "./model/index.js";
export {
  createDocumentEngine,
  DocumentEngine,
  allParagraphs,
} from "./core/document.js";
export type {
  EngineSnapshot,
  EngineOptions,
  EngineStatistics,
} from "./core/document.js";
export type { DocumentSession, DocumentStyle } from "./core/types.js";
export type { StyleCatalog, StyleDefinition } from "./core/styles.js";
export { getParagraphs, getTextRuns } from "./model/index.js";
export { importDocx } from "./docx/import-docx.js";
export { openDocxPackage, findMainDocumentPart } from "./docx/package.js";
export type { DocxPackage, ImportOptions } from "./docx/package.js";
export { openDocxSession, createBlankDocx } from "./docx/session.js";
export type {
  DocxSession,
  DocxStyle,
  EditCommand,
  ParagraphCapability,
  RunFormat,
  ParagraphFormat,
} from "./docx/session.js";
