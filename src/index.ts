export type * from './model/index.js';
export { getParagraphs, getTextRuns } from './model/index.js';
export { importDocx } from './docx/import-docx.js';
export { openDocxPackage, findMainDocumentPart } from './docx/package.js';
export type { DocxPackage, ImportOptions } from './docx/package.js';
