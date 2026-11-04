/** Microsoft Word/OOXML compatibility layer. The portable engine lives in ../core. */
export { importDocx } from "./import-docx.js";
export { openDocxPackage, findMainDocumentPart } from "./package.js";
export type { DocxPackage, ImportOptions } from "./package.js";
export { openDocxSession, createBlankDocx } from "./session.js";
export type { DocxSession, DocxStyle } from "./session.js";
