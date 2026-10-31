import type { DocumentIR } from '../model/document.js';
import { findMainDocumentPart, openDocxPackage, type ImportOptions } from './package.js';
import { parseXml } from './parser/xml.js';
import { parseDocumentXml } from './parser/document.js';
/** Imports the supported Milestone 1 subset; not a full fidelity DOCX round-trip. */
export async function importDocx(input: ArrayBuffer | Uint8Array, options: ImportOptions = {}): Promise<DocumentIR> {
  const pkg = await openDocxPackage(input, options);
  const path = await findMainDocumentPart(pkg);
  return parseDocumentXml(parseXml(await pkg.readText(path)));
}
