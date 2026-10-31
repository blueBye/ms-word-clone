import { unzipSync, strFromU8 } from 'fflate';
import { parseXml, childElements } from './parser/xml.js';
import { NS } from './namespaces.js';

export interface DocxPackage {
  hasPart(path: string): boolean;
  readBinary(path: string): Promise<Uint8Array>;
  readText(path: string): Promise<string>;
  listParts(): string[];
}
export interface ImportOptions {
  /** Maximum uncompressed package bytes, default 50 MiB. */
  maxUncompressedBytes?: number;
}
function normalizedPath(path: string): string {
  const normalized = path.replace(/^\/+/, '').replace(/\\/g, '/');
  const chunks = normalized.split('/');
  if (!normalized || chunks.some(chunk => !chunk || chunk === '.' || chunk === '..')) throw new Error(`Unsafe DOCX part name: ${path}`);
  return normalized;
}
export async function openDocxPackage(bytes: ArrayBuffer | Uint8Array, options: ImportOptions = {}): Promise<DocxPackage> {
  const input = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (input.length < 4 || input[0] !== 0x50 || input[1] !== 0x4b) throw new Error('Input is not a ZIP/DOCX file');
  const max = options.maxUncompressedBytes ?? 50 * 1024 * 1024;
  // fflate unzipSync inflates the whole archive: enforce both input and output budgets.
  if (input.byteLength > max) throw new Error('DOCX exceeds input size limit');
  const entries = unzipSync(input);
  const parts = new Map<string, Uint8Array>();
  let size = 0;
  for (const [path, value] of Object.entries(entries)) {
    if (path.endsWith('/')) continue;
    const key = normalizedPath(path);
    size += value.byteLength;
    if (size > max) throw new Error('DOCX exceeds uncompressed size limit');
    parts.set(key, value);
  }
  if (!parts.has('[Content_Types].xml') || !parts.has('_rels/.rels')) throw new Error('Not a valid OPC DOCX package');
  return {
    hasPart: p => parts.has(normalizedPath(p)),
    readBinary: async p => {
      const value = parts.get(normalizedPath(p));
      if (!value) throw new Error(`Missing DOCX part: ${p}`);
      return value;
    },
    readText: async p => strFromU8(await (async () => {
      const value = parts.get(normalizedPath(p));
      if (!value) throw new Error(`Missing DOCX part: ${p}`);
      return value;
    })()),
    listParts: () => [...parts.keys()],
  };
}
/** Discover the main document part through the OPC package relationships. */
export async function findMainDocumentPart(pkg: DocxPackage): Promise<string> {
  const rels = parseXml(await pkg.readText('_rels/.rels'));
  const root = rels.documentElement;
  if (root.localName !== 'Relationships' || root.namespaceURI !== NS.relationships) throw new Error('Invalid package relationships');
  const candidate = childElements(root).find(el =>
    el.localName === 'Relationship' && el.namespaceURI === NS.relationships &&
    [NS.officeRelationships + '/officeDocument', NS.officeRelationshipsStrict + '/officeDocument'].includes(el.getAttribute('Type') || '')
  );
  if (!candidate) throw new Error('DOCX main-document relationship not found');
  if ((candidate.getAttribute('TargetMode') || '').toLowerCase() === 'external') throw new Error('External main-document target rejected');
  let target = candidate.getAttribute('Target');
  if (!target) throw new Error('DOCX main-document relationship has no target');
  // External URI schemes, fragment/query navigation, and parent traversal are not valid local main-part paths.
  if (/^[a-z][\w+.-]*:/i.test(target) || /[?#]/.test(target)) throw new Error('Invalid DOCX document target');
  target = decodeURIComponent(target);
  const path = normalizedPath(target);
  if (!pkg.hasPart(path)) throw new Error(`Missing DOCX main document: ${path}`);
  return path;
}
