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
  /** Maximum compressed input and declared/actual uncompressed bytes, default 50 MiB. */
  maxUncompressedBytes?: number;
  /** Maximum ZIP entries (including directories), default 2,000. */
  maxEntries?: number;
}
function normalizedPath(path: string): string {
  const chunks = path.split('/');
  if (!path || /[\\\u0000-\u001f]/.test(path) || chunks.some(chunk => !chunk || chunk === '.' || chunk === '..')) throw new Error(`Unsafe DOCX part name: ${path}`);
  return path;
}
export async function openDocxPackage(bytes: ArrayBuffer | Uint8Array, options: ImportOptions = {}): Promise<DocxPackage> {
  const input = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (input.length < 4 || input[0] !== 0x50 || input[1] !== 0x4b) throw new Error('Input is not a ZIP/DOCX file');
  const max = options.maxUncompressedBytes ?? 50 * 1024 * 1024;
  const maxEntries = options.maxEntries ?? 2_000;
  if (!Number.isSafeInteger(max) || max <= 0 || !Number.isSafeInteger(maxEntries) || maxEntries <= 0) {
    throw new Error('DOCX size and entry limits must be positive safe integers');
  }
  if (input.byteLength > max) throw new Error('DOCX exceeds input size limit');
  // Scan the central directory without inflating anything, before allocating output buffers.
  const declaredSizes = new Map<string, number>();
  let declaredBytes = 0;
  let entryCount = 0;
  unzipSync(input, { filter: entry => {
    if (++entryCount > maxEntries) throw new Error('DOCX exceeds entry count limit');
    const key = normalizedPath(entry.name.endsWith('/') ? entry.name.slice(0, -1) : entry.name);
    if (declaredSizes.has(key)) throw new Error(`Duplicate DOCX part: ${key}`);
    declaredSizes.set(key, entry.originalSize);
    declaredBytes += Math.max(entry.originalSize, entry.size);
    if (!Number.isSafeInteger(declaredBytes) || declaredBytes > max) throw new Error('DOCX exceeds uncompressed size limit');
    return false;
  } });
  const entries = unzipSync(input);
  const parts = new Map<string, Uint8Array>();
  let size = 0;
  for (const [path, value] of Object.entries(entries)) {
    if (path.endsWith('/')) continue;
    const key = normalizedPath(path);
    if (value.byteLength !== declaredSizes.get(key)) throw new Error(`DOCX part size mismatch: ${key}`);
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
    readText: async p => {
      const value = parts.get(normalizedPath(p));
      if (!value) throw new Error(`Missing DOCX part: ${p}`);
      return strFromU8(value);
    },
    listParts: () => [...parts.keys()],
  };
}
/** Discover the main document part through the OPC package relationships. */
export async function findMainDocumentPart(pkg: DocxPackage): Promise<string> {
  const rels = parseXml(await pkg.readText('_rels/.rels'));
  const root = rels.documentElement;
  if (!root || root.localName !== 'Relationships' || root.namespaceURI !== NS.relationships) throw new Error('Invalid package relationships');
  const candidate = childElements(root).find(el =>
    el.localName === 'Relationship' && el.namespaceURI === NS.relationships &&
    [NS.officeRelationships + '/officeDocument', NS.officeRelationshipsStrict + '/officeDocument'].includes(el.getAttribute('Type') || '')
  );
  if (!candidate) throw new Error('DOCX main-document relationship not found');
  if ((candidate.getAttribute('TargetMode') || '').toLowerCase() === 'external') throw new Error('External main-document target rejected');
  let target = candidate.getAttribute('Target');
  if (!target) throw new Error('DOCX main-document relationship has no target');
  // External URI schemes, fragment/query navigation, and parent traversal are not valid local main-part paths.
  try { target = decodeURIComponent(target); }
  catch { throw new Error('Invalid DOCX document target encoding'); }
  if (/^[a-z][\w+.-]*:/i.test(target) || /[?#]/.test(target) || target.startsWith('//')) throw new Error('Invalid DOCX document target');
  const path = normalizedPath(target.startsWith('/') ? target.slice(1) : target);
  if (!pkg.hasPart(path)) throw new Error(`Missing DOCX main document: ${path}`);
  return path;
}
