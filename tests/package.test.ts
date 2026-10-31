import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { findMainDocumentPart, openDocxPackage } from '../src/docx/package.js';
import { docx, relationships } from './helpers/docx.js';

describe('OPC package boundary', () => {
  it('reads binary/text entries and discovers a nonstandard main part', async () => {
    const pkg = await openDocxPackage(docx(undefined, { 'custom/media.bin': new Uint8Array([0, 255, 42]) }));
    expect(await findMainDocumentPart(pkg)).toBe('custom/main.xml');
    expect(await pkg.readBinary('custom/media.bin')).toEqual(new Uint8Array([0, 255, 42]));
    expect(await pkg.readText('custom/main.xml')).toContain('w:document');
    expect(pkg.listParts()).toHaveLength(4);
    await expect(pkg.readText('missing.xml')).rejects.toThrow('Missing DOCX part');
  });
  it('accepts a package-root absolute relationship target', async () => {
    const pkg = await openDocxPackage(docx(undefined, { '_rels/.rels': relationships('/custom/main.xml') }));
    expect(await findMainDocumentPart(pkg)).toBe('custom/main.xml');
  });
  it.each(['https://example.invalid/main.xml', '//example.invalid/main.xml', 'custom/main.xml?x=1', 'custom/main.xml#part', 'custom/%2e%2e/main.xml', 'custom\\main.xml', 'http%3Aevil.xml', '%zz'])('rejects unsafe main target %s', async target => {
    const pkg = await openDocxPackage(docx(undefined, { '_rels/.rels': relationships(target) }));
    await expect(findMainDocumentPart(pkg)).rejects.toThrow();
  });
  it('rejects external, missing and invalid main relationships', async () => {
    for (const xml of [relationships('custom/main.xml', 'External'), relationships('missing.xml'), relationships('', '', 'urn:other'), '<Relationships/>']) {
      const pkg = await openDocxPackage(docx(undefined, { '_rels/.rels': xml }));
      await expect(findMainDocumentPart(pkg)).rejects.toThrow();
    }
  });
  it.each(['../evil.xml', '/evil.xml', 'a/./evil.xml', 'a//evil.xml', 'a\\evil.xml'])('rejects unsafe ZIP entry %s', async path => {
    await expect(openDocxPackage(docx(undefined, { [path]: 'evil' }))).rejects.toThrow('Unsafe DOCX part name');
  });
  it.each([0, -1, NaN, Infinity, 1.5])('rejects invalid limits %s', async limit => {
    await expect(openDocxPackage(docx(), { maxUncompressedBytes: limit })).rejects.toThrow('positive safe integers');
    await expect(openDocxPackage(docx(), { maxEntries: limit })).rejects.toThrow('positive safe integers');
  });
  it('enforces input, uncompressed and entry-count budgets', async () => {
    await expect(openDocxPackage(docx(), { maxUncompressedBytes: 32 })).rejects.toThrow('input size limit');
    await expect(openDocxPackage(docx(undefined, { 'large.txt': 'a'.repeat(100_000) }), { maxUncompressedBytes: 4_096 })).rejects.toThrow('uncompressed size limit');
    await expect(openDocxPackage(docx(), { maxEntries: 2 })).rejects.toThrow('entry count limit');
  });
  it('checks the entire declared budget before attempting to inflate the first entry', async () => {
    const bytes = docx(undefined, { 'large.txt': 'a'.repeat(100_000) });
    // Corrupt the first entry's compression method in its central directory. Inflation
    // would fail with "unknown compression type" if it happened before preflight ended.
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    for (let i = 0; i < bytes.length - 46; i++) {
      if (view.getUint32(i, true) === 0x02014b50) { view.setUint16(i + 10, 99, true); break; }
    }
    await expect(openDocxPackage(bytes, { maxUncompressedBytes: 4_096 })).rejects.toThrow('uncompressed size limit');
  });
  it('rejects duplicate central-directory part names before extraction overwrites one', async () => {
    const bytes = docx(undefined, { 'duplicate-a': 'a', 'duplicate-b': 'b' });
    const needle = strToU8('duplicate-b');
    for (let i = 0; i <= bytes.length - needle.length; i++) {
      if (needle.every((byte, j) => bytes[i + j] === byte)) bytes[i + needle.length - 1] = 'a'.charCodeAt(0);
    }
    await expect(openDocxPackage(bytes)).rejects.toThrow('Duplicate DOCX part');
  });
  it('rejects archives missing mandatory OPC entries', async () => {
    await expect(openDocxPackage(zipSync({ 'document.xml': strToU8('<document/>') }))).rejects.toThrow('OPC');
  });
});
