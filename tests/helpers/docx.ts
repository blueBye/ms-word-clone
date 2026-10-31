import { zipSync, strToU8 } from 'fflate';

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const W_STRICT = 'http://purl.oclc.org/ooxml/wordprocessingml/main';
export const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function documentXml(body: string, ns = W): string {
  return `<w:document xmlns:w="${ns}"><w:body>${body}</w:body></w:document>`;
}

export function relationships(target = 'custom/main.xml', mode = '', type = `${OFFICE_REL}/officeDocument`): string {
  return `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${type}" Target="${target}"${mode ? ` TargetMode="${mode}"` : ''}/></Relationships>`;
}

/** Synthetic OPC fixtures; Word-authored interoperability files remain a separate milestone. */
export function docx(xml = documentXml('<w:p/>'), overrides: Record<string, string | Uint8Array> = {}): Uint8Array {
  const parts: Record<string, string | Uint8Array> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/custom/main.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    '_rels/.rels': relationships(),
    'custom/main.xml': xml,
    ...overrides,
  };
  return zipSync(Object.fromEntries(Object.entries(parts).map(([name, data]) => [name, typeof data === 'string' ? strToU8(data) : data])));
}
