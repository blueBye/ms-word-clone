import { describe, expect, it } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { importDocx } from '../src/docx/import-docx.js';
import { parseXml } from '../src/docx/parser/xml.js';
import { parseDocumentXml } from '../src/docx/parser/document.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function docx(xml: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
    '_rels/.rels': strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/main.xml"/></Relationships>'),
    'custom/main.xml': strToU8(xml),
  });
}
describe('Milestone 1 DOCX import', () => {
  it('imports RTL paragraph, basic formatting and complex-script overrides', async () => {
    const xml = `<w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:cs="B Zar"/><w:b w:val="0"/><w:bCs/><w:iCs w:val="false"/><w:sz w:val="22"/><w:szCs w:val="24"/><w:rtl/></w:rPr><w:t>قرارداد کار</w:t></w:r></w:p></w:body></w:document>`;
    const doc = await importDocx(docx(xml));
    expect(doc.blocks[0].properties).toEqual({ direction: 'rtl', alignment: 'right' });
    expect(doc.blocks[0].children[0]).toEqual({ type: 'text', text: 'قرارداد کار', properties: {
      direction: 'rtl', fontFamily: 'Times New Roman', fontSizePt: 11, bold: false,
      complexScript: { fontFamily: 'B Zar', fontSizePt: 12, bold: true, italic: false },
      italic: undefined,
    } });
  });
  it('preserves split runs, whitespace, tabs and breaks', async () => {
    const xml = `<x:document xmlns:x="${W}"><x:body><x:p><x:r><x:t xml:space="preserve"> A </x:t><x:tab/><x:t>B</x:t></x:r><x:r><x:t>C</x:t><x:br/></x:r></x:p><x:p/></x:body></x:document>`;
    const result = parseDocumentXml(parseXml(xml));
    expect(result.blocks).toHaveLength(2);
    expect(result.blocks[0].children.map(run => run.text)).toEqual([' A \tB', 'C\n']);
  });
  it('rejects unrelated XML document roots', () => {
    expect(() => parseDocumentXml(parseXml('<document/>'))).toThrow('WordprocessingML');
  });
  it('rejects invalid ZIP data', async () => {
    await expect(importDocx(strToU8('hello'))).rejects.toThrow('ZIP');
  });
});
