import { describe, expect, it } from 'vitest';
import { strToU8 } from 'fflate';
import { importDocx } from '../src/docx/import-docx.js';
import { getTextRuns } from '../src/model/traversal.js';
import { parseXml } from '../src/docx/parser/xml.js';
import { parseDocumentXml } from '../src/docx/parser/document.js';
import {
  docx,
  documentXml,
  relationships,
  W,
  W_STRICT,
} from './helpers/docx.js';

describe('Milestone 1 DOCX import', () => {
  it('imports RTL paragraph, basic formatting and complex-script overrides', async () => {
    const xml = `<w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:cs="B Zar"/><w:b w:val="0"/><w:bCs/><w:iCs w:val="false"/><w:sz w:val="22"/><w:szCs w:val="24"/><w:rtl/></w:rPr><w:t>قرارداد کار</w:t></w:r></w:p></w:body></w:document>`;
    const doc = await importDocx(docx(xml));
    expect(doc.blocks[0].properties).toEqual({
      direction: 'rtl',
      alignment: 'right',
    });
    expect(getTextRuns(doc.blocks)[0]).toMatchObject({
      type: 'text',
      text: 'قرارداد کار',
      properties: {
        direction: 'rtl',
        fontFamily: 'Times New Roman',
        fontSizePt: 11,
        bold: false,
        complexScript: {
          fontFamily: 'B Zar',
          fontSizePt: 12,
          bold: true,
          italic: false,
        },
      },
    });
  });
  it('preserves split runs, whitespace, tabs and breaks', async () => {
    const xml = `<x:document xmlns:x="${W}"><x:body><x:p><x:r><x:t xml:space="preserve"> A </x:t><x:tab/><x:t>B</x:t></x:r><x:r><x:t>C</x:t><x:br/></x:r></x:p><x:p/></x:body></x:document>`;
    const result = parseDocumentXml(parseXml(xml));
    expect(result.blocks).toHaveLength(2);
    expect(getTextRuns(result.blocks).map((run) => run.text)).toEqual([
      ' A \tB',
      'C\n',
    ]);
  });
  it('rejects unrelated XML document roots', () => {
    expect(() => parseDocumentXml(parseXml('<document/>'))).toThrow(
      'WordprocessingML',
    );
  });
  it('rejects invalid ZIP data', async () => {
    await expect(importDocx(strToU8('hello'))).rejects.toThrow('ZIP');
  });
  it('imports Strict namespaces and Strict office relationships', async () => {
    const xml = documentXml(
      '<w:p><w:pPr><w:bidi w:val="0"/><w:jc w:val="end"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="23"/></w:rPr><w:t>Strict</w:t></w:r></w:p>',
      W_STRICT,
    );
    const result = await importDocx(
      docx(xml, {
        '_rels/.rels': relationships(
          '/custom/main.xml',
          '',
          'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument',
        ),
      }),
    );
    expect(result.blocks[0].properties).toEqual({
      direction: 'ltr',
      alignment: 'end',
    });
    expect(getTextRuns(result.blocks)[0].properties).toMatchObject({
      bold: true,
      fontSizePt: 11.5,
    });
  });
  it('keeps Unicode, ZWNJ, combining marks and digit forms without normalization', async () => {
    const text = 'می\u200cروم عَرَبِيّ עברית English ۱۲۳ ١٢٣ 123 (a/b) \u200f';
    const result = await importDocx(
      docx(documentXml(`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)),
    );
    expect(getTextRuns(result.blocks)[0].text).toBe(text);
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });
  it('traverses inline content without leaking SDT metadata or deleted text', () => {
    const result = parseDocumentXml(
      parseXml(
        documentXml(`<w:p>
      <w:hyperlink><w:r><w:t>link</w:t></w:r></w:hyperlink>
      <w:sdt><w:sdtPr><w:r><w:t>metadata</w:t></w:r></w:sdtPr><w:r><w:t>outside-content</w:t></w:r><w:sdtContent><w:ins><w:r><w:t>content</w:t></w:r></w:ins></w:sdtContent></w:sdt>
      <w:del><w:r><w:delText>deleted</w:delText></w:r></w:del>
      <w:r><w:instrText>PAGE</w:instrText><w:t>1</w:t><w:noBreakHyphen/><w:softHyphen/></w:r>
    </w:p>`),
      ),
    );
    expect(getTextRuns(result.blocks).map((run) => run.text)).toEqual([
      'link',
      'content',
      '1\u2011\u00ad',
    ]);
  });
  it('rejects a missing body and does not treat foreign elements as Word text', () => {
    expect(() =>
      parseDocumentXml(parseXml(`<w:document xmlns:w="${W}"/>`)),
    ).toThrow('w:body');
    const result = parseDocumentXml(
      parseXml(
        documentXml(
          '<p><r><t>foreign</t></r></p><w:p><w:r><t>foreign</t><w:t>word</w:t></w:r></w:p>',
        ),
      ),
    );
    expect(result.blocks).toHaveLength(1);
    expect(getTextRuns(result.blocks)[0].text).toBe('word');
  });
});
