import { describe, expect, it } from 'vitest';
import { renderToStaticMarkup } from 'react-dom/server';
import { DocumentRenderer, getRequiredFonts } from '../src/react/index.js';
import {
  getParagraphs,
  getTextRuns,
  importDocx,
  type Block,
  type DocumentIR,
  type Paragraph,
  type Table,
} from '../src/index.js';
import {
  paginateBlocks,
  pageVariant,
  splitPageBreaks,
} from '../src/react/pagination.js';
import { localizeDigits } from '../src/model/numerals.js';
import { docx, documentXml, W, OFFICE_REL, REL } from './helpers/docx.js';

const text = (value: string, properties = '') =>
  `<w:r><w:rPr>${properties}</w:rPr><w:t xml:space="preserve">${value}</w:t></w:r>`;
const rels = (body: string) =>
  `<Relationships xmlns="${REL}">${body}</Relationships>`;
const relationship = (id: string, type: string, target: string) =>
  `<Relationship Id="${id}" Type="${OFFICE_REL}/${type}" Target="${target}"/>`;
const paragraph = (
  name: string,
  properties: Paragraph['properties'] = {},
): Paragraph => ({
  type: 'paragraph',
  properties,
  children: [{ type: 'text', text: name, properties: {} }],
});
const document = (blocks: Block[]): DocumentIR => ({
  type: 'document',
  schemaVersion: 2,
  blocks,
});
const render = (document: DocumentIR, number = 1, count = 1) =>
  renderToStaticMarkup(
    DocumentRenderer({
      document,
      pageLayout: true,
      pageNumber: number,
      pageCount: count,
    }),
  );

describe('opacity and bidi regressions', () => {
  it.each([
    ['0', 0],
    ['20000', 0.2],
    ['100000', 1],
  ])('applies alphaModFix amt=%s to the image', async (amount, expected) => {
    const body = `<w:p><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${OFFICE_REL}"><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:blip r:embed="image"><a:clrChange><a:clrTo><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange><a:alphaModFix amt="${amount}"/></a:blip></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const doc = await importDocx(
      docx(documentXml(body), {
        'custom/_rels/main.xml.rels': rels(
          relationship('image', 'image', 'pixel.png'),
        ),
        'custom/pixel.png': new Uint8Array([137, 80, 78, 71]),
      }),
    );
    const inline = getParagraphs(doc.blocks)[0].children[0];
    expect(inline.type).toBe('drawing');
    if (inline.type !== 'drawing') throw new Error('Expected drawing');
    expect(inline.image?.opacity).toBe(expected);
    expect(render(doc)).toContain(`opacity:${expected}`);
  });
  it('keeps mixed runs in one paragraph bidi context instead of embedding each run', async () => {
    const body = `<w:p><w:pPr><w:bidi/></w:pPr>${text('متن فارسی ', '<w:rtl/>')}${text('English words', '<w:rtl w:val="0"/>')}${text(' ادامه (۱۲۳)', '<w:rtl/>')}</w:p>`;
    const doc = await importDocx(docx(documentXml(body)));
    const markup = render(doc);
    expect(markup).toContain('<p dir="rtl"');
    expect(markup).toContain('unicode-bidi:isolate');
    expect(markup).not.toContain('unicode-bidi:embed');
    expect(markup).not.toMatch(/<span[^>]*dir=/);
    expect(
      getTextRuns(doc.blocks)
        .map((run) => run.text)
        .join(''),
    ).toBe('متن فارسی English words ادامه (۱۲۳)');
  });
  it('maps legacy left/right indents to paragraph start/end and keeps RTL markers inside the cell', async () => {
    const body = `<w:tbl><w:tblPr><w:bidiVisual/></w:tblPr><w:tr><w:tc><w:p><w:pPr><w:bidi/><w:ind w:left="384" w:hanging="360"/><w:numPr><w:numId w:val="1"/></w:numPr><w:rPr><w:rFonts w:cs="B Nazanin"/><w:szCs w:val="26"/></w:rPr></w:pPr>${text('گزارش فارسی')}</w:p></w:tc></w:tr></w:tbl>`;
    const doc = await importDocx(
      docx(documentXml(body), {
        'custom/_rels/main.xml.rels': rels(
          relationship('num', 'numbering', 'numbering.xml'),
        ),
        'custom/numbering.xml': `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
      }),
    );
    const p = getParagraphs(doc.blocks)[0];
    expect(p.properties.list).toMatchObject({
      marker: '۱.',
      locale: 'fa-IR',
      properties: {
        complexScript: { fontFamily: 'B Nazanin', fontSizePt: 13 },
      },
    });
    const markup = render(doc);
    expect(markup).toContain('margin-inline-start:19.2pt');
    expect(markup).not.toContain('margin-left:19.2pt');
    expect(markup).toContain('inset-inline-start:-18pt');
    expect(markup).toContain('۱.');
  });
  it('localizes marker digits without touching Latin body text', () => {
    expect(localizeDigits('12.3', 'fa-IR')).toBe('۱۲.۳');
    expect(localizeDigits('12.3', 'en-US')).toBe('12.3');
    expect(localizeDigits('12', 'ar-SA')).toBe('١٢');
  });
});

describe('header/footer fonts and page variants', () => {
  it('includes marker-only fonts when checking document font availability', () => {
    const doc = document([
      paragraph('Body', {
        list: {
          marker: '•',
          suffix: 'tab',
          properties: {
            fontSlots: { ascii: 'ASCII face', hAnsi: 'Marker face' },
          },
        },
      }),
    ]);
    expect(getRequiredFonts(doc)).toContainEqual({
      family: 'Marker face',
      locations: ['body'],
    });
  });
  it('reports a separate header textbox font and applies it to its text', async () => {
    const header = `<w:hdr xmlns:w="${W}"><w:p><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wp:inline><wp:extent cx="914400" cy="914400"/><wps:wsp><wps:txbx><w:txbxContent><w:p><w:pPr><w:bidi/></w:pPr>${text('عنوان فارسی', '<w:rFonts w:cs="B Compset"/><w:rtl/>')}</w:p></w:txbxContent></wps:txbx></wps:wsp></wp:inline></w:drawing></w:r></w:p></w:hdr>`;
    const doc = await importDocx(
      docx(
        documentXml(
          `<w:p>${text('Body')}</w:p><w:sectPr xmlns:r="${OFFICE_REL}"><w:headerReference w:type="default" r:id="header"/></w:sectPr>`,
        ),
        {
          'custom/_rels/main.xml.rels': rels(
            relationship('header', 'header', 'header.xml'),
          ),
          'custom/header.xml': header,
        },
      ),
    );
    expect(getRequiredFonts(doc)).toContainEqual({
      family: 'B Compset',
      locations: ['header'],
    });
    expect(render(doc)).toContain('font-family:&quot;B Compset&quot;');
  });
  it('uses first/default/even stories independently and leaves an absent first footer blank', () => {
    const doc = document([paragraph('Body')]);
    doc.page = { differentFirstPage: true, evenAndOddHeaders: true };
    doc.headers = {
      first: [paragraph('First header')],
      default: [paragraph('Default header')],
      even: [paragraph('Even header')],
    };
    doc.footers = { default: [paragraph('Default footer')] };
    expect(render(doc, 1, 3)).toContain('First header');
    expect(render(doc, 1, 3)).not.toContain('Default footer');
    expect(render(doc, 2, 3)).toContain('Even header');
    expect(render(doc, 3, 3)).toContain('Default footer');
    expect(pageVariant(doc, 2)).toBe('even');
  });
  it('updates basic PAGE/NUMPAGES results per sheet and retains cached source text', async () => {
    const field = (instruction: string, cached: string) =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>${instruction}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${text(cached, '<w:rtl/>')}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const doc = await importDocx(
      docx(
        documentXml(
          `<w:p><w:pPr><w:bidi/></w:pPr>${field(' PAGE ', '1')}${text(' از ')}${field(' NUMPAGES ', '1')}${text(' گزارش فارسی')}</w:p>`,
        ),
      ),
    );
    expect(
      getTextRuns(doc.blocks)
        .filter((run) => run.field)
        .map((run) => run.text),
    ).toEqual(['1', '1']);
    const markup = render(doc, 3, 5);
    expect(markup).toContain('>۳</span>');
    expect(markup).toContain('>۵</span>');
  });
});

describe('measured page construction', () => {
  it('creates separate sheets at the content height and respects keep-with-next', () => {
    const blocks = [
      paragraph('First'),
      paragraph('Heading', { keepNext: true }),
      paragraph('Next'),
    ];
    const pages = paginateBlocks(
      blocks,
      [{ heightPt: 70 }, { heightPt: 20 }, { heightPt: 30 }],
      () => 100,
    );
    expect(
      pages.map((page) => getTextRuns(page.blocks).map((run) => run.text)),
    ).toEqual([['First'], ['Heading', 'Next']]);
  });
  it('honors explicit paragraph and inline page breaks without reordering text', () => {
    const p = paragraph('A');
    p.children.push(
      { type: 'break', kind: 'page', properties: {} },
      { type: 'text', text: 'B', properties: {} },
    );
    const source = document([p]);
    const before = JSON.stringify(source);
    const blocks = splitPageBreaks(source.blocks);
    const pages = paginateBlocks(
      blocks,
      blocks.map(() => ({ heightPt: 10 })),
      () => 100,
    );
    expect(
      pages.map((page) => getTextRuns(page.blocks).map((run) => run.text)),
    ).toEqual([['A'], ['B']]);
    expect(JSON.stringify(source)).toBe(before);
  });
  it('splits tables at row boundaries and repeats the leading table header', () => {
    const rows = ['Heading', 'A', 'B', 'C'].map((value, index) => ({
      header: index === 0,
      cells: [
        { column: 0, columnSpan: 1, rowSpan: 1, blocks: [paragraph(value)] },
      ],
    }));
    const table: Table = {
      type: 'table',
      properties: {},
      columnsPt: [100],
      rows,
    };
    const pages = paginateBlocks(
      [table],
      [{ heightPt: 140, rowHeightsPt: [20, 40, 40, 40] }],
      () => 100,
    );
    expect(
      pages.map((page) => getTextRuns(page.blocks).map((run) => run.text)),
    ).toEqual([
      ['Heading', 'A', 'B'],
      ['Heading', 'C'],
    ]);
    expect(table.rows).toHaveLength(4);
  });
  it('keeps vertical merges together and retains oversize content without clipping', () => {
    const table: Table = {
      type: 'table',
      properties: {},
      columnsPt: [100],
      rows: [
        {
          cells: [
            {
              column: 0,
              columnSpan: 1,
              rowSpan: 2,
              blocks: [paragraph('Merged')],
            },
          ],
        },
        {
          cells: [
            {
              column: 0,
              columnSpan: 1,
              rowSpan: 1,
              continuation: true,
              blocks: [],
            },
          ],
        },
      ],
    };
    const pages = paginateBlocks(
      [table, paragraph('After')],
      [{ heightPt: 120, rowHeightsPt: [60, 60] }, { heightPt: 10 }],
      () => 100,
    );
    expect(pages[0].overflow).toBe(true);
    expect(pages[0].blocks[0].type).toBe('table');
    expect(pages[1].blocks).toHaveLength(1);
  });
  it('uses the correct content capacity for first and later page headers', () => {
    const pages = paginateBlocks(
      [paragraph('A'), paragraph('B'), paragraph('C')],
      [{ heightPt: 60 }, { heightPt: 40 }, { heightPt: 40 }],
      (number) => (number === 1 ? 70 : 100),
    );
    expect(pages.map((page) => page.blocks.length)).toEqual([1, 2]);
  });
});
