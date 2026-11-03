import { describe, expect, it } from 'vitest';
import { renderToStaticMarkup } from 'react-dom/server';
import {
  getParagraphs,
  getTextRuns,
  importDocx,
  type DocumentIR,
} from '../src/index.js';
import { DocumentRenderer } from '../src/react/index.js';
import { resolvePartTarget } from '../src/docx/relationships.js';
import {
  docx,
  documentXml,
  W,
  W_STRICT,
  OFFICE_REL,
  REL,
} from './helpers/docx.js';

const p = (text: string, props = '', runProps = '') =>
  `<w:p><w:pPr>${props}</w:pPr><w:r><w:rPr>${runProps}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;
const cell = (body: string, props = '') =>
  `<w:tc><w:tcPr>${props}</w:tcPr>${body}</w:tc>`;
const table = (body: string, props = '') =>
  `<w:tbl><w:tblPr>${props}</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid>${body}</w:tbl>`;
const styles = (body: string) => `<w:styles xmlns:w="${W}">${body}</w:styles>`;
const rels = (
  records: { id: string; type: string; target: string; external?: boolean }[],
) =>
  `<Relationships xmlns="${REL}">${records.map((r) => `<Relationship Id="${r.id}" Type="${OFFICE_REL}/${r.type}" Target="${r.target}"${r.external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;

async function load(
  body: string,
  parts: Record<string, string | Uint8Array> = {},
) {
  const references = Object.keys(parts)
    .filter((name) =>
      ['styles', 'theme', 'numbering'].includes(name.replace('.xml', '')),
    )
    .map((name) => ({
      id: name,
      type: name.replace('.xml', ''),
      target: name,
    }));
  return importDocx(
    docx(documentXml(body), {
      'custom/_rels/main.xml.rels': rels(references),
      ...Object.fromEntries(
        Object.entries(parts).map(([name, data]) => [
          name.startsWith('custom/') ? name : `custom/${name}`,
          data,
        ]),
      ),
    }),
  );
}
const html = (document: DocumentIR, pageLayout = false) =>
  renderToStaticMarkup(DocumentRenderer({ document, pageLayout }));

describe('sample-derived table coverage', () => {
  it('keeps all résumé text inside a two-column table, in source order', async () => {
    const document = await load(
      table(
        `<w:tr>${cell(p('Name'), '<w:gridSpan w:val="2"/>')}</w:tr><w:tr>${cell(p('Experience'))}${cell(p('Skills'))}</w:tr>`,
        '<w:tblBorders><w:top w:val="nil"/></w:tblBorders>',
      ),
    );
    const t = document.blocks[0];
    expect(t.type).toBe('table');
    if (t.type !== 'table') throw new Error('Expected table');
    expect(t.columnsPt).toEqual([100, 200]);
    expect(t.rows[0].cells[0].columnSpan).toBe(2);
    expect(getTextRuns(document.blocks).map((r) => r.text)).toEqual([
      'Name',
      'Experience',
      'Skills',
    ]);
    expect(html(document)).toMatch(/colspan="2"/i);
    expect(html(document)).toContain('Skills');
  });
  it('separates RTL table column order from paragraph direction and respects defaults', async () => {
    const document = await load(
      table(
        `<w:tr>${cell(p('فارسی'))}${cell(p('English', '<w:bidi w:val="0"/>'))}</w:tr>`,
        '<w:bidiVisual/>',
      ),
      {
        'styles.xml': styles(
          '<w:docDefaults><w:pPrDefault><w:pPr><w:bidi/></w:pPr></w:pPrDefault></w:docDefaults>',
        ),
      },
    );
    expect(document.blocks[0].properties.direction).toBe('rtl');
    expect(
      getParagraphs(document.blocks).map((p) => p.properties.direction),
    ).toEqual(['rtl', 'ltr']);
    expect(html(document)).toContain('<table dir="rtl"');
    expect(html(document)).toContain('<p dir="ltr"');
  });
  it('supports nested tables, widths, shading, explicit borders and cell margins', async () => {
    const nested = table(`<w:tr>${cell(p('Nested'))}</w:tr>`);
    const document = await load(
      table(
        `<w:tr>${cell(nested, '<w:tcW w:type="pct" w:w="2500"/><w:shd w:fill="EEECE1"/><w:tcMar><w:left w:w="100"/></w:tcMar><w:tcBorders><w:top w:val="nil"/></w:tcBorders>')}</w:tr>`,
        '<w:tblW w:type="dxa" w:w="6000"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="single" w:sz="8" w:color="FF0000"/></w:tblBorders>',
      ),
    );
    expect(getTextRuns(document.blocks)[0].text).toBe('Nested');
    const output = html(document);
    expect(output.match(/<table /g)).toHaveLength(2);
    expect(output).toContain('table-layout:fixed');
    expect(output).toContain('background-color:#EEECE1');
    expect(output).toContain('border-top:hidden');
    expect(output).toContain('width:50%');
  });
  it('resolves vertical merges only across contiguous matching grid spans', async () => {
    const document = await load(
      table(
        `<w:tr>${cell(p('Merged'), '<w:vMerge w:val="restart"/>')}${cell(p('A'))}</w:tr><w:tr>${cell(p(''), '<w:vMerge/>')}${cell(p('B'))}</w:tr><w:tr>${cell(p('Gap'))}${cell(p('C'))}</w:tr><w:tr>${cell(p('New'), '<w:vMerge/>')}${cell(p('D'))}</w:tr>`,
      ),
    );
    const t = document.blocks[0];
    if (t.type !== 'table') throw new Error('Expected table');
    expect(t.rows[0].cells[0].rowSpan).toBe(2);
    expect(t.rows[1].cells[0].continuation).toBe(true);
    expect(t.rows[3].cells[0].continuation).toBeUndefined();
    expect(html(document)).toMatch(/rowspan="2"/i);
  });
  it('inherits table styles without replacing direct paragraph overrides', async () => {
    const document = await load(
      table(
        `<w:tr>${cell(p('Text', '<w:spacing w:after="60"/>'))}</w:tr>`,
        '<w:tblStyle w:val="Grid"/>',
      ),
      {
        'styles.xml': styles(
          '<w:style w:type="table" w:styleId="Grid"><w:pPr><w:bidi/><w:spacing w:after="0"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr><w:tblPr><w:tblBorders><w:insideH w:val="single" w:sz="4"/></w:tblBorders><w:tblCellMar><w:right w:w="120"/></w:tblCellMar></w:tblPr></w:style>',
        ),
      },
    );
    expect(getParagraphs(document.blocks)[0].properties).toMatchObject({
      direction: 'rtl',
      spaceAfterPt: 3,
    });
    expect(getTextRuns(document.blocks)[0].properties.fontSizePt).toBe(10);
    expect(document.blocks[0].properties).toMatchObject({
      cellMargins: { right: 6 },
      borders: { insideH: { widthPt: 0.5 } },
    });
  });
});

describe('styles, fonts and typography', () => {
  it('uses natural single-line font height inside shaded cells while retaining inherited padding', async () => {
    const runProps =
      '<w:rFonts w:ascii="Example Body" w:hAnsi="Example Body" w:cs="Example Body"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/>';
    const document = await load(
      table(
        `<w:tr>${cell(
          p('First line', '', runProps) + p('Second line', '', runProps),
          '<w:shd w:fill="FFFFCC"/>',
        )}</w:tr>`,
        '<w:tblStyle w:val="Grid"/><w:bidiVisual/>',
      ),
      {
        'styles.xml': styles(
          '<w:docDefaults><w:pPrDefault><w:pPr><w:bidi/><w:spacing w:line="240" w:lineRule="auto" w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="table" w:styleId="Grid"><w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>',
        ),
      },
    );
    const before = JSON.stringify(document);
    const output = html(document);
    expect(output).toContain('background-color:#FFFFCC');
    expect(output).toContain('padding:0pt 5.4pt 0pt 5.4pt');
    expect(output.match(/<p\b[^>]*line-height:normal/g)).toHaveLength(2);
    expect(output).not.toContain('line-height:1;');
    expect(output).toContain('font-family:&quot;Example Body&quot;');
    expect(output).toContain('font-size:12pt');
    expect(JSON.stringify(document)).toBe(before);
  });

  it('preserves exact line spacing as an explicit point value', async () => {
    const document = await load(
      p('Text', '<w:spacing w:line="240" w:lineRule="exact"/>'),
    );
    expect(html(document)).toContain('line-height:12pt');
    expect(getParagraphs(document.blocks)[0].properties.lineSpacing).toEqual({
      rule: 'exact',
      value: 12,
    });
  });

  it('resolves document defaults, basedOn and explicit character/direct overrides', async () => {
    const document = await load(
      p(
        'Test',
        '<w:pStyle w:val="Child"/>',
        '<w:rStyle w:val="Emphasis"/><w:b w:val="0"/>',
      ),
      {
        'styles.xml': styles(
          '<w:docDefaults><w:pPrDefault><w:pPr><w:bidi/></w:pPr></w:pPrDefault><w:rPrDefault><w:rPr><w:rFonts w:ascii="Latin" w:cs="B Nazanin"/><w:sz w:val="24"/><w:szCs w:val="28"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:spacing w:before="240"/></w:pPr><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/><w:pPr><w:jc w:val="start"/></w:pPr></w:style><w:style w:type="character" w:styleId="Emphasis"><w:rPr><w:i/></w:rPr></w:style>',
        ),
      },
    );
    expect(getParagraphs(document.blocks)[0].properties).toMatchObject({
      direction: 'rtl',
      alignment: 'start',
      spaceBeforePt: 12,
    });
    expect(getTextRuns(document.blocks)[0].properties).toMatchObject({
      fontFamily: 'Latin',
      fontSizePt: 12,
      bold: false,
      italic: true,
      complexScript: { fontFamily: 'B Nazanin', fontSizePt: 14 },
    });
    expect(getTextRuns(document.blocks)[0].directProperties).toMatchObject({
      styleId: 'Emphasis',
      bold: false,
    });
  });
  it('uses toggle semantics within styles, and absolute on/off for direct formatting', async () => {
    const document = await load(
      p('Toggle', '<w:pStyle w:val="Child"/>') +
        p('Direct', '<w:pStyle w:val="Child"/>', '<w:b/>'),
      {
        'styles.xml': styles(
          '<w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:b/><w:i/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/><w:rPr><w:b/><w:i w:val="0"/></w:rPr></w:style>',
        ),
      },
    );
    expect(
      getTextRuns(document.blocks).map((r) => [
        r.properties.bold,
        r.properties.italic,
      ]),
    ).toEqual([
      [false, true],
      [true, true],
    ]);
  });
  it('stops cyclic styles and reports missing styles', async () => {
    const document = await load(
      p('Cycle', '<w:pStyle w:val="A"/>') +
        p('Missing', '<w:pStyle w:val="Missing"/>'),
      {
        'styles.xml': styles(
          '<w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="B"/></w:style><w:style w:type="paragraph" w:styleId="B"><w:basedOn w:val="A"/></w:style>',
        ),
      },
    );
    expect(document.diagnostics?.map((d) => d.code)).toEqual(
      expect.arrayContaining(['style-cycle', 'missing-style']),
    );
  });
  it('resolves theme fonts and color references, with literal font slot overrides', async () => {
    const document = await load(
      p(
        'Theme',
        '',
        '<w:rFonts w:ascii="Custom"/><w:color w:themeColor="accent1"/>',
      ),
      {
        'styles.xml': styles(
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>',
        ),
        'theme.xml':
          '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="test"><a:accent1><a:srgbClr val="23DFA2"/></a:accent1></a:clrScheme><a:fontScheme name="test"><a:minorFont><a:latin typeface="Theme Sans"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>',
      },
    );
    expect(getTextRuns(document.blocks)[0].properties).toMatchObject({
      fontSlots: { ascii: 'Custom', hAnsi: 'Theme Sans' },
      color: '#23DFA2',
      fontSizePt: 11,
    });
  });
  it('keeps Latin, Hebrew and Persian text in logical order with separate fonts', async () => {
    const document = await load(
      p(
        'English فارسی עברית',
        '',
        '<w:rFonts w:ascii="Latin" w:hAnsi="Accent" w:cs="Persian"/>',
      ),
    );
    const output = html(document);
    expect(output).toContain('font-family:&quot;Latin&quot;');
    expect(output).toContain('font-family:&quot;Persian&quot;');
    expect(output.indexOf('English')).toBeLessThan(output.indexOf('فارسی'));
    expect(getTextRuns(document.blocks)[0].text).toBe('English فارسی עברית');
  });
  it.each(['<w:rtl/>', '<w:cs/>'])(
    'forces complex-script typography for neutral runs with %s',
    async (property) => {
      const document = await load(
        p(
          ' 123 ',
          '',
          `<w:rFonts w:ascii="Latin" w:cs="Persian"/><w:b/>${property}`,
        ),
      );
      const output = html(document);
      expect(output).toContain('font-family:&quot;Persian&quot;');
      expect(output).toContain('font-weight:400');
      expect(output).not.toContain('font-weight:700');
    },
  );
  it('supports color, highlight, underline, strike, script position and paragraph spacing', async () => {
    const document = await load(
      p(
        'Marks',
        '<w:spacing w:before="120" w:after="60" w:line="360"/><w:ind w:left="240" w:hanging="100"/>',
        '<w:color w:val="23DFA2"/><w:highlight w:val="yellow"/><w:u w:val="double"/><w:strike/><w:vertAlign w:val="superscript"/>',
      ),
    );
    expect(getParagraphs(document.blocks)[0].properties).toMatchObject({
      spaceBeforePt: 6,
      spaceAfterPt: 3,
      lineSpacing: { rule: 'auto', value: 1.5 },
      indentLeftPt: 12,
      firstLinePt: -5,
    });
    const output = html(document);
    expect(output).toContain('color:#23DFA2');
    expect(output).toContain('background-color:#FFFF00');
    expect(output).toContain('text-decoration-line:underline');
    expect(output).toContain('text-decoration:line-through solid');
    expect(output).toContain('vertical-align:super');
  });
});

describe('numbering, stories and assets', () => {
  it('numbers independent instances with start overrides and suppresses numId=0', async () => {
    const np = (id: string) =>
      `<w:numPr><w:numId w:val="${id}"/><w:ilvl w:val="0"/></w:numPr>`;
    const document = await load(
      p('A', np('1')) + p('B', np('1')) + p('C', np('2')) + p('D', np('0')),
      {
        'numbering.xml': `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1)"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num></w:numbering>`,
      },
    );
    expect(
      getParagraphs(document.blocks).map((p) => p.properties.list?.marker),
    ).toEqual(['1)', '2)', '5)', undefined]);
  });
  it('normalizes common private-use bullets and preserves marker font settings', async () => {
    const document = await load(
      p('Bullet', '<w:numPr><w:numId w:val="1"/></w:numPr>'),
      {
        'numbering.xml': `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:rPr><w:rFonts w:ascii="Symbol"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
      },
    );
    expect(getParagraphs(document.blocks)[0].properties.list).toMatchObject({
      marker: '•',
      properties: { fontFamily: 'Symbol' },
    });
  });
  it('reads document page dimensions and a dark document background', async () => {
    const bytes = docx(
      documentXml(
        p('Dark') +
          '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:left="720" w:right="720" w:top="720" w:bottom="720"/></w:sectPr>',
      ).replace('<w:body>', '<w:background w:color="1A1A1A"/><w:body>'),
    );
    const document = await importDocx(bytes);
    expect(document.page).toMatchObject({
      widthPt: 595.3,
      heightPt: 841.9,
      margins: { left: 36, top: 36 },
    });
    expect(html(document, true)).toContain('background-color:#1A1A1A');
    expect(html(document, true)).toContain('color:#fff');
  });
  it('keeps inline drawing order and embedded raster images with dimensions', async () => {
    const document = await load(
      `<w:p><w:r><w:t>Before</w:t>${imageDrawing()}<w:t>After</w:t></w:r></w:p>`,
      {
        'custom/_rels/main.xml.rels': rels([
          { id: 'img', type: 'image', target: 'media/pixel.png' },
        ]),
        'media/pixel.png': pixel,
      },
    );
    expect(
      getParagraphs(document.blocks)[0].children.map((n) => n.type),
    ).toEqual(['text', 'drawing', 'text']);
    const output = html(document);
    expect(output).toContain('data:image/png;base64,');
    expect(output).toContain('width:72pt');
    expect(output.indexOf('Before')).toBeLessThan(output.indexOf('<img'));
    expect(output.indexOf('<img')).toBeLessThan(output.indexOf('After'));
  });
  it('blocks external image loading and unsafe hyperlinks without dropping their text', async () => {
    const document = await load(
      `<w:p><w:hyperlink xmlns:r="${OFFICE_REL}" r:id="bad"><w:r><w:t>Visible</w:t></w:r></w:hyperlink><w:r>${imageDrawing()}</w:r></w:p>`,
      {
        'custom/_rels/main.xml.rels': rels([
          {
            id: 'bad',
            type: 'hyperlink',
            target: 'javascript:alert(1)',
            external: true,
          },
          {
            id: 'img',
            type: 'image',
            target: 'https://example.invalid/pixel.png',
            external: true,
          },
        ]),
      },
    );
    expect(getTextRuns(document.blocks)[0].href).toBeUndefined();
    expect(html(document)).not.toContain('src="https:');
    expect(document.diagnostics?.map((d) => d.code)).toContain(
      'unsupported-image',
    );
  });
  it('renders safe hyperlink targets with character-style color', async () => {
    const document = await load(
      `<w:p><w:hyperlink xmlns:r="${OFFICE_REL}" r:id="link"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Profile</w:t></w:r></w:hyperlink></w:p>`,
      {
        'custom/_rels/main.xml.rels': rels([
          {
            id: 'link',
            type: 'hyperlink',
            target: 'https://example.com',
            external: true,
          },
          { id: 'styles', type: 'styles', target: 'styles.xml' },
        ]),
        'styles.xml': styles(
          '<w:style w:type="character" w:styleId="Hyperlink"><w:rPr><w:color w:val="467886"/><w:u/></w:rPr></w:style>',
        ),
      },
    );
    expect(html(document)).toContain('href="https://example.com"');
    expect(html(document)).toContain('color:#467886');
  });
  it('loads header-local image relationships and preserves story variants', async () => {
    const document = await load(
      p('Body') +
        `<w:sectPr xmlns:r="${OFFICE_REL}"><w:headerReference w:type="first" r:id="header"/><w:titlePg/></w:sectPr>`,
      {
        'custom/_rels/main.xml.rels': rels([
          { id: 'header', type: 'header', target: 'stories/header.xml' },
        ]),
        'stories/header.xml': `<w:hdr xmlns:w="${W}">${p('Header')}<w:p><w:r>${imageDrawing()}</w:r></w:p></w:hdr>`,
        'stories/_rels/header.xml.rels': rels([
          { id: 'img', type: 'image', target: '../media/pixel.png' },
        ]),
        'media/pixel.png': pixel,
      },
    );
    expect(getTextRuns(document.headers!.first!)[0].text).toBe('Header');
    expect(html(document, true)).toContain('data-docx-story="header"');
    expect(html(document, true)).toContain('data:image/png;base64,');
  });
  it('selects a single supported AlternateContent branch and imports textbox text once', async () => {
    const drawing = `<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wp:inline><wp:extent cx="914400" cy="914400"/><wps:wsp><wps:txbx><w:txbxContent>${p('Textbox')}</w:txbxContent></wps:txbx></wps:wsp></wp:inline></w:drawing>`;
    const document = await load(
      `<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><mc:Choice Requires="wps">${drawing}</mc:Choice><mc:Fallback><w:t>Duplicate fallback</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>`,
    );
    expect(getTextRuns(document.blocks).map((r) => r.text)).toEqual([
      'Textbox',
    ]);
    expect(html(document).match(/Textbox/g)).toHaveLength(1);
    expect(html(document)).not.toContain('Duplicate fallback');
  });
  it('uses AlternateContent fallback when a required namespace is unsupported', async () => {
    const document = await load(
      `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:Choice Requires="x">${p('Unsupported')}</mc:Choice><mc:Fallback>${p('Fallback')}</mc:Fallback></mc:AlternateContent>`,
    );
    expect(getTextRuns(document.blocks)[0].text).toBe('Fallback');
  });
  it('preserves typed page and column breaks instead of converting them to line breaks', async () => {
    const document = await load(
      '<w:p><w:r><w:t>A</w:t><w:br w:type="page"/><w:t>B</w:t><w:br w:type="column"/><w:t>C</w:t><w:br/><w:t>D</w:t></w:r></w:p>',
    );
    const children = getParagraphs(document.blocks)[0].children;
    expect(children.map((n) => n.type)).toEqual([
      'text',
      'break',
      'text',
      'break',
      'text',
    ]);
    expect(
      children.filter((n) => n.type === 'break').map((n) => n.kind),
    ).toEqual(['page', 'column']);
    expect(getTextRuns(document.blocks).map((n) => n.text)).toEqual([
      'A',
      'B',
      'C\nD',
    ]);
  });
  it('supports Strict table markup too', async () => {
    const document = await importDocx(
      docx(documentXml(table(`<w:tr>${cell(p('Strict'))}</w:tr>`), W_STRICT)),
    );
    expect(document.blocks[0].type).toBe('table');
    expect(getTextRuns(document.blocks)[0].text).toBe('Strict');
  });
});

describe('source-relative OPC paths', () => {
  it('resolves valid parent targets without permitting package escape', () => {
    expect(
      resolvePartTarget('word/headers/header.xml', '../media/image.png'),
    ).toBe('word/media/image.png');
    expect(resolvePartTarget('word/document.xml', '/media/image.png')).toBe(
      'media/image.png',
    );
    expect(() =>
      resolvePartTarget('word/document.xml', '../../outside.xml'),
    ).toThrow('escapes');
    expect(() =>
      resolvePartTarget('word/document.xml', 'http%3Aevil.xml'),
    ).toThrow('Unsafe');
  });
});

describe('additional drawing and typography regressions', () => {
  it('rejects foreign namespaces in related Word parts', async () => {
    await expect(
      load(p('Text'), { 'styles.xml': '<styles xmlns="urn:foreign"/>' }),
    ).rejects.toThrow('Invalid styles part root');
  });
  it('uses theme fill references for anchored résumé decoration', async () => {
    const drawing = `<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wp:anchor behindDoc="1" relativeHeight="3"><wp:positionH relativeFrom="column"><wp:posOffset>127000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>254000</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wps:wsp><wps:spPr><a:prstGeom prst="ellipse"/></wps:spPr><wps:style><a:fillRef idx="1"><a:schemeClr val="dk1"/></a:fillRef></wps:style></wps:wsp></wp:anchor></w:drawing>`;
    const document = await load(`<w:p><w:r>${drawing}</w:r></w:p>`, {
      'theme.xml':
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme><a:dk1><a:srgbClr val="123456"/></a:dk1></a:clrScheme></a:themeElements></a:theme>',
    });
    const output = html(document);
    expect(output).toContain('background-color:#123456');
    expect(output).toContain('border-radius:50%');
    expect(output).toContain('left:10pt');
    expect(output).toContain('top:20pt');
  });
  it('preserves explicit zero run spacing and baseline overrides from inherited formatting', async () => {
    const document = await load(
      p(
        'Zero',
        '<w:pStyle w:val="Shift"/>',
        '<w:position w:val="0"/><w:spacing w:val="0"/>',
      ),
      {
        'styles.xml': styles(
          '<w:style w:type="paragraph" w:styleId="Shift"><w:rPr><w:position w:val="6"/><w:spacing w:val="20"/></w:rPr></w:style>',
        ),
      },
    );
    expect(getTextRuns(document.blocks)[0].properties).toMatchObject({
      positionPt: 0,
      letterSpacingPt: 0,
    });
  });
  it('applies paragraph mark typography only to the empty line, not to existing runs', async () => {
    const document = await load(
      '<w:p><w:pPr><w:rPr><w:sz w:val="4"/></w:rPr></w:pPr></w:p>' +
        p('Normal', '<w:rPr><w:sz w:val="4"/></w:rPr>'),
    );
    const paragraphs = getParagraphs(document.blocks);
    expect(paragraphs[0].markProperties?.fontSizePt).toBe(2);
    expect(
      getTextRuns(document.blocks)[0].properties.fontSizePt,
    ).toBeUndefined();
    expect(html(document)).toContain('font-size:2pt');
    expect(html(document)).toContain('<br/>');
  });
  it('restarts nested counters when the parent level advances', async () => {
    const np = (level: number) =>
      `<w:numPr><w:numId w:val="1"/><w:ilvl w:val="${level}"/></w:numPr>`;
    const document = await load(
      p('A', np(0)) +
        p('B', np(1)) +
        p('C', np(1)) +
        p('D', np(0)) +
        p('E', np(1)),
      {
        'numbering.xml': `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
      },
    );
    expect(
      getParagraphs(document.blocks).map((p) => p.properties.list?.marker),
    ).toEqual(['1.', '1.1', '1.2', '2.', '2.1']);
  });
});

const pixel = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
function imageDrawing() {
  return `<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${OFFICE_REL}"><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Test image"/><a:graphic><a:blip r:embed="img"/></a:graphic></wp:inline></w:drawing>`;
}
