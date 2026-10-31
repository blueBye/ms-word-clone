import { describe, expect, it } from 'vitest';
import { DocumentRenderer } from '../src/react/DocumentRenderer.js';
import { renderToStaticMarkup } from 'react-dom/server';
import type { DocumentIR, RunProperties } from '../src/model/document.js';

function render(text: string, properties: RunProperties = {}) {
  const document: DocumentIR = {
    type: 'document',
    schemaVersion: 2,
    blocks: [
      {
        type: 'paragraph',
        properties: { direction: 'rtl', alignment: 'right' },
        children: [{ type: 'text', text, properties }],
      },
    ],
  };
  const before = JSON.stringify(document);
  const html = renderToStaticMarkup(
    DocumentRenderer({ document, className: 'preview' }),
  );
  expect(JSON.stringify(document)).toBe(before);
  return html;
}

// Component-output checks only. Browser shaping, bidi and line layout need visual fixtures.
describe('React preview output', () => {
  it('preserves text, explicit paragraph direction and base font settings', () => {
    const html = render(' A\tB\n ', {
      fontFamily: 'Example Font',
      fontSizePt: 11.5,
      bold: true,
      italic: true,
    });
    expect(html).toContain('class="preview"');
    expect(html).toContain('<p dir="rtl"');
    expect(html).toContain('text-align:right');
    expect(html).toContain(' A\tB\n ');
    expect(html).toContain('font-family:&quot;Example Font&quot;');
    expect(html).toContain('font-size:11.5pt');
    expect(html).toContain('font-weight:700');
    expect(html).toContain('font-style:italic');
  });
  it('applies explicit complex-script overrides, including disabled values', () => {
    const html = render('قرارداد', {
      direction: 'rtl',
      bold: true,
      italic: true,
      complexScript: {
        fontFamily: 'B Zar',
        fontSizePt: 12,
        bold: false,
        italic: false,
      },
    });
    expect(html).toContain('font-family:&quot;B Zar&quot;');
    expect(html).toContain('font-size:12pt');
    expect(html).toContain('font-weight:400');
    expect(html).toContain('font-style:normal');
  });
  it('keeps explicit base off values rather than inheriting browser emphasis', () => {
    const html = render('Text', { bold: false, italic: false });
    expect(html).toContain('font-weight:400');
    expect(html).toContain('font-style:normal');
  });
});
