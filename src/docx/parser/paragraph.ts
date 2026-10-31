import type { Paragraph, ParagraphAlignment, ParagraphProperties, TextRun } from '../../model/document.js';
import { childElements, childWord, isWord, onOff, wordAttribute } from './xml.js';
import { parseRun } from './run.js';

function paragraphProperties(p: Element): ParagraphProperties {
  const result: ParagraphProperties = {};
  const pPr = childWord(p, 'pPr');
  if (!pPr) return result;
  const bidi = onOff(childWord(pPr, 'bidi'));
  if (bidi !== undefined) result.direction = bidi ? 'rtl' : 'ltr';
  const alignment = wordAttribute(childWord(pPr, 'jc') ?? pPr, 'val');
  const mapping: Record<string, ParagraphAlignment> = {
    left: 'left', right: 'right', center: 'center', both: 'justify',
    start: result.direction === 'rtl' ? 'right' : 'left',
    end: result.direction === 'rtl' ? 'left' : 'right',
  };
  if (alignment && mapping[alignment]) result.alignment = mapping[alignment];
  return result;
}

export function parseParagraph(p: Element): Paragraph {
  const children: TextRun[] = [];
  function visit(parent: Element): void {
    for (const c of childElements(parent)) {
      if (isWord(c, 'r')) {
        const run = parseRun(c);
        if (run) children.push(run);
      } else if (isWord(c, 'hyperlink') || isWord(c, 'sdt') || isWord(c, 'sdtContent') || isWord(c, 'ins') || isWord(c, 'smartTag')) {
        // Hyperlink URLs and SDT metadata are not represented in Milestone 1.
        visit(c);
      }
    }
  }
  visit(p);
  return { type: 'paragraph', properties: paragraphProperties(p), children };
}
