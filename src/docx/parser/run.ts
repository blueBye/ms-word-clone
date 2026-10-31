import type { Element } from '@xmldom/xmldom';
import type { Inline, RunProperties } from '../../model/document.js';
import { childWord, contentElements, isWord, wordAttribute } from './xml.js';
import { parseRunPropertyElement } from './properties.js';
import { parseDrawing } from './drawing.js';
import type { BlockParser, ParseContext } from './context.js';

export function parseRunProperties(run: Element): RunProperties {
  return parseRunPropertyElement(childWord(run, 'rPr'));
}

/** Flatten supported text-bearing run children, preserving run text order and whitespace. */
export function parseRun(
  run: Element,
  base: RunProperties = {},
  context?: ParseContext,
  parseBlocks?: BlockParser,
): Inline[] {
  const direct = parseRunPropertyElement(childWord(run, 'rPr'), context?.theme);
  const properties = context ? context.styles.run(base, direct) : direct;
  const result: Inline[] = [];
  let text = '';
  let hasText = false;
  const flush = () => {
    if (hasText)
      result.push({
        type: 'text',
        text,
        properties,
        ...(context ? { directProperties: direct } : {}),
      });
    text = '';
    hasText = false;
  };
  for (const child of contentElements(run)) {
    if (isWord(child, 't')) {
      text += child.textContent || '';
      hasText = true;
    } else if (isWord(child, 'tab')) {
      text += '\t';
      hasText = true;
    } else if (
      isWord(child, 'br') &&
      ['page', 'column'].includes(wordAttribute(child, 'type') ?? '')
    ) {
      flush();
      result.push({
        type: 'break',
        kind: wordAttribute(child, 'type') as 'page' | 'column',
        properties,
      });
    } else if (isWord(child, 'br') || isWord(child, 'cr')) {
      text += '\n';
      hasText = true;
    } else if (isWord(child, 'noBreakHyphen')) {
      text += '\u2011';
      hasText = true;
    } else if (isWord(child, 'softHyphen')) {
      text += '\u00AD';
      hasText = true;
    } else if (isWord(child, 'drawing') || isWord(child, 'pict')) {
      flush();
      result.push(...parseDrawing(child, context, parseBlocks));
    }
  }
  flush();
  return result;
}
