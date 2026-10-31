import type { Element } from '@xmldom/xmldom';
import type { Paragraph, Inline } from '../../model/document.js';
import { childWord, contentElements, isWord, wordAttribute } from './xml.js';
import { parseRun } from './run.js';
import {
  mergeParagraph,
  mergeRun,
  parseParagraphProperties,
  parseRunPropertyElement,
} from './properties.js';
import { relationshipAttribute, safeHyperlink } from '../relationships.js';
import type { BlockParser, ParseContext } from './context.js';

export function parseParagraph(
  p: Element,
  context?: ParseContext,
  parseBlocks?: BlockParser,
): Paragraph {
  const pPr = childWord(p, 'pPr');
  const direct = parseParagraphProperties(pPr, context?.theme);
  const resolved = context?.styles.paragraph(direct, context.tableLayers);
  let properties = resolved?.properties ?? direct;
  const numbering = properties.numbering;
  if (numbering?.id && numbering.id !== '0') {
    const level = context?.numbering.next(numbering.id, numbering.level ?? 0);
    if (level)
      properties = {
        ...mergeParagraph(level.paragraph, properties),
        list: level.list,
      };
    else
      context?.warn(
        'missing-numbering',
        'A list definition could not be resolved.',
      );
  }
  const children: Inline[] = [];
  function visit(parent: Element, href?: string): void {
    for (const c of contentElements(parent)) {
      if (isWord(c, 'r')) {
        const inlines = parseRun(c, resolved?.run, context, parseBlocks);
        for (const inline of inlines)
          if (inline.type === 'text' && href) inline.href = href;
        children.push(...inlines);
      } else if (isWord(c, 'sdt')) {
        const content = childWord(c, 'sdtContent');
        if (content) visit(content, href);
      } else if (isWord(c, 'hyperlink')) {
        const id = relationshipAttribute(c, 'id');
        const rel = id && context?.relationships.get(id);
        const anchor = wordAttribute(c, 'anchor');
        visit(
          c,
          anchor
            ? `#${encodeURIComponent(anchor)}`
            : rel && rel.type === 'hyperlink'
              ? safeHyperlink(rel.target)
              : undefined,
        );
      } else if (
        [
          'sdtContent',
          'ins',
          'smartTag',
          'customXml',
          'fldSimple',
          'moveTo',
        ].some((name) => isWord(c, name))
      ) {
        visit(c, href);
      }
    }
  }
  visit(p);
  const mark = pPr && childWord(pPr, 'rPr');
  return {
    type: 'paragraph',
    properties,
    children,
    ...(context ? { directProperties: direct } : {}),
    ...(mark
      ? {
          markProperties: mergeRun(
            resolved?.run ?? {},
            parseRunPropertyElement(mark, context?.theme),
          ),
        }
      : {}),
  };
}
