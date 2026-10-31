import type { Element } from '@xmldom/xmldom';
import type { Paragraph, Inline } from '../../model/document.js';
import {
  childWord,
  childElements,
  contentElements,
  isWord,
  wordAttribute,
} from './xml.js';
import { localizeDigits } from '../../model/numerals.js';
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
  const fields: { instruction: string; result: boolean }[] = [];
  const fieldKind = (instruction: string) =>
    /^\s*NUMPAGES\b/i.test(instruction)
      ? ('pages' as const)
      : /^\s*PAGE\b/i.test(instruction)
        ? ('page' as const)
        : undefined;
  function visit(parent: Element, href?: string): void {
    for (const c of contentElements(parent)) {
      if (isWord(c, 'r')) {
        for (const node of childElements(c)) {
          if (
            isWord(node, 'fldChar') &&
            wordAttribute(node, 'fldCharType') === 'begin'
          )
            fields.push({ instruction: '', result: false });
          if (isWord(node, 'instrText') && fields.length)
            fields[fields.length - 1].instruction += node.textContent ?? '';
          if (
            isWord(node, 'fldChar') &&
            wordAttribute(node, 'fldCharType') === 'separate' &&
            fields.length
          )
            fields[fields.length - 1].result = true;
        }
        const inlines = parseRun(c, resolved?.run, context, parseBlocks);
        for (const inline of inlines)
          if (inline.type === 'text') {
            if (href) inline.href = href;
            const field = fields.at(-1);
            if (field?.result) inline.field = fieldKind(field.instruction);
          }
        children.push(...inlines);
        for (const node of childElements(c))
          if (
            isWord(node, 'fldChar') &&
            wordAttribute(node, 'fldCharType') === 'end'
          )
            fields.pop();
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
      } else if (isWord(c, 'fldSimple')) {
        fields.push({
          instruction: wordAttribute(c, 'instr') ?? '',
          result: true,
        });
        visit(c, href);
        fields.pop();
      } else if (
        ['sdtContent', 'ins', 'smartTag', 'customXml', 'moveTo'].some((name) =>
          isWord(c, name),
        )
      ) {
        visit(c, href);
      }
    }
  }
  visit(p);
  const mark = pPr && childWord(pPr, 'rPr');
  const markProperties = mergeRun(
    resolved?.run ?? {},
    parseRunPropertyElement(mark, context?.theme),
  );
  if (properties.list) {
    const textRuns = children.filter((inline) => inline.type === 'text');
    const text = textRuns.map((run) => run.text).join('');
    const language = /[پچژگکی۰-۹]/u.test(text)
      ? 'fa-IR'
      : markProperties.complexScriptLanguage?.startsWith('fa')
        ? markProperties.complexScriptLanguage
        : (context?.numberingLanguage ?? markProperties.complexScriptLanguage);
    const locale = properties.direction === 'rtl' ? language : undefined;
    properties = {
      ...properties,
      list: {
        ...properties.list,
        properties: mergeRun(markProperties, properties.list.properties),
        locale,
        marker: properties.list.format?.startsWith('decimal')
          ? localizeDigits(properties.list.marker, locale)
          : properties.list.marker,
      },
    };
  }
  return {
    type: 'paragraph',
    properties,
    children,
    ...(context ? { directProperties: direct } : {}),
    ...(mark
      ? {
          markProperties,
        }
      : {}),
  };
}
