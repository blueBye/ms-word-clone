import type { Document } from '@xmldom/xmldom';
import type { Element } from '@xmldom/xmldom';
import type { Block, DocumentIR } from '../../model/document.js';
import { contentElements, childWord, isWord } from './xml.js';
import { parseParagraph } from './paragraph.js';
import { parseTable } from './table.js';
import type { ParseContext } from './context.js';

export function parseBlocks(parent: Element, context?: ParseContext): Block[] {
  return contentElements(parent).flatMap((element): Block[] => {
    if (isWord(element, 'p'))
      return [parseParagraph(element, context, parseBlocks)];
    if (isWord(element, 'tbl'))
      return [parseTable(element, context, parseBlocks)];
    if (isWord(element, 'sdt')) {
      const content = childWord(element, 'sdtContent');
      return content ? parseBlocks(content, context) : [];
    }
    if (
      ['sdtContent', 'customXml', 'ins', 'moveTo'].some((name) =>
        isWord(element, name),
      )
    )
      return parseBlocks(element, context);
    return [];
  });
}
export function parseDocumentXml(
  xml: Document,
  context?: ParseContext,
): DocumentIR {
  const root = xml.documentElement;
  if (!root || !isWord(root, 'document'))
    throw new Error('Expected a WordprocessingML document root');
  const body = childWord(root, 'body');
  if (!body) throw new Error('DOCX document is missing w:body');
  return {
    type: 'document',
    schemaVersion: 2,
    blocks: parseBlocks(body, context),
  };
}
