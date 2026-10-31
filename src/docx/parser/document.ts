import type { DocumentIR } from '../../model/document.js';
import { childElements, childWord, isWord } from './xml.js';
import { parseParagraph } from './paragraph.js';
export function parseDocumentXml(xml: Document): DocumentIR {
  const root = xml.documentElement;
  if (!isWord(root, 'document')) throw new Error('Expected a WordprocessingML document root');
  const body = childWord(root, 'body');
  if (!body) throw new Error('DOCX document is missing w:body');
  return { type: 'document', blocks: childElements(body).filter(el => isWord(el, 'p')).map(parseParagraph) };
}
