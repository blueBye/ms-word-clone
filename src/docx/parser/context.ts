import type { Element } from '@xmldom/xmldom';
import type { Block, Paragraph } from '../../model/document.js';
import type { Styles, StyleLayer } from '../styles.js';
import type { Numbering } from '../numbering.js';
import type { Relationship } from '../relationships.js';
import type { Theme } from './properties.js';

export interface ParseContext {
  styles: Styles;
  theme: Theme;
  numbering: Numbering;
  relationships: Map<string, Relationship>;
  assets: Map<string, string>;
  tableLayers?: StyleLayer[];
  numberingLanguage?: string;
  onParagraph?(source: Element, paragraph: Paragraph): void;
  warn(code: string, message: string): void;
}
export type BlockParser = (parent: Element, context?: ParseContext) => Block[];
