/** An OOXML-independent, deliberately small milestone-1 document model. */
export interface DocumentIR {
  type: 'document';
  blocks: Paragraph[];
}
export type TextDirection = 'ltr' | 'rtl';
export type ParagraphAlignment = 'left' | 'right' | 'center' | 'justify';
export interface ParagraphProperties {
  alignment?: ParagraphAlignment;
  direction?: TextDirection;
}
export interface Paragraph {
  type: 'paragraph';
  properties: ParagraphProperties;
  children: TextRun[];
}
export interface ComplexScriptProperties {
  fontFamily?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
}
export interface RunProperties {
  direction?: TextDirection;
  fontFamily?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  complexScript?: ComplexScriptProperties;
}
export interface TextRun {
  type: 'text';
  text: string;
  properties: RunProperties;
}
