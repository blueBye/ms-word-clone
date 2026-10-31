import type { RunProperties, TextRun, ComplexScriptProperties } from '../../model/document.js';
import { childElements, childWord, halfPoints, isWord, onOff, wordAttribute } from './xml.js';

export function parseRunProperties(run: Element): RunProperties {
  const props: RunProperties = {};
  const rPr = childWord(run, 'rPr');
  if (!rPr) return props;
  const fonts = childWord(rPr, 'rFonts');
  if (fonts) {
    // Latin/high-ANSI collapsed intentionally; genuine script-specific font kept separately.
    props.fontFamily = wordAttribute(fonts, 'hAnsi') ?? wordAttribute(fonts, 'ascii');
  }
  props.bold = onOff(childWord(rPr, 'b'));
  props.italic = onOff(childWord(rPr, 'i'));
  props.fontSizePt = halfPoints(childWord(rPr, 'sz'));
  const direction = onOff(childWord(rPr, 'rtl'));
  if (direction !== undefined) props.direction = direction ? 'rtl' : 'ltr';
  const cs: ComplexScriptProperties = {};
  if (fonts) cs.fontFamily = wordAttribute(fonts, 'cs');
  cs.fontSizePt = halfPoints(childWord(rPr, 'szCs'));
  cs.bold = onOff(childWord(rPr, 'bCs'));
  cs.italic = onOff(childWord(rPr, 'iCs'));
  if (Object.values(cs).some(v => v !== undefined)) props.complexScript = cs;
  return props;
}

/** Flatten supported text-bearing run children, preserving run text order and whitespace. */
export function parseRun(run: Element): TextRun | undefined {
  const strings: string[] = [];
  for (const child of childElements(run)) {
    if (isWord(child, 't')) strings.push(child.textContent || '');
    else if (isWord(child, 'tab')) strings.push('\t');
    else if (isWord(child, 'br') || isWord(child, 'cr')) strings.push('\n');
    else if (isWord(child, 'noBreakHyphen')) strings.push('\u2011');
    else if (isWord(child, 'softHyphen')) strings.push('\u00AD');
    // Deleted/field instruction text and unsupported drawing elements are not visual text.
  }
  if (!strings.length) return undefined;
  return { type: 'text', text: strings.join(''), properties: parseRunProperties(run) };
}
