import type { Element } from '@xmldom/xmldom';
import type { ParagraphProperties, RunProperties } from '../model/document.js';
import {
  childElements,
  childWord,
  numericAttribute,
  wordAttribute,
  wordValue,
} from './parser/xml.js';
import {
  parseParagraphProperties,
  parseRunPropertyElement,
  type Theme,
} from './parser/properties.js';

interface Level {
  start: number;
  format: string;
  text: string;
  suffix: string;
  restart?: number;
  paragraph: ParagraphProperties;
  run: RunProperties;
}
function formatCounter(value: number, format: string): string {
  if (format === 'decimalZero') return String(value).padStart(2, '0');
  if (format === 'lowerLetter' || format === 'upperLetter') {
    let text = '';
    for (let n = value; n > 0; n = Math.floor((n - 1) / 26))
      text = String.fromCharCode(65 + ((n - 1) % 26)) + text;
    return format === 'lowerLetter' ? text.toLowerCase() : text;
  }
  if (format === 'upperRoman' || format === 'lowerRoman') {
    let text = '',
      n = value;
    for (const [amount, symbol] of [
      [1000, 'M'],
      [900, 'CM'],
      [500, 'D'],
      [400, 'CD'],
      [100, 'C'],
      [90, 'XC'],
      [50, 'L'],
      [40, 'XL'],
      [10, 'X'],
      [9, 'IX'],
      [5, 'V'],
      [4, 'IV'],
      [1, 'I'],
    ] as const)
      while (n >= amount) {
        text += symbol;
        n -= amount;
      }
    return format === 'lowerRoman' ? text.toLowerCase() : text;
  }
  return String(value);
}
export class Numbering {
  private definitions = new Map<string, Map<number, Level>>();
  private instances = new Map<string, Map<number, Level>>();
  private counters = new Map<string, number[]>();
  constructor(root: Element | undefined, theme: Theme) {
    if (!root) return;
    const parseLevel = (node: Element): Level => ({
      start: numericAttribute(childWord(node, 'start'), 'val') ?? 1,
      format: wordValue(node, 'numFmt') ?? 'decimal',
      text: wordValue(node, 'lvlText') ?? '%1.',
      suffix: wordValue(node, 'suff') ?? 'tab',
      restart: numericAttribute(childWord(node, 'lvlRestart'), 'val'),
      paragraph: parseParagraphProperties(childWord(node, 'pPr'), theme),
      run: parseRunPropertyElement(childWord(node, 'rPr'), theme),
    });
    for (const abstract of childElements(root).filter(
      (n) => n.localName === 'abstractNum',
    )) {
      this.definitions.set(
        wordAttribute(abstract, 'abstractNumId') ?? '',
        new Map(
          childElements(abstract)
            .filter((n) => n.localName === 'lvl')
            .map((node) => [
              numericAttribute(node, 'ilvl') ?? 0,
              parseLevel(node),
            ]),
        ),
      );
    }
    for (const num of childElements(root).filter(
      (n) => n.localName === 'num',
    )) {
      const levels = new Map(
        this.definitions.get(wordValue(num, 'abstractNumId') ?? ''),
      );
      for (const override of childElements(num).filter(
        (n) => n.localName === 'lvlOverride',
      )) {
        const index = numericAttribute(override, 'ilvl') ?? 0;
        const element = childWord(override, 'lvl');
        const level = element ? parseLevel(element) : levels.get(index);
        if (level)
          levels.set(index, {
            ...level,
            start:
              numericAttribute(childWord(override, 'startOverride'), 'val') ??
              level.start,
          });
      }
      this.instances.set(wordAttribute(num, 'numId') ?? '', levels);
    }
  }
  next(id: string, index: number) {
    const levels = this.instances.get(id);
    const level = levels?.get(index);
    if (!level || index < 0 || index > 8) return undefined;
    const counters = this.counters.get(id) ?? [];
    counters[index] =
      counters[index] === undefined ? level.start : counters[index] + 1;
    for (let i = index + 1; i <= 8; i++)
      if (
        levels!.get(i)?.restart !== 0 &&
        (levels!.get(i)?.restart ?? i) > index
      )
        delete counters[i];
    this.counters.set(id, counters);
    let marker =
      level.format === 'none'
        ? ''
        : level.format === 'bullet'
          ? level.text
          : level.text.replace(/%([1-9])/g, (_, digit: string) => {
              const slot = Number(digit) - 1,
                definition = levels!.get(slot);
              return formatCounter(
                counters[slot] ?? definition?.start ?? 1,
                definition?.format ?? 'decimal',
              );
            });
    // Common legacy Symbol/Wingdings bullets need Unicode fallbacks in browser fonts.
    if (level.format === 'bullet')
      marker = marker.replace(/\uF0B7/g, '•').replace(/\uF0A7/g, '▪');
    return {
      paragraph: level.paragraph,
      list: { marker, suffix: level.suffix, properties: level.run },
    };
  }
}
