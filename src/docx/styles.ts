import type { Element } from '@xmldom/xmldom';
import type { ParagraphProperties, RunProperties } from '../model/document.js';
import {
  childElements,
  childWord,
  descendants,
  wordAttribute,
  wordValue,
} from './parser/xml.js';
import {
  mergeParagraph,
  mergeRun,
  parseParagraphProperties,
  parseRunPropertyElement,
  type Theme,
} from './parser/properties.js';

export function parseTheme(root: Element | undefined): Theme {
  const theme: Theme = { colors: {}, fonts: {} };
  if (!root) return theme;
  const colors = descendants(root, 'clrScheme')[0];
  for (const child of colors ? childElements(colors) : []) {
    const value = childElements(child)[0];
    const color = value?.getAttribute('lastClr') ?? value?.getAttribute('val');
    if (child.localName && color && /^[\da-f]{6}$/i.test(color))
      theme.colors[child.localName] = color;
  }
  for (const [key, source] of Object.entries({
    text1: 'dk1',
    text2: 'dk2',
    background1: 'lt1',
    background2: 'lt2',
  }))
    theme.colors[key] = theme.colors[source];
  for (const family of ['major', 'minor']) {
    const fonts = descendants(root, `${family}Font`)[0];
    for (const child of fonts ? childElements(fonts) : []) {
      const slot =
        child.localName === 'ea'
          ? 'eastAsia'
          : child.localName === 'font'
            ? child.getAttribute('script')
            : child.localName;
      const name = child.getAttribute('typeface');
      if (slot && name) theme.fonts[`${family}.${slot}`] = name;
    }
  }
  return theme;
}
export interface StyleLayer {
  paragraph: ParagraphProperties;
  run: RunProperties;
  table?: Element;
  cell?: Element;
}
export class Styles {
  private styles = new Map<string, { element: Element; type: string }>();
  private defaultIds = new Map<string, string>();
  readonly defaults: StyleLayer;
  constructor(
    root: Element | undefined,
    readonly theme: Theme,
    private warn: (code: string, message: string) => void,
  ) {
    const defaults = root && childWord(root, 'docDefaults');
    this.defaults = {
      paragraph: parseParagraphProperties(
        defaults &&
          childWord(childWord(defaults, 'pPrDefault') ?? defaults, 'pPr'),
        theme,
      ),
      run: parseRunPropertyElement(
        defaults &&
          childWord(childWord(defaults, 'rPrDefault') ?? defaults, 'rPr'),
        theme,
      ),
    };
    for (const element of root
      ? childElements(root).filter((n) => n.localName === 'style')
      : []) {
      const id = wordAttribute(element, 'styleId');
      const type = wordAttribute(element, 'type') ?? '';
      if (!id) continue;
      this.styles.set(id, { element, type });
      if (['1', 'true'].includes(wordAttribute(element, 'default') ?? ''))
        this.defaultIds.set(type, id);
    }
  }
  chain(type: string, id?: string): StyleLayer[] {
    let next = id ?? this.defaultIds.get(type);
    const result: StyleLayer[] = [];
    const visited = new Set<string>();
    while (next) {
      if (visited.has(next)) {
        this.warn('style-cycle', 'A style inheritance cycle was stopped.');
        break;
      }
      visited.add(next);
      const style = this.styles.get(next);
      if (!style || style.type !== type) {
        this.warn('missing-style', 'A referenced style could not be resolved.');
        break;
      }
      const element = style.element;
      result.unshift({
        paragraph: parseParagraphProperties(
          childWord(element, 'pPr'),
          this.theme,
        ),
        run: parseRunPropertyElement(childWord(element, 'rPr'), this.theme),
        table: childWord(element, 'tblPr'),
        cell: childWord(element, 'tcPr'),
      });
      next = wordValue(element, 'basedOn');
    }
    return result;
  }
  paragraph(
    direct: ParagraphProperties,
    tableLayers: StyleLayer[] = [],
  ): { properties: ParagraphProperties; run: RunProperties } {
    let properties = this.defaults.paragraph;
    let run = this.defaults.run;
    for (const layer of [
      ...tableLayers,
      ...this.chain('paragraph', direct.styleId),
    ]) {
      properties = mergeParagraph(properties, layer.paragraph);
      run = mergeRun(run, layer.run, true);
    }
    return { properties: mergeParagraph(properties, direct), run };
  }
  run(base: RunProperties, direct: RunProperties): RunProperties {
    let result = base;
    // A default character style is not automatically applied to all text.
    if (direct.styleId)
      for (const layer of this.chain('character', direct.styleId))
        result = mergeRun(result, layer.run, true);
    return mergeRun(result, direct);
  }
}
