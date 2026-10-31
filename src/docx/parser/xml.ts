import { DOMParser } from '@xmldom/xmldom';
import { WORD_NS } from '../namespaces.js';

export function parseXml(xml: string): Document {
  const errors: string[] = [];
  const document = new DOMParser({
    errorHandler: {
      warning: (msg: string) => { errors.push(msg); },
      error: (msg: string) => { errors.push(msg); },
      fatalError: (msg: string) => { errors.push(msg); },
    },
  }).parseFromString(xml, 'application/xml');
  if (!document?.documentElement || errors.length || document.documentElement.localName === 'parsererror') {
    throw new Error(`Invalid XML: ${errors.join('; ') || 'no document element'}`);
  }
  return document;
}
export function childElements(element: Element): Element[] {
  const children: Element[] = [];
  for (let node = element.firstChild; node; node = node.nextSibling) {
    if (node.nodeType === 1) children.push(node as Element);
  }
  return children;
}
export function isWord(element: Element, local: string): boolean {
  return WORD_NS.has(element.namespaceURI || '') && element.localName === local;
}
export function childWord(element: Element, local: string): Element | undefined {
  return childElements(element).find(c => isWord(c, local));
}
export function wordAttribute(element: Element, local: string): string | undefined {
  for (const ns of WORD_NS) {
    const val = element.getAttributeNS(ns, local);
    if (val !== null) return val;
  }
  return undefined;
}
/** ST_OnOff: omitted val means on; explicit 0/false/off means off. */
export function onOff(element: Element | undefined): boolean | undefined {
  if (!element) return undefined;
  const value = wordAttribute(element, 'val');
  if (value === undefined) return true;
  if (['1', 'true', 'on'].includes(value.toLowerCase())) return true;
  if (['0', 'false', 'off'].includes(value.toLowerCase())) return false;
  return undefined;
}
export function halfPoints(element: Element | undefined): number | undefined {
  if (!element) return undefined;
  const raw = wordAttribute(element, 'val');
  if (!raw || !/^[1-9]\d*$/.test(raw)) return undefined;
  const value = Number(raw) / 2;
  return Number.isFinite(value) ? value : undefined;
}
