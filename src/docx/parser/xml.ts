import { DOMParser, type Document, type Element } from '@xmldom/xmldom';
import { WORD_NS } from '../namespaces.js';

export function parseXml(xml: string): Document {
  try {
    const document = new DOMParser({
      onError: (_level, message) => {
        throw new Error(message);
      },
    }).parseFromString(xml, 'application/xml');
    if (!document.documentElement) throw new Error('no document element');
    // OPC XML parts cannot contain DTDs. xmldom does not load external entities.
    if (document.doctype) throw new Error('DOCTYPE is not allowed in OPC XML');
    return document;
  } catch (error) {
    throw new Error(
      `Invalid XML: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
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
export function childWord(
  element: Element,
  local: string,
): Element | undefined {
  return childElements(element).find((c) => isWord(c, local));
}
export function descendants(
  element: Element,
  local: string,
  namespaces?: ReadonlySet<string>,
): Element[] {
  return Array.from(element.getElementsByTagName('*')).filter(
    (el) =>
      el.localName === local &&
      (!namespaces || namespaces.has(el.namespaceURI ?? '')),
  );
}
export function wordValue(
  element: Element | undefined,
  child: string,
): string | undefined {
  const node = element && childWord(element, child);
  return node ? wordAttribute(node, 'val') : undefined;
}
export function numericAttribute(
  element: Element | undefined,
  name: string,
): number | undefined {
  const raw = element && wordAttribute(element, name);
  if (raw === undefined || raw.trim() === '') return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const SUPPORTED_CHOICES = new Set([
  ...WORD_NS,
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  'http://schemas.openxmlformats.org/drawingml/2006/main',
  'http://schemas.openxmlformats.org/drawingml/2006/picture',
]);
/** Select exactly one compatibility branch, never duplicate choice/fallback content. */
export function contentElements(parent: Element): Element[] {
  return childElements(parent).flatMap((child) => {
    if (child.namespaceURI !== MC || child.localName !== 'AlternateContent')
      return [child];
    const branches = childElements(child);
    const selected =
      branches.find(
        (branch) =>
          branch.localName === 'Choice' &&
          (branch.getAttribute('Requires') ?? '')
            .split(/\s+/)
            .filter(Boolean)
            .every((prefix) =>
              SUPPORTED_CHOICES.has(branch.lookupNamespaceURI(prefix) ?? ''),
            ),
      ) ?? branches.find((branch) => branch.localName === 'Fallback');
    return selected ? contentElements(selected) : [];
  });
}
export function wordAttribute(
  element: Element,
  local: string,
): string | undefined {
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
