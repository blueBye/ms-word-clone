import type { Element } from '@xmldom/xmldom';
import type { DocxPackage } from './package.js';
import { NS } from './namespaces.js';
import { childElements, parseXml } from './parser/xml.js';

export interface Relationship {
  type: string;
  target: string;
  external: boolean;
}
export function resolvePartTarget(source: string, target: string): string {
  let decoded: string;
  try {
    decoded = decodeURIComponent(target);
  } catch {
    throw new Error('Invalid relationship target encoding');
  }
  if (
    /^[a-z][\w+.-]*:/i.test(decoded) ||
    decoded.startsWith('//') ||
    /[\\?#\u0000-\u001f]/.test(decoded)
  )
    throw new Error('Unsafe internal relationship target');
  const result = decoded.startsWith('/') ? [] : source.split('/').slice(0, -1);
  for (const component of decoded.replace(/^\//, '').split('/')) {
    if (component === '..') {
      if (!result.length) throw new Error('Relationship escapes the package');
      result.pop();
    } else if (component !== '.') {
      if (!component) throw new Error('Empty relationship path segment');
      result.push(component);
    }
  }
  if (!result.length) throw new Error('Empty relationship target');
  return result.join('/');
}
export async function readRelationships(
  pkg: DocxPackage,
  source: string,
): Promise<Map<string, Relationship>> {
  const parts = source.split('/');
  const filename = parts.pop()!;
  const path = [...parts, '_rels', `${filename}.rels`].join('/');
  const result = new Map<string, Relationship>();
  if (!pkg.hasPart(path)) return result;
  const root = parseXml(await pkg.readText(path)).documentElement!;
  if (
    root.localName !== 'Relationships' ||
    root.namespaceURI !== NS.relationships
  )
    throw new Error('Invalid part relationships');
  for (const node of childElements(root)) {
    if (
      node.localName !== 'Relationship' ||
      node.namespaceURI !== NS.relationships
    )
      continue;
    const id = node.getAttribute('Id');
    const target = node.getAttribute('Target');
    if (!id || !target || result.has(id))
      throw new Error('Invalid or duplicate relationship');
    const external = node.getAttribute('TargetMode') === 'External';
    result.set(id, {
      type: node.getAttribute('Type')?.split('/').pop() ?? '',
      target: external ? target : resolvePartTarget(source, target),
      external,
    });
  }
  return result;
}
export function relationshipAttribute(
  element: Element,
  name: string,
): string | undefined {
  return (
    element.getAttributeNS(NS.officeRelationships, name) ??
    element.getAttributeNS(NS.officeRelationshipsStrict, name) ??
    undefined
  );
}
export function safeHyperlink(target: string): string | undefined {
  return /^(https?:|mailto:)/i.test(target) && !/[\u0000-\u0020]/.test(target)
    ? target
    : undefined;
}
