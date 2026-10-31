import { readFileSync } from 'node:fs';
import { basename } from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';

for (const path of process.argv.slice(2)) {
  const parts = unzipSync(readFileSync(path));
  const xml = (name) =>
    parts[name]
      ? new DOMParser().parseFromString(
          strFromU8(parts[name]),
          'application/xml',
        )
      : null;
  const main = xml('word/document.xml');
  const nodes = Array.from(main.getElementsByTagName('*'));
  const counts = {};
  const values = {};
  const selected = new Set([
    'rFonts',
    'sz',
    'szCs',
    'pStyle',
    'rStyle',
    'jc',
    'spacing',
    'ind',
    'numId',
    'ilvl',
    'pgSz',
    'pgMar',
    'cols',
    'tblW',
    'tblLayout',
    'gridSpan',
    'vMerge',
    'textDirection',
    'vertAlign',
    'color',
    'shd',
    'u',
    'tab',
    'position',
    'bdr',
  ]);
  for (const node of nodes) {
    const key = node.tagName;
    counts[key] = (counts[key] ?? 0) + 1;
    if (selected.has(node.localName)) {
      const attrs = Array.from(node.attributes)
        .map((a) => `${a.localName}=${a.value}`)
        .sort()
        .join(' ');
      (values[key] ??= new Set()).add(attrs);
    }
  }
  const styles = xml('word/styles.xml');
  const styleInfo = styles
    ? Array.from(styles.getElementsByTagName('*'))
        .filter((n) => n.localName === 'style')
        .map((n) => ({
          id: n.getAttribute('w:styleId'),
          type: n.getAttribute('w:type'),
          default: n.getAttribute('w:default'),
          basedOn: Array.from(n.childNodes)
            .find((c) => c.localName === 'basedOn')
            ?.getAttribute('w:val'),
          props: Array.from(n.childNodes)
            .filter((c) =>
              ['pPr', 'rPr', 'tblPr', 'tcPr'].includes(c.localName),
            )
            .map((n) => n.toString()),
        }))
    : [];
  const rels = xml('word/_rels/document.xml.rels');
  const relTypes = rels
    ? Array.from(rels.getElementsByTagName('*'))
        .filter((n) => n.localName === 'Relationship')
        .map((n) => ({
          type: n.getAttribute('Type').split('/').pop(),
          external: n.getAttribute('TargetMode') === 'External',
          extension: n.getAttribute('Target').split('.').pop(),
        }))
    : [];
  const used = new Set(
    nodes
      .filter((n) => ['pStyle', 'rStyle', 'tblStyle'].includes(n.localName))
      .map((n) => n.getAttribute('w:val')),
  );
  for (const style of styleInfo) if (style.default === '1') used.add(style.id);
  for (let i = 0; i < styleInfo.length; i++)
    for (const style of styleInfo)
      if (used.has(style.id) && style.basedOn) used.add(style.basedOn);
  const stories = Object.keys(parts)
    .filter((n) => /^word\/(header|footer)\d+\.xml$/.test(n))
    .map((name) => {
      const counts = {};
      for (const n of Array.from(xml(name).getElementsByTagName('*')))
        counts[n.tagName] = (counts[n.tagName] ?? 0) + 1;
      return { name, counts };
    });
  const numbering = xml('word/numbering.xml');
  const levels =
    numbering &&
    Array.from(numbering.getElementsByTagName('*'))
      .filter((n) => n.localName === 'lvl' && n.getAttribute('w:ilvl') === '0')
      .map((n) => n.toString());
  const drawings = nodes
    .filter((n) => n.tagName === 'wp:anchor' || n.tagName === 'wp:inline')
    .map((n) => ({
      type: n.localName,
      attrs: Array.from(n.attributes).map((a) => [a.localName, a.value]),
      layout: Array.from(n.getElementsByTagName('*'))
        .filter((c) =>
          [
            'positionH',
            'positionV',
            'extent',
            'srcRect',
            'prstGeom',
            'solidFill',
          ].includes(c.localName),
        )
        .map((c) => c.toString()),
      inCell: !!(() => {
        let p = n.parentNode;
        while (p) {
          if (p.localName === 'tc') return true;
          p = p.parentNode;
        }
        return false;
      })(),
    }));
  const result = {
    file: basename(path),
    counts: Object.fromEntries(
      Object.entries(counts).filter(
        ([k]) =>
          k.startsWith('w:') ||
          ['wp:anchor', 'wp:inline', 'mc:AlternateContent'].includes(k),
      ),
    ),
    values: Object.fromEntries(
      Object.entries(values).map(([k, v]) => [k, [...v]]),
    ),
    defaults:
      styles &&
      Array.from(styles.getElementsByTagName('*'))
        .find((n) => n.localName === 'docDefaults')
        ?.toString(),
    styles: styleInfo.filter((s) => used.has(s.id)),
    numbering: levels,
    stories,
    drawings,
  };
  console.log(
    JSON.stringify(
      process.env.DOCX_LAYOUT_ONLY
        ? {
            file: result.file,
            defaults: result.defaults,
            background: nodes
              .find((n) => n.localName === 'background')
              ?.getAttribute('w:color'),
            drawings,
            stories,
            layout: result.values['w:pgMar'],
          }
        : result,
      null,
      2,
    ),
  );
}
