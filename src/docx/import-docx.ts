import type {
  DocumentDiagnostic,
  DocumentIR,
  PageLayout,
  StoryVariants,
} from '../model/document.js';
import {
  findMainDocumentPart,
  openDocxPackage,
  type ImportOptions,
} from './package.js';
import {
  childElements,
  childWord,
  descendants,
  isWord,
  numericAttribute,
  onOff,
  parseXml,
  wordAttribute,
} from './parser/xml.js';
import { parseBlocks, parseDocumentXml } from './parser/document.js';
import {
  readRelationships,
  relationshipAttribute,
  type Relationship,
} from './relationships.js';
import { Styles, parseTheme } from './styles.js';
import { Numbering } from './numbering.js';
import { colorValue } from './parser/properties.js';
import type { ParseContext } from './parser/context.js';

/** Import to preview IR; source packages are not yet round-trippable. */
export async function importDocx(
  input: ArrayBuffer | Uint8Array,
  options: ImportOptions = {},
): Promise<DocumentIR> {
  const pkg = await openDocxPackage(input, options);
  const path = await findMainDocumentPart(pkg);
  const relationships = await readRelationships(pkg, path);
  const relatedXml = async (type: string) => {
    const relation = [...relationships.values()].find(
      (rel) => rel.type === type && !rel.external,
    );
    if (!relation || !pkg.hasPart(relation.target)) return undefined;
    const root = parseXml(await pkg.readText(relation.target)).documentElement!;
    const valid =
      type === 'theme'
        ? root.localName === 'theme' &&
          [
            'http://schemas.openxmlformats.org/drawingml/2006/main',
            'http://purl.oclc.org/ooxml/drawingml/main',
          ].includes(root.namespaceURI ?? '')
        : isWord(root, type);
    if (!valid) throw new Error(`Invalid ${type} part root`);
    return root;
  };
  const [styleXml, themeXml, numberingXml] = await Promise.all([
    relatedXml('styles'),
    relatedXml('theme'),
    relatedXml('numbering'),
  ]);
  const theme = parseTheme(themeXml);
  const diagnostics: DocumentDiagnostic[] = [];
  const warningKeys = new Set<string>();
  const warning = (part: string) => (code: string, message: string) => {
    const key = `${part}:${code}`;
    if (!warningKeys.has(key)) {
      warningKeys.add(key);
      diagnostics.push({ code, message, part });
    }
  };
  const styles = new Styles(styleXml, theme, warning(path));
  const assets = new Map<string, string>();
  async function context(
    part: string,
    rels: Map<string, Relationship>,
  ): Promise<ParseContext> {
    for (const rel of rels.values()) {
      if (
        rel.external ||
        rel.type !== 'image' ||
        assets.has(rel.target) ||
        !pkg.hasPart(rel.target)
      )
        continue;
      const extension = rel.target.split('.').pop()?.toLowerCase() ?? '';
      const mime: Record<string, string> = {
        png: 'image/png',
        jpg: 'image/jpeg',
        jpeg: 'image/jpeg',
        gif: 'image/gif',
        webp: 'image/webp',
        bmp: 'image/bmp',
      };
      if (!mime[extension]) continue; // Prefer DrawingML's raster fallback to active SVG/unsupported formats.
      const bytes = await pkg.readBinary(rel.target);
      let binary = '';
      for (let i = 0; i < bytes.length; i += 8192)
        binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
      assets.set(rel.target, `data:${mime[extension]};base64,${btoa(binary)}`);
    }
    return {
      styles,
      theme,
      numbering: new Numbering(numberingXml, theme),
      relationships: rels,
      assets,
      warn: warning(part),
    };
  }
  const xml = parseXml(await pkg.readText(path));
  const ctx = await context(path, relationships);
  const document = parseDocumentXml(xml, ctx);
  const body = childWord(xml.documentElement!, 'body')!;
  const section = childWord(body, 'sectPr');
  if (section) {
    const size = childWord(section, 'pgSz'),
      margins = childWord(section, 'pgMar');
    const page: PageLayout = {
      widthPt: (numericAttribute(size, 'w') ?? 11906) / 20,
      heightPt: (numericAttribute(size, 'h') ?? 16838) / 20,
      margins: {},
      differentFirstPage: onOff(childWord(section, 'titlePg')),
    };
    for (const side of [
      'top',
      'right',
      'bottom',
      'left',
      'header',
      'footer',
    ] as const) {
      const value = numericAttribute(margins, side);
      if (value !== undefined) page.margins![side] = value / 20;
    }
    document.page = page;
    for (const kind of ['header', 'footer'] as const) {
      const variants: StoryVariants = {};
      for (const ref of childElements(section).filter(
        (n) => n.localName === `${kind}Reference`,
      )) {
        const id = relationshipAttribute(ref, 'id');
        const rel = id && relationships.get(id);
        if (
          !rel ||
          rel.external ||
          rel.type !== kind ||
          !pkg.hasPart(rel.target)
        ) {
          ctx.warn(
            'missing-story',
            'A header or footer part could not be resolved.',
          );
          continue;
        }
        const type = wordAttribute(ref, 'type');
        if (!['first', 'default', 'even'].includes(type ?? '')) continue;
        const story = parseXml(await pkg.readText(rel.target)).documentElement!;
        variants[type as keyof StoryVariants] = parseBlocks(
          story,
          await context(rel.target, await readRelationships(pkg, rel.target)),
        );
      }
      if (Object.keys(variants).length)
        document[kind === 'header' ? 'headers' : 'footers'] = variants;
    }
  }
  document.backgroundColor = colorValue(
    childWord(xml.documentElement!, 'background'),
    'color',
    theme,
  );
  if (!document.backgroundColor) delete document.backgroundColor;
  if (descendants(xml.documentElement!, 'sectPr').length > 1)
    ctx.warn(
      'multiple-sections',
      'Multiple section layouts are approximated using the final section settings.',
    );
  if (document.headers || document.footers)
    ctx.warn(
      'story-pagination',
      'Headers and footers are shown once in the flowing preview; repeating page variants and page numbers are not recalculated.',
    );
  if (descendants(xml.documentElement!, 'anchor').length)
    ctx.warn(
      'anchored-layout',
      'Positioned drawings use browser layout; Word wrapping and pagination can differ.',
    );
  if (diagnostics.length) document.diagnostics = diagnostics;
  return document;
}
