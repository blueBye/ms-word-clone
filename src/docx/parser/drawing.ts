import type { Element } from '@xmldom/xmldom';
import type { DrawingInline, DrawingPosition } from '../../model/document.js';
import { relationshipAttribute } from '../relationships.js';
import { childElements, descendants } from './xml.js';
import type { BlockParser, ParseContext } from './context.js';

function numberAttribute(
  node: Element | undefined,
  name: string,
  divisor = 1,
): number | undefined {
  const raw = node?.getAttribute(name);
  return raw && Number.isFinite(Number(raw))
    ? Number(raw) / divisor
    : undefined;
}
function drawingColor(
  node: Element | undefined,
  context?: ParseContext,
): string | undefined {
  if (!node) return undefined;
  const color = childElements(node)[0];
  const value =
    color?.localName === 'schemeClr'
      ? context?.theme.colors[color.getAttribute('val') ?? '']
      : color?.getAttribute('val');
  return value && /^[\da-f]{6}$/i.test(value) ? `#${value}` : undefined;
}
export function parseDrawing(
  root: Element,
  context?: ParseContext,
  parseBlocks?: BlockParser,
): DrawingInline[] {
  const anchors = descendants(root, 'anchor').concat(
    descendants(root, 'inline'),
  );
  if (!anchors.length) {
    context?.warn(
      'unsupported-vml',
      'A legacy VML drawing cannot be positioned accurately.',
    );
    return [];
  }
  return anchors.map((anchor) => {
    const nodes = descendants(anchor, 'extent');
    const extent = nodes[0];
    const image = descendants(anchor, 'blip')[0];
    const id = image && relationshipAttribute(image, 'embed');
    const rel = id && context?.relationships.get(id);
    const dataUrl =
      rel && !rel.external && rel.type === 'image'
        ? context?.assets.get(rel.target)
        : undefined;
    const shapeRoot = descendants(anchor, 'wsp')[0];
    const shapeProperties =
      shapeRoot && childElements(shapeRoot).find((n) => n.localName === 'spPr');
    const pictureProperties =
      descendants(anchor, 'pic')[0] &&
      descendants(descendants(anchor, 'pic')[0], 'spPr')[0];
    const geometry = descendants(
      shapeProperties ?? pictureProperties ?? anchor,
      'prstGeom',
    )[0];
    const propertyChildren = shapeProperties
      ? childElements(shapeProperties)
      : [];
    const shapeStyle =
      shapeRoot &&
      childElements(shapeRoot).find((n) => n.localName === 'style');
    const fillReference =
      shapeStyle &&
      childElements(shapeStyle).find((n) => n.localName === 'fillRef');
    const fill = propertyChildren.some((n) => n.localName === 'noFill')
      ? undefined
      : drawingColor(
          propertyChildren.find((n) => n.localName === 'solidFill') ??
            fillReference,
          context,
        );
    const line = propertyChildren.find((n) => n.localName === 'ln');
    const crop = descendants(anchor, 'srcRect')[0];
    const text = descendants(anchor, 'txbxContent').find((n) =>
      n.namespaceURI?.includes('wordprocessingml'),
    );
    const docPr = descendants(anchor, 'docPr')[0];
    const position = (local: string): DrawingPosition => {
      const node = childElements(anchor).find((n) => n.localName === local);
      const offset =
        node &&
        childElements(node).find((n) => n.localName === 'posOffset')
          ?.textContent;
      return {
        relativeTo: node?.getAttribute('relativeFrom') ?? 'paragraph',
        offsetPt:
          offset && Number.isFinite(Number(offset))
            ? Number(offset) / 12700
            : undefined,
        alignment:
          (node &&
            childElements(node).find((n) => n.localName === 'align')
              ?.textContent) ||
          undefined,
      };
    };
    if (image && !dataUrl)
      context?.warn(
        'unsupported-image',
        'An external, missing or unsupported image was not loaded.',
      );
    const kind = geometry?.getAttribute('prst') ?? 'rect';
    if (shapeRoot && !['rect', 'ellipse', 'roundRect'].includes(kind))
      context?.warn(
        'shape-approximation',
        'A drawing geometry is approximated by its bounding box.',
      );
    return {
      type: 'drawing' as const,
      widthPt: numberAttribute(extent, 'cx', 12700) ?? 24,
      heightPt: numberAttribute(extent, 'cy', 12700) ?? 24,
      alt: docPr?.getAttribute('descr') || docPr?.getAttribute('name') || '',
      image: dataUrl
        ? {
            dataUrl,
            crop: crop
              ? {
                  left: numberAttribute(crop, 'l', 100000),
                  top: numberAttribute(crop, 't', 100000),
                  right: numberAttribute(crop, 'r', 100000),
                  bottom: numberAttribute(crop, 'b', 100000),
                }
              : undefined,
          }
        : undefined,
      shape:
        shapeRoot || geometry
          ? {
              kind,
              fill,
              stroke:
                line &&
                drawingColor(
                  childElements(line).find((n) => n.localName === 'solidFill'),
                  context,
                ),
              strokeWidthPt: numberAttribute(line, 'w', 12700),
              rotation: numberAttribute(
                descendants(anchor, 'xfrm')[0],
                'rot',
                60000,
              ),
            }
          : undefined,
      blocks: text && parseBlocks ? parseBlocks(text, context) : undefined,
      position:
        anchor.localName === 'anchor'
          ? {
              horizontal: position('positionH'),
              vertical: position('positionV'),
              behindText: anchor.getAttribute('behindDoc') === '1',
              zIndex: Number(anchor.getAttribute('relativeHeight') ?? 0),
              wrap:
                childElements(anchor).find((n) =>
                  n.localName?.startsWith('wrap'),
                )?.localName ?? 'wrapNone',
            }
          : undefined,
    };
  });
}
