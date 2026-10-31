import type { Element } from '@xmldom/xmldom';
import type { Table, TableCell } from '../../model/document.js';
import {
  childWord,
  contentElements,
  isWord,
  numericAttribute,
  onOff,
  wordAttribute,
  wordValue,
} from './xml.js';
import { colorValue, defined, parseBorders } from './properties.js';
import type { BlockParser, ParseContext } from './context.js';

function width(node: Element | undefined): TableCell['width'] {
  const value = numericAttribute(node, 'w');
  const type = node && wordAttribute(node, 'type');
  return value === undefined || !['pct', 'dxa'].includes(type ?? '')
    ? undefined
    : {
        value: value / (type === 'pct' ? 50 : 20),
        unit: type === 'pct' ? 'percent' : 'pt',
      };
}
function margins(node: Element | undefined) {
  if (!node) return undefined;
  const result: NonNullable<TableCell['margins']> = {};
  for (const side of ['top', 'left', 'bottom', 'right'] as const) {
    const n = numericAttribute(childWord(node, side), 'w');
    if (n !== undefined) result[side] = n / 20;
  }
  return result;
}
export function parseTable(
  element: Element,
  context: ParseContext | undefined,
  parseBlocks: BlockParser,
): Table {
  const direct = childWord(element, 'tblPr');
  const styleId = wordValue(direct, 'tblStyle');
  const layers = context?.styles.chain('table', styleId) ?? [];
  let properties: Table['properties'] = {};
  for (const node of [...layers.map((layer) => layer.table), direct]) {
    if (!node) continue;
    const bidi = onOff(childWord(node, 'bidiVisual'));
    const alignment = wordValue(node, 'jc');
    const border = parseBorders(childWord(node, 'tblBorders'), context?.theme);
    properties = {
      ...properties,
      ...defined({
        styleId,
        direction:
          bidi === undefined
            ? undefined
            : bidi
              ? ('rtl' as const)
              : ('ltr' as const),
        width: width(childWord(node, 'tblW')),
        alignment: ['left', 'right', 'center'].includes(alignment ?? '')
          ? (alignment as 'left' | 'right' | 'center')
          : undefined,
        indentPt:
          numericAttribute(childWord(node, 'tblInd'), 'w') === undefined
            ? undefined
            : numericAttribute(childWord(node, 'tblInd'), 'w')! / 20,
        layout: childWord(node, 'tblLayout')
          ? (wordAttribute(childWord(node, 'tblLayout')!, 'type') as
              'fixed' | 'autofit' | undefined)
          : undefined,
        backgroundColor: colorValue(
          childWord(node, 'shd'),
          'fill',
          context?.theme,
        ),
      }),
      borders: border
        ? { ...properties.borders, ...border }
        : properties.borders,
      cellMargins: {
        ...properties.cellMargins,
        ...margins(childWord(node, 'tblCellMar')),
      },
    };
  }
  if (direct && childWord(direct, 'tblpPr'))
    context?.warn(
      'floating-table',
      'Floating tables are displayed in document order.',
    );
  const columnsPt = contentElements(childWord(element, 'tblGrid') ?? element)
    .filter((n) => isWord(n, 'gridCol'))
    .map((n) => (numericAttribute(n, 'w') ?? 0) / 20);
  const active = new Map<number, TableCell>();
  const rows = contentElements(element)
    .filter((n) => isWord(n, 'tr'))
    .map((row) => {
      const trPr = childWord(row, 'trPr'),
        height = trPr && childWord(trPr, 'trHeight');
      let column =
        numericAttribute(trPr && childWord(trPr, 'gridBefore'), 'val') ?? 0;
      const nextActive = new Map<number, TableCell>();
      const cells = contentElements(row)
        .filter((n) => isWord(n, 'tc'))
        .map((cell) => {
          const tcPr = childWord(cell, 'tcPr');
          const result: TableCell = {
            blocks: parseBlocks(
              cell,
              context ? { ...context, tableLayers: layers } : undefined,
            ),
            column,
            columnSpan: Math.max(
              1,
              Math.min(
                1000,
                numericAttribute(tcPr && childWord(tcPr, 'gridSpan'), 'val') ??
                  1,
              ),
            ),
            rowSpan: 1,
          };
          for (const node of [...layers.map((layer) => layer.cell), tcPr]) {
            if (!node) continue;
            Object.assign(
              result,
              defined({
                width: width(childWord(node, 'tcW')),
                backgroundColor: colorValue(
                  childWord(node, 'shd'),
                  'fill',
                  context?.theme,
                ),
                verticalAlign: wordValue(node, 'vAlign'),
              }),
            );
            const border = parseBorders(
              childWord(node, 'tcBorders'),
              context?.theme,
            );
            if (border) result.borders = { ...result.borders, ...border };
            result.margins = {
              ...result.margins,
              ...margins(childWord(node, 'tcMar')),
            };
          }
          const merge = tcPr && childWord(tcPr, 'vMerge');
          if (merge) {
            const previous = active.get(column);
            if (
              wordAttribute(merge, 'val') !== 'restart' &&
              previous &&
              previous.columnSpan === result.columnSpan
            ) {
              previous.rowSpan++;
              result.continuation = true;
              nextActive.set(column, previous);
            } else nextActive.set(column, result);
          }
          column += result.columnSpan;
          return result;
        });
      active.clear();
      for (const [key, value] of nextActive) active.set(key, value);
      return {
        cells,
        heightPt: (numericAttribute(height, 'val') ?? NaN) / 20 || undefined,
        heightRule: height && wordAttribute(height, 'hRule'),
        header: onOff(trPr && childWord(trPr, 'tblHeader')),
        cantSplit: onOff(trPr && childWord(trPr, 'cantSplit')),
      };
    });
  return { type: 'table', properties, columnsPt, rows };
}
