import type {
  Block,
  DocumentIR,
  Inline,
  Paragraph,
  Table,
  TableRow,
} from '../model/document.js';

export interface BlockMeasurement {
  heightPt: number;
  rowHeightsPt?: number[];
  cells?: CellMeasurement[][];
  tableWidthPt?: number;
  columnWidthsPt?: number[];
}
export interface CellMeasurement {
  /** Block bounds include paragraph margins, in the cell's coordinate space. */
  blocks: { startPt: number; endPt: number }[];
  insetPt: number;
}

function cellHeight(measurement: CellMeasurement): number {
  const { blocks, insetPt } = measurement;
  return (
    insetPt +
    (blocks.length ? blocks[blocks.length - 1].endPt - blocks[0].startPt : 0)
  );
}

/** Split copies at cell block boundaries; the imported table stays unchanged. */
function splitRow(
  row: TableRow,
  measurements: CellMeasurement[] | undefined,
  available: number,
) {
  if (
    row.cantSplit ||
    row.heightRule === 'exact' ||
    !measurements ||
    measurements.length !== row.cells.length ||
    row.cells.some(
      (cell, index) =>
        cell.rowSpan !== 1 ||
        cell.continuation ||
        measurements[index].blocks.length !== cell.blocks.length,
    )
  )
    return null;

  const counts = row.cells.map((cell, index) => {
    const measurement = measurements[index];
    let count = 0;
    for (let end = 1; end <= cell.blocks.length; end++) {
      if (
        measurement.insetPt +
          measurement.blocks[end - 1].endPt -
          measurement.blocks[0].startPt >
        available
      )
        break;
      const block = cell.blocks[end - 1];
      if (
        end === cell.blocks.length ||
        block.type !== 'paragraph' ||
        !block.properties.keepNext
      )
        count = end;
    }
    return count;
  });
  if (
    !counts.some(Boolean) ||
    counts.every((count, index) => count === row.cells[index].blocks.length)
  )
    return null;

  const part = (remaining: boolean) => {
    const cells = row.cells.map((cell, index) => ({
      ...cell,
      blocks: remaining
        ? cell.blocks.slice(counts[index])
        : cell.blocks.slice(0, counts[index]),
      verticalAlign: 'top' as const,
    }));
    const measured = measurements.map((cell, index) => ({
      ...cell,
      blocks: remaining
        ? cell.blocks.slice(counts[index])
        : cell.blocks.slice(0, counts[index]),
    }));
    return {
      row: { ...row, heightPt: undefined, heightRule: undefined, cells },
      measurements: measured,
      heightPt: Math.max(0, ...measured.map(cellHeight)),
    };
  };
  return { first: part(false), rest: part(true) };
}
export interface PageSlice {
  blocks: Block[];
  overflow: boolean;
}

function isEmptyTableEndParagraph(
  block: Block,
  previous: Block | undefined,
): boolean {
  if (block.type !== 'paragraph' || previous?.type !== 'table') return false;
  const p = block.properties;
  return (
    !p.pageBreakBefore &&
    !p.list &&
    !p.backgroundColor &&
    !Object.values(p.borders ?? {}).some(
      (border) => !['nil', 'none'].includes(border.style),
    ) &&
    block.children.every(
      (child) =>
        child.type === 'text' &&
        !child.field &&
        !child.text.trim() &&
        !child.properties.backgroundColor &&
        !child.properties.strike &&
        (!child.properties.underline || child.properties.underline === 'none'),
    )
  );
}

export function pageVariant(
  document: DocumentIR,
  number: number,
): 'first' | 'default' | 'even' {
  if (number === 1 && document.page?.differentFirstPage) return 'first';
  return document.page?.evenAndOddHeaders && number % 2 === 0
    ? 'even'
    : 'default';
}

/** Preserve hard breaks without inserting text or mutating the imported model. */
export function splitPageBreaks(blocks: readonly Block[]): Block[] {
  return blocks.flatMap((block): Block[] => {
    if (
      block.type !== 'paragraph' ||
      !block.children.some(
        (child) => child.type === 'break' && child.kind === 'page',
      )
    )
      return [block];
    const parts: Paragraph[] = [];
    let children: Inline[] = [];
    let pageBreakBefore = block.properties.pageBreakBefore;
    const flush = () => {
      parts.push({
        ...block,
        children,
        properties: {
          ...block.properties,
          pageBreakBefore,
          firstLinePt: parts.length ? 0 : block.properties.firstLinePt,
          list: parts.length ? undefined : block.properties.list,
        },
      });
      children = [];
    };
    for (const inline of block.children) {
      if (inline.type === 'break' && inline.kind === 'page') {
        if (children.length) flush();
        else if (pageBreakBefore) flush();
        pageBreakBefore = true;
      } else children.push(inline);
    }
    if (children.length || !parts.length) flush();
    return parts;
  });
}

/** Browser-measured block/row pagination. Oversize indivisible content is retained, never clipped. */
export function paginateBlocks(
  blocks: readonly Block[],
  measurements: readonly BlockMeasurement[],
  capacityForPage: (number: number) => number,
): PageSlice[] {
  const pages: PageSlice[] = [];
  let page: PageSlice = { blocks: [], overflow: false };
  let used = 0;
  const capacity = () => Math.max(1, capacityForPage(pages.length + 1));
  const finish = () => {
    if (page.blocks.length) pages.push(page);
    page = { blocks: [], overflow: false };
    used = 0;
  };
  const add = (block: Block, height: number) => {
    page.blocks.push(block);
    used += height;
    page.overflow ||= used > capacity() + 0.5;
  };
  for (let index = 0; index < blocks.length; index++) {
    const block = blocks[index];
    const measurement = measurements[index] ?? { heightPt: 0 };
    const height = Math.max(0, measurement.heightPt);
    // A final empty paragraph after a table must not create a header-only sheet.
    // Retain it when it fits; never change the source IR or strip interior spacing.
    if (
      index === blocks.length - 1 &&
      page.blocks.length &&
      used + height > capacity() &&
      isEmptyTableEndParagraph(block, blocks[index - 1])
    )
      continue;
    if (block.type === 'paragraph' && block.properties.pageBreakBefore)
      finish();
    if (block.type === 'paragraph' && block.properties.keepNext) {
      let chain = height;
      for (let next = index + 1; next < blocks.length; next++) {
        chain += measurements[next]?.heightPt ?? 0;
        const following = blocks[next];
        if (following.type !== 'paragraph' || !following.properties.keepNext)
          break;
      }
      if (
        page.blocks.length &&
        chain <= capacity() &&
        used + chain > capacity()
      )
        finish();
    }
    if (
      block.type !== 'table' ||
      !block.rows.length ||
      height <= capacity() - used
    ) {
      if (page.blocks.length && used + height > capacity()) finish();
      add(block, height);
      continue;
    }

    const rowHeights = block.rows.map((_, row) =>
      Math.max(
        0,
        measurement.rowHeightsPt?.[row] ?? height / block.rows.length,
      ),
    );
    const overhead = Math.max(
      0,
      height - rowHeights.reduce((a, b) => a + b, 0),
    );
    let headerCount = 0;
    while (block.rows[headerCount]?.header) headerCount++;
    if (
      block.rows
        .slice(0, headerCount)
        .some((row, i) =>
          row.cells.some((cell) => i + cell.rowSpan > headerCount),
        )
    )
      headerCount = 0;
    const headerHeight = rowHeights
      .slice(0, headerCount)
      .reduce((a, b) => a + b, 0);
    let rows: TableRow[] = [];
    let fragmentHeight = overhead;
    let continued = false;
    const commit = () => {
      if (rows.length)
        add(
          {
            ...block,
            // Auto-layout must not resize columns when a fragment omits other rows.
            ...(measurement.columnWidthsPt?.length
              ? {
                  columnsPt: measurement.columnWidthsPt,
                  properties: {
                    ...block.properties,
                    layout: 'fixed',
                    width: measurement.tableWidthPt
                      ? { value: measurement.tableWidthPt, unit: 'pt' }
                      : block.properties.width,
                  },
                }
              : {}),
            rows,
          } as Table,
          fragmentHeight,
        );
      rows = [];
      fragmentHeight = overhead;
    };
    for (let row = 0; row < block.rows.length;) {
      let end = Math.min(
        block.rows.length,
        row === 0 ? Math.max(1, headerCount + 1) : row + 1,
      );
      // Keep intersecting vertical merges on one page.
      for (let scan = row; scan < end; scan++) {
        for (const cell of block.rows[scan].cells)
          end = Math.min(block.rows.length, Math.max(end, scan + cell.rowSpan));
      }
      let group = block.rows.slice(row, end);
      let groupHeight = rowHeights.slice(row, end).reduce((a, b) => a + b, 0);
      let cells = measurement.cells?.[row];
      for (;;) {
        const repeatHeight =
          !rows.length && continued && row >= headerCount ? headerHeight : 0;
        const available = capacity() - used - fragmentHeight - repeatHeight;
        const split =
          groupHeight > available && group.length === 1
            ? splitRow(group[0], cells, available)
            : null;
        if (
          groupHeight > available &&
          !split &&
          (page.blocks.length || rows.length)
        ) {
          commit();
          finish();
          continued = true;
          continue;
        }
        if (!rows.length && continued && row >= headerCount) {
          rows.push(...block.rows.slice(0, headerCount));
          fragmentHeight += headerHeight;
        }
        // Continue the row border to the content bottom when a paragraph moves
        // to the next page, rather than closing it immediately after the text.
        rows.push(
          ...(split
            ? [
                {
                  ...split.first.row,
                  heightPt: available,
                  heightRule: 'atLeast' as const,
                },
              ]
            : group),
        );
        fragmentHeight += split ? available : groupHeight;
        if (!split) break;
        commit();
        finish();
        continued = true;
        group = [split.rest.row];
        groupHeight = split.rest.heightPt;
        cells = split.rest.measurements;
      }
      row = end;
    }
    commit();
  }
  finish();
  return pages.length ? pages : [{ blocks: [], overflow: false }];
}
