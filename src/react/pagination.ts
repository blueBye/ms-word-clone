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
}
export interface PageSlice {
  blocks: Block[];
  overflow: boolean;
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
      if (rows.length) add({ ...block, rows } as Table, fragmentHeight);
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
      const groupHeight = rowHeights.slice(row, end).reduce((a, b) => a + b, 0);
      const repeatHeight =
        !rows.length && continued && row >= headerCount ? headerHeight : 0;
      if (
        (page.blocks.length || rows.length) &&
        used + fragmentHeight + repeatHeight + groupHeight > capacity()
      ) {
        commit();
        finish();
        continued = true;
      }
      if (!rows.length && continued && row >= headerCount) {
        rows.push(...block.rows.slice(0, headerCount));
        fragmentHeight += headerHeight;
      }
      rows.push(...block.rows.slice(row, end));
      fragmentHeight += groupHeight;
      row = end;
    }
    commit();
  }
  finish();
  return pages.length ? pages : [{ blocks: [], overflow: false }];
}
