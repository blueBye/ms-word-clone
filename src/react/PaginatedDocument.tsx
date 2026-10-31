'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import type { DocumentIR } from '../model/document.js';
import { DocumentRenderer } from './DocumentRenderer.js';
import {
  paginateBlocks,
  pageVariant,
  splitPageBreaks,
  type PageSlice,
  type BlockMeasurement,
} from './pagination.js';

export interface PaginatedDocumentProps {
  document: DocumentIR;
  className?: string;
  /** Increment after registering a FontFace that was already loaded. */
  fontRevision?: number;
  onPageCountChange?: (count: number) => void;
}
type Insets = Record<
  'first' | 'default' | 'even',
  { top: number; bottom: number }
>;
interface Layout {
  source: DocumentIR;
  pages: PageSlice[];
  insets: Insets;
}

export function PaginatedDocument({
  document: doc,
  className,
  fontRevision = 0,
  onPageCountChange,
}: PaginatedDocumentProps) {
  const blocks = useMemo(() => splitPageBreaks(doc.blocks), [doc]);
  const measureRef = useRef<HTMLDivElement>(null);
  const [layout, setLayout] = useState<Layout | null>(null);
  const margins = doc.page?.margins;
  const width = doc.page?.widthPt ?? 595.3;
  const height = doc.page?.heightPt ?? 841.9;

  useEffect(() => {
    const container = measureRef.current;
    if (!container) return;
    let frame = 0;
    let disposed = false;
    function measure() {
      if (disposed) return;
      const body = container!.querySelector<HTMLElement>(
        '[data-measure-body] [data-docx-body]',
      );
      if (!body) return;
      const outerHeight = (element: Element) => {
        const style = getComputedStyle(element);
        return (
          (element.getBoundingClientRect().height +
            (parseFloat(style.marginTop) || 0) +
            (parseFloat(style.marginBottom) || 0)) *
          0.75
        );
      };
      const measurements: BlockMeasurement[] = Array.from(body.children).map(
        (element) => {
          const rows =
            element.tagName === 'TABLE'
              ? Array.from(
                  element.querySelectorAll<HTMLTableRowElement>(
                    ':scope > tbody > tr',
                  ),
                )
              : undefined;
          return {
            heightPt: outerHeight(element),
            rowHeightsPt: rows?.map(
              (row) => row.getBoundingClientRect().height * 0.75,
            ),
            cells: rows?.map((row) =>
              Array.from(row.cells).map((cell) => {
                const style = getComputedStyle(cell);
                const cellTop = cell.getBoundingClientRect().top;
                return {
                  insetPt:
                    ((parseFloat(style.paddingTop) || 0) +
                      (parseFloat(style.paddingBottom) || 0) +
                      ((parseFloat(style.borderTopWidth) || 0) +
                        (parseFloat(style.borderBottomWidth) || 0)) /
                        2) *
                    0.75,
                  blocks: Array.from(cell.children).map((child) => {
                    const rect = child.getBoundingClientRect();
                    const childStyle = getComputedStyle(child);
                    return {
                      startPt:
                        (rect.top -
                          cellTop -
                          (parseFloat(childStyle.marginTop) || 0)) *
                        0.75,
                      endPt:
                        (rect.bottom -
                          cellTop +
                          (parseFloat(childStyle.marginBottom) || 0)) *
                        0.75,
                    };
                  }),
                };
              }),
            ),
          };
        },
      );
      const insets = {} as Insets;
      for (const variant of ['first', 'default', 'even'] as const) {
        const stories = container!.querySelector(
          `[data-measure-story="${variant}"]`,
        );
        const header = stories?.querySelector('[data-docx-story="header"]');
        const footer = stories?.querySelector('[data-docx-story="footer"]');
        insets[variant] = {
          top: Math.max(
            margins?.top ?? 72,
            header ? (margins?.header ?? 36) + outerHeight(header) : 0,
          ),
          bottom: Math.max(
            margins?.bottom ?? 72,
            footer ? (margins?.footer ?? 36) + outerHeight(footer) : 0,
          ),
        };
      }
      const pages = paginateBlocks(blocks, measurements, (number) => {
        const inset = insets[pageVariant(doc, number)];
        return height - inset.top - inset.bottom - 1;
      });
      setLayout({ source: doc, pages, insets });
    }
    const schedule = () => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(measure);
    };
    schedule();
    // Reflow after font loading, zoom/viewport changes and measured content resize.
    const observer = new ResizeObserver(schedule);
    observer.observe(container);
    const body = container.querySelector(
      '[data-measure-body] [data-docx-body]',
    );
    if (body) observer.observe(body);
    window.document.fonts.ready.then(() => {
      if (!disposed) schedule();
    });
    window.document.fonts.addEventListener('loadingdone', schedule);
    window.addEventListener('resize', schedule);
    return () => {
      disposed = true;
      cancelAnimationFrame(frame);
      observer.disconnect();
      window.document.fonts.removeEventListener('loadingdone', schedule);
      window.removeEventListener('resize', schedule);
    };
  }, [doc, blocks, fontRevision, height, margins]);

  const current = layout?.source === doc ? layout : null;
  useEffect(() => {
    if (current) onPageCountChange?.(current.pages.length);
  }, [current, onPageCountChange]);
  const storyDocument = (
    variant: 'first' | 'default' | 'even',
  ): DocumentIR => ({
    ...doc,
    blocks: [],
    page: { ...doc.page, differentFirstPage: false, evenAndOddHeaders: false },
    headers: { default: doc.headers?.[variant] },
    footers: { default: doc.footers?.[variant] },
  });
  const sheetStyle: CSSProperties = {
    boxSizing: 'content-box',
    width: `${width}pt`,
    boxShadow: '0 2px 12px #243c2c1a',
    border: '1px solid #d7ded9',
    background: doc.backgroundColor ?? '#fff',
    overflow: 'hidden',
  };
  return (
    <div
      className={className}
      data-docx-pages=""
      style={{ position: 'relative' }}
    >
      <div
        ref={measureRef}
        aria-hidden="true"
        style={{
          position: 'absolute',
          left: '-100000px',
          top: 0,
          width: `${width}pt`,
          visibility: 'hidden',
          pointerEvents: 'none',
        }}
      >
        <div data-measure-body="">
          <DocumentRenderer document={{ ...doc, blocks }} pageLayout />
        </div>
        {(['first', 'default', 'even'] as const).map((variant) => (
          <div key={variant} data-measure-story={variant}>
            <DocumentRenderer document={storyDocument(variant)} pageLayout />
          </div>
        ))}
      </div>
      {!current && (
        <div
          role="status"
          style={{
            ...sheetStyle,
            minHeight: `${height}pt`,
            padding: '60px',
            color: '#718178',
          }}
        >
          Laying out document pages…
        </div>
      )}
      {current?.pages.map((page, index) => {
        const inset = current.insets[pageVariant(doc, index + 1)];
        return (
          <section
            key={index}
            aria-label={`Page ${index + 1}`}
            data-docx-page={index + 1}
            style={{ width: 'fit-content', margin: '0 auto 28px' }}
          >
            <div style={sheetStyle}>
              <DocumentRenderer
                document={{ ...doc, blocks: page.blocks }}
                pageLayout
                pageNumber={index + 1}
                pageCount={current.pages.length}
                contentTopPt={inset.top}
                contentBottomPt={inset.bottom}
              />
            </div>
            <div
              style={{
                fontFamily: 'Segoe UI, Arial, sans-serif',
                fontSize: '11px',
                color: '#77877d',
                textAlign: 'center',
                marginTop: '9px',
              }}
            >
              Page {index + 1} of {current.pages.length}
              {page.overflow ? ' · Oversize content kept together' : ''}
            </div>
          </section>
        );
      })}
    </div>
  );
}
