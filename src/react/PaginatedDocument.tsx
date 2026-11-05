"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import type { CSSProperties } from "react";
import type { DocumentIR, Block } from "../model/document.js";
import { DocumentRenderer } from "./DocumentRenderer.js";
import { measureBlocks, measureStoryInset } from "./measurement.js";
import { useLineNumbers } from "./useLineNumbers.js";
import { splitSections } from "../core/sections.js";
import {
  paginateBlocks,
  groupColumns,
  sectionPagePositions,
  pageVariant,
  splitPageBreaks,
  type PageSlice,
} from "./pagination.js";

export interface PaginatedDocumentProps {
  pageOffset?: number;
  totalPageCount?: number;
  fieldPageOffset?: number;
  document: DocumentIR;
  className?: string;
  /** Increment after registering a FontFace that was already loaded. */
  fontRevision?: number;
  onPageCountChange?: (count: number) => void;
  /** View-only zoom; measurement remains at physical page size. */
  zoom?: number;
  showFormattingMarks?: boolean;
}
type Insets = Record<
  "first" | "default" | "even",
  { top: number; bottom: number }
>;
interface Layout {
  source: DocumentIR;
  pages: (PageSlice & { columns?: Block[][] })[];
  insets: Insets;
}

export function PaginatedDocument(props: PaginatedDocumentProps) {
  const sections = useMemo(() => {
    const source = splitSections(props.document),
      result: DocumentIR[] = [];
    const geometry = (doc: DocumentIR) =>
      JSON.stringify([
        doc.page?.widthPt,
        doc.page?.heightPt,
        doc.page?.margins,
        doc.page?.columns,
        doc.page?.gutterPt,
        doc.page?.mirrorMargins,
        doc.page?.direction,
        doc.page?.verticalAlignment,
        doc.page?.pageNumberStart,
        doc.page?.differentFirstPage,
        doc.page?.evenAndOddHeaders,
        doc.page?.lineNumbers,
        doc.page?.hyphenation,
        doc.headers,
        doc.footers,
      ]);
    for (const section of source) {
      const previous = result.at(-1);
      if (
        previous &&
        section.page?.sectionStart === "continuous" &&
        geometry(previous) === geometry(section)
      )
        result[result.length - 1] = {
          ...previous,
          blocks: [...previous.blocks, ...section.blocks],
        };
      else result.push(section);
    }
    return result;
  }, [props.document]);
  return sections.length === 1 ? (
    <PaginatedSheetDocument {...props} document={sections[0]} />
  ) : (
    <SectionPages {...props} sections={sections} />
  );
}
function SectionPages({
  sections,
  ...props
}: PaginatedDocumentProps & { sections: DocumentIR[] }) {
  const [counts, setCounts] = useState<Record<number, number>>({});
  const { positions, total: offset } = sectionPagePositions(sections, counts);
  useEffect(
    () => props.onPageCountChange?.(offset),
    [props.onPageCountChange, offset],
  );
  return (
    <div data-docx-sections="">
      {sections.map((section, i) => (
        <div key={i}>
          {positions[i].blank && (
            <section
              data-docx-page={positions[i].before + 1}
              style={{
                width: "fit-content",
                margin: "0 auto 28px",
                zoom: props.zoom ?? 1,
                background: "white",
                boxShadow: "0 2px 12px #243c2c1a",
                border: "1px solid #d7ded9",
              }}
            >
              <DocumentRenderer
                document={{ ...sections[Math.max(0, i - 1)], blocks: [] }}
                pageLayout
                pageNumber={positions[i].before + 1}
                pageCount={offset}
              />
            </section>
          )}
          <PaginatedSheetDocument
            {...props}
            document={section}
            pageOffset={positions[i].offset}
            totalPageCount={offset}
            fieldPageOffset={positions[i].fieldOffset}
            onPageCountChange={(n) =>
              setCounts((previous) =>
                previous[i] === n ? previous : { ...previous, [i]: n },
              )
            }
          />
        </div>
      ))}
    </div>
  );
}
function PaginatedSheetDocument({
  document: doc,
  className,
  fontRevision = 0,
  onPageCountChange,
  zoom = 1,
  showFormattingMarks = false,
  pageOffset = 0,
  totalPageCount,
  fieldPageOffset = pageOffset,
}: PaginatedDocumentProps) {
  const blocks = useMemo(() => splitPageBreaks(doc.blocks), [doc]);
  const measureRef = useRef<HTMLDivElement>(null);
  const rootRef = useRef<HTMLDivElement>(null);
  const [layout, setLayout] = useState<Layout | null>(null);
  const lineNumbers = useLineNumbers(rootRef, doc, layout, zoom);
  const margins = doc.page?.margins;
  const width = doc.page?.widthPt ?? 595.3;
  const height = doc.page?.heightPt ?? 841.9;
  const columnCount = Math.max(1, Math.min(4, doc.page?.columns?.count ?? 1));
  const measurementWidth =
    columnCount === 1
      ? width
      : (width -
          (margins?.left ?? 72) -
          (margins?.right ?? 72) -
          (doc.page?.gutterPt ?? 0) -
          (columnCount - 1) * (doc.page?.columns?.gapPt ?? 36)) /
          columnCount +
        (margins?.left ?? 72) +
        (margins?.right ?? 72) +
        (doc.page?.gutterPt ?? 0);

  useEffect(() => {
    const container = measureRef.current;
    if (!container) return;
    let frame = 0;
    let disposed = false;
    function measure() {
      if (disposed) return;
      const body = container!.querySelector<HTMLElement>(
        "[data-measure-body] [data-docx-body]",
      );
      if (!body) return;
      const measurements = measureBlocks(body);
      const insets = {} as Insets;
      for (const variant of ["first", "default", "even"] as const) {
        const stories = container!.querySelector(
          `[data-measure-story="${variant}"]`,
        );
        const header = stories?.querySelector('[data-docx-story="header"]');
        const footer = stories?.querySelector('[data-docx-story="footer"]');
        insets[variant] = {
          top: Math.max(
            margins?.top ?? 72,
            measureStoryInset(header, "header"),
          ),
          bottom: Math.max(
            margins?.bottom ?? 72,
            measureStoryInset(footer, "footer"),
          ),
        };
      }
      const slices = paginateBlocks(blocks, measurements, (number) => {
        const inset = insets[pageVariant(doc, Math.ceil(number / columnCount))];
        return height - inset.top - inset.bottom - 1;
      });
      const pages = groupColumns(slices, columnCount);
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
      "[data-measure-body] [data-docx-body]",
    );
    if (body) observer.observe(body);
    window.document.fonts.ready.then(() => {
      if (!disposed) schedule();
    });
    window.document.fonts.addEventListener("loadingdone", schedule);
    window.addEventListener("resize", schedule);
    return () => {
      disposed = true;
      cancelAnimationFrame(frame);
      observer.disconnect();
      window.document.fonts.removeEventListener("loadingdone", schedule);
      window.removeEventListener("resize", schedule);
    };
  }, [doc, blocks, fontRevision, height, margins]);

  // Keep the previous sheets during measurement; the editing adapter restores
  // its caret after the new source revision has finished laying out.
  const current = layout;
  useEffect(() => {
    if (current) onPageCountChange?.(current.pages.length);
  }, [current, onPageCountChange]);
  const storyDocument = (
    variant: "first" | "default" | "even",
  ): DocumentIR => ({
    ...doc,
    blocks: [],
    page: { ...doc.page, differentFirstPage: false, evenAndOddHeaders: false },
    headers: { default: doc.headers?.[variant] },
    footers: { default: doc.footers?.[variant] },
  });
  const sheetStyle: CSSProperties = {
    boxSizing: "content-box",
    width: `${width}pt`,
    boxShadow: "0 2px 12px #243c2c1a",
    border: "1px solid #d7ded9",
    background: doc.backgroundColor ?? "#fff",
    overflow: "hidden",
  };
  return (
    <div
      ref={rootRef}
      className={className}
      data-docx-pages=""
      data-layout-pending={layout?.source !== doc ? "true" : undefined}
      style={{ position: "relative" }}
    >
      <div
        ref={measureRef}
        aria-hidden="true"
        contentEditable={false}
        style={{
          position: "absolute",
          left: "-100000px",
          top: 0,
          width: `${width}pt`,
          visibility: "hidden",
          pointerEvents: "none",
          height: 0,
          overflow: "hidden",
        }}
      >
        <div data-measure-body="">
          <DocumentRenderer
            document={{
              ...doc,
              page: { ...doc.page, widthPt: measurementWidth },
              blocks,
            }}
            pageLayout
          />
        </div>
        {(["first", "default", "even"] as const).map((variant) => (
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
            padding: "60px",
            color: "#718178",
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
            aria-label={`Page ${pageOffset + index + 1}`}
            data-docx-page={pageOffset + index + 1}
            data-docx-section-page={index + 1}
            style={{ width: "fit-content", margin: "0 auto 28px", zoom }}
          >
            <div style={{ ...sheetStyle, position: "relative" }}>
              <DocumentRenderer
                document={{
                  ...doc,
                  blocks: page.blocks,
                  page: {
                    ...doc.page,
                    differentFirstPage: false,
                    evenAndOddHeaders: false,
                  },
                  headers: {
                    default: doc.headers?.[pageVariant(doc, index + 1)],
                  },
                  footers: {
                    default: doc.footers?.[pageVariant(doc, index + 1)],
                  },
                }}
                pageLayout
                pageNumber={
                  doc.page?.pageNumberStart
                    ? index + 1
                    : fieldPageOffset + index + 1
                }
                pageCount={totalPageCount ?? current.pages.length}
                contentTopPt={inset.top}
                contentBottomPt={inset.bottom}
                columnBlocks={page.columns}
                showFormattingMarks={showFormattingMarks}
              />
              {lineNumbers
                .filter((line) => line.page === index + 1)
                .map((line, i) => (
                  <span
                    key={i}
                    contentEditable={false}
                    data-docx-line-number=""
                    style={{
                      position: "absolute",
                      top: `${line.top}pt`,
                      ...(doc.page?.direction === "rtl"
                        ? {
                            right: `${(doc.page?.margins?.right ?? 72) - (doc.page?.lineNumbers?.distancePt ?? 18)}pt`,
                          }
                        : {
                            left: `${(doc.page?.margins?.left ?? 72) - (doc.page?.lineNumbers?.distancePt ?? 18)}pt`,
                          }),
                      fontSize: "9pt",
                      color: "#777",
                      pointerEvents: "none",
                      userSelect: "none",
                    }}
                  >
                    {new Intl.NumberFormat(doc.language ?? "en").format(
                      line.number,
                    )}
                  </span>
                ))}
            </div>
            <div
              contentEditable={false}
              style={{
                fontFamily: "Segoe UI, Arial, sans-serif",
                fontSize: "11px",
                color: "#77877d",
                textAlign: "center",
                marginTop: "9px",
              }}
            >
              Page {pageOffset + index + 1} of{" "}
              {totalPageCount ?? current.pages.length}
              {page.overflow ? " · Oversize content kept together" : ""}
            </div>
          </section>
        );
      })}
    </div>
  );
}
