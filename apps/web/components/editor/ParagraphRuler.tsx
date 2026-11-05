import {
  useLayoutEffect,
  useState,
  type PointerEvent,
  type RefObject,
} from "react";
import type {
  DocumentIR,
  Paragraph,
  ParagraphFormat,
} from "@local/docx-ir-engine/core";
import { rulerChange, type RulerHandle } from "../../lib/ruler";

export function ParagraphRuler({
  document,
  selected,
  zoom,
  disabled,
  onChange,
  container,
}: {
  document: DocumentIR;
  selected?: Paragraph;
  zoom: number;
  disabled: boolean;
  onChange: (format: ParagraphFormat) => void;
  container: RefObject<HTMLElement | null>;
}) {
  const [draft, setDraft] = useState<ParagraphFormat | null>(null);
  const [insets, setInsets] = useState<{ left: number; right: number } | null>(
    null,
  );
  useLayoutEffect(() => {
    const host = container.current;
    if (!host) return;
    const measure = () => {
      const node = selected?.id
        ? [
            ...host.querySelectorAll<HTMLElement>(
              `[data-docx-page] [data-paragraph-id="${CSS.escape(selected.id)}"]`,
            ),
          ].find((node) => node.style.display !== "none")
        : undefined;
      const parent = node?.closest<HTMLElement>("td,[data-docx-column]"),
        page = node?.closest<HTMLElement>("[data-docx-engine]");
      let next: typeof insets = null;
      if (parent && page) {
        const a = parent.getBoundingClientRect(),
          b = page.getBoundingClientRect(),
          width = document.page?.widthPt ?? 595.3,
          scale = b.width / width,
          style = getComputedStyle(parent);
        next = {
          left:
            (a.left - b.left) / scale +
            (parseFloat(style.paddingLeft) +
              parseFloat(style.borderLeftWidth) || 0) *
              0.75,
          right:
            width -
            (a.right - b.left) / scale +
            (parseFloat(style.paddingRight) +
              parseFloat(style.borderRightWidth) || 0) *
              0.75,
        };
      }
      setInsets((previous) =>
        JSON.stringify(previous) === JSON.stringify(next) ? previous : next,
      );
    };
    measure();
    const observer = new MutationObserver(measure);
    observer.observe(host, {
      subtree: true,
      attributes: true,
      attributeFilter: ["data-layout-pending"],
    });
    return () => observer.disconnect();
  }, [container, selected?.id, document, zoom]);
  const width = document.page?.widthPt ?? 595.3,
    left = insets?.left ?? document.page?.margins?.left ?? 72,
    right = insets?.right ?? document.page?.margins?.right ?? 72;
  const content = width - left - right,
    rtl = selected?.properties.direction === "rtl",
    p = { ...selected?.properties, ...draft };
  const start =
    p.indentStartPt ?? (rtl ? p.indentRightPt : p.indentLeftPt) ?? 0;
  const end = p.indentEndPt ?? (rtl ? p.indentLeftPt : p.indentRightPt) ?? 0;
  const first = p.firstLinePt ?? (p.list ? -18 : 0);
  const physical = (offset: number) =>
    rtl ? width - right - offset : left + offset;
  const change = (
    event: PointerEvent<HTMLButtonElement>,
    handle: RulerHandle,
  ) => {
    const rect = event.currentTarget.parentElement!.getBoundingClientRect();
    const point = ((event.clientX - rect.left) / rect.width) * width;
    const value =
      handle === "end"
        ? rtl
          ? point - left
          : width - right - point
        : rtl
          ? width - right - point
          : point - left;
    return rulerChange(handle, value, start, first, content);
  };
  return (
    <div
      className="paragraph-ruler"
      contentEditable={false}
      role="group"
      aria-label="Paragraph ruler"
      style={{ width: `${((width * 4) / 3) * zoom}px` }}
    >
      <div
        className="ruler-margin"
        style={{ width: `${(left / width) * 100}%` }}
      />
      <div
        className="ruler-margin ruler-margin-right"
        style={{ width: `${(right / width) * 100}%` }}
      />
      {Array.from({ length: Math.floor(content / 36) + 1 }, (_, i) => (
        <span
          className={`ruler-tick ${i % 2 ? "minor" : ""}`}
          key={i}
          style={{ left: `${(physical(i * 36) / width) * 100}%` }}
        >
          {i % 2 ? "" : i / 2}
        </span>
      ))}
      {(
        [
          [
            "first",
            p.list ? "Bullet or number position" : "First line indent",
            physical(start + first),
          ],
          ["hanging", "Hanging indent", physical(start)],
          ["start", "Move paragraph indent", physical(start)],
          [
            "end",
            "Paragraph end indent",
            rtl ? left + end : width - right - end,
          ],
        ] as const
      ).map(([handle, label, position]) => (
        <button
          key={handle}
          className={`ruler-handle ruler-${handle}`}
          type="button"
          aria-label={label}
          title={`${label}: ${Math.round(handle === "first" ? first : handle === "end" ? end : start)} pt — drag or use arrow keys`}
          disabled={disabled || !selected}
          style={{ left: `${(position / width) * 100}%` }}
          onPointerDown={(e) => {
            e.preventDefault();
            e.currentTarget.setPointerCapture(e.pointerId);
          }}
          onPointerMove={(e) => {
            if (e.currentTarget.hasPointerCapture(e.pointerId))
              setDraft(change(e, handle));
          }}
          onPointerUp={(e) => {
            if (!e.currentTarget.hasPointerCapture(e.pointerId)) return;
            const format = change(e, handle);
            e.currentTarget.releasePointerCapture(e.pointerId);
            setDraft(null);
            onChange(format);
          }}
          onPointerCancel={() => setDraft(null)}
          onKeyDown={(e) => {
            if (e.key !== "ArrowLeft" && e.key !== "ArrowRight") return;
            e.preventDefault();
            const delta =
              (e.key === "ArrowRight" ? 1 : -1) *
              (rtl ? -1 : 1) *
              (e.shiftKey ? 6 : 0.5);
            const value =
              handle === "first"
                ? start + first
                : handle === "end"
                  ? end
                  : start;
            onChange(
              rulerChange(
                handle,
                value + (handle === "end" ? -delta : delta),
                start,
                first,
                content,
              ),
            );
          }}
        />
      ))}
    </div>
  );
}
