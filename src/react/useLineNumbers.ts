import { useLayoutEffect, useState, type RefObject } from "react";
import type { DocumentIR } from "../model/document.js";
export function useLineNumbers(
  root: RefObject<HTMLElement | null>,
  document: DocumentIR,
  revision: unknown,
  zoom: number,
) {
  const [lines, setLines] = useState<
    { page: number; top: number; number: number }[]
  >([]);
  useLayoutEffect(() => {
    if (!document.page?.lineNumbers || !root.current) {
      setLines([]);
      return;
    }
    const settings = document.page.lineNumbers,
      result: typeof lines = [];
    let counter = settings.start;
    root.current
      .querySelectorAll<HTMLElement>("[data-docx-page]")
      .forEach((page, index) => {
        const engine = page.querySelector<HTMLElement>("[data-docx-engine]");
        if (!engine) return;
        const rect = engine.getBoundingClientRect(),
          scale = rect.width / (document.page?.widthPt ?? 595.3);
        if (settings.mode === "page") counter = settings.start;
        page
          .querySelectorAll<HTMLElement>("[data-docx-body] [data-paragraph-id]")
          .forEach((p) => {
            if (
              p.dataset.suppressLineNumbers ||
              p.style.display === "none" ||
              p.closest("[data-docx-drawing],table")
            )
              return;
            const boxes: { top: number; bottom: number }[] = [],
              walker = window.document.createTreeWalker(
                p,
                NodeFilter.SHOW_TEXT,
              );
            let node: Node | null;
            while ((node = walker.nextNode())) {
              if (
                node.parentElement?.closest(
                  "[data-docx-marker],[data-docx-mark]",
                )
              )
                continue;
              const range = window.document.createRange();
              range.selectNodeContents(node);
              for (const r of Array.from(range.getClientRects()))
                if (
                  r.height &&
                  r.width &&
                  !boxes.some(
                    (b) =>
                      Math.min(b.bottom, r.bottom) - Math.max(b.top, r.top) >
                      Math.min(b.bottom - b.top, r.height) / 2,
                  )
                )
                  boxes.push({ top: r.top, bottom: r.bottom });
            }
            if (!boxes.length) boxes.push(p.getBoundingClientRect());
            boxes
              .sort((a, b) => a.top - b.top)
              .forEach((box) => {
                if ((counter - settings.start) % settings.countBy === 0)
                  result.push({
                    page: index + 1,
                    top: (box.top - rect.top) / scale,
                    number: counter,
                  });
                counter++;
              });
          });
      });
    setLines(result);
  }, [root, document, revision, zoom]);
  return lines;
}
