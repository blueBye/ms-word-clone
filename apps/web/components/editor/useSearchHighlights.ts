import { useEffect, type RefObject } from "react";
import { textRange } from "../../lib/editor";
import type { SearchHit } from "../../lib/search";

/** Highlight ranges without inserting markup into the editor or the measured document. */
export function useSearchHighlights(
  container: RefObject<HTMLElement | null>,
  hits: SearchHit[],
  active: SearchHit | null,
  enabled: boolean,
  scope: { id: string; start: number; end: number }[],
) {
  useEffect(() => {
    const root = container.current;
    if (!root || typeof Highlight === "undefined" || !CSS.highlights) return;
    let frame = 0;
    const paint = () => {
      frame = 0;
      const ranges: Range[] = [],
        selected: Range[] = [];
      if (enabled) {
        const byId = new Map<string, SearchHit[]>();
        for (const hit of hits)
          byId.set(hit.id, [...(byId.get(hit.id) ?? []), hit]);
        root
          .querySelectorAll<HTMLElement>("[data-docx-page] [data-paragraph-id]")
          .forEach((node) => {
            if (node.style.display === "none") return;
            for (const hit of byId.get(node.dataset.paragraphId!) ?? []) {
              const range = textRange(node, hit.start, hit.end);
              ranges.push(range);
              if (
                active?.id === hit.id &&
                active.start === hit.start &&
                active.end === hit.end
              )
                selected.push(range);
            }
          });
      }
      CSS.highlights.set("docx-search", new Highlight(...ranges));
      CSS.highlights.set("docx-search-active", new Highlight(...selected));
      const picked: Range[] = [];
      for (const r of scope)
        root
          .querySelectorAll<HTMLElement>(
            `[data-docx-page] [data-paragraph-id="${CSS.escape(r.id)}"]`,
          )
          .forEach((node) => {
            if (node.style.display !== "none")
              picked.push(textRange(node, r.start, r.end));
          });
      CSS.highlights.set("docx-selection", new Highlight(...picked));
    };
    const schedule = () => {
      if (!frame) frame = requestAnimationFrame(paint);
    };
    paint();
    const observer = new MutationObserver(schedule);
    observer.observe(root, {
      childList: true,
      subtree: true,
      characterData: true,
      attributes: true,
      attributeFilter: ["style", "data-layout-pending"],
    });
    return () => {
      observer.disconnect();
      cancelAnimationFrame(frame);
      CSS.highlights.delete("docx-search");
      CSS.highlights.delete("docx-search-active");
      CSS.highlights.delete("docx-selection");
    };
  }, [container, hits, active, enabled, scope]);
}
