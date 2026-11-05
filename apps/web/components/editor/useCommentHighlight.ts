import { useEffect, type RefObject } from "react";
import { textRange } from "../../lib/editor";
import { commentSelection } from "../../lib/comments";
export function useCommentHighlight(
  root: RefObject<HTMLElement | null>,
  range: { id: string; start: number; end: number } | null,
) {
  useEffect(() => {
    if (!root.current || typeof Highlight === "undefined" || !CSS.highlights)
      return;
    const paint = () => {
      const ranges: Range[] = [];
      if (range)
        root
          .current!.querySelectorAll<HTMLElement>(
            `[data-docx-page] [data-paragraph-id="${CSS.escape(range.id)}"]`,
          )
          .forEach((p) => {
            if (p.style.display !== "none") {
              const selection = commentSelection(p.textContent ?? "", range);
              ranges.push(textRange(p, selection.start, selection.end));
            }
          });
      CSS.highlights.set("docx-comment-hover", new Highlight(...ranges));
    };
    paint();
    const observer = new MutationObserver(paint);
    observer.observe(root.current, {
      childList: true,
      subtree: true,
      characterData: true,
    });
    return () => {
      observer.disconnect();
      CSS.highlights.delete("docx-comment-hover");
    };
  }, [root, range]);
}
