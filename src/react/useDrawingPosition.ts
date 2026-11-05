import { useLayoutEffect, useRef } from "react";
import type {
  DrawingInline,
  DrawingPosition,
  PageLayout,
} from "../model/document.js";

/** Resolve DrawingML page/margin/column coordinates against the actual paginated frame. */
export function useDrawingPosition(
  drawing: DrawingInline,
  page?: PageLayout,
  enabled = true,
) {
  const ref = useRef<HTMLElement | null>(null);
  useLayoutEffect(() => {
    const node = ref.current,
      anchor = drawing.position;
    if (!node || !anchor || !enabled) return;
    const sheet = node.closest<HTMLElement>("[data-docx-engine]"),
      paragraph = node.closest<HTMLElement>("[data-docx-paragraph]");
    if (!sheet || !paragraph) return;
    let frame = 0;
    const measure = () => {
      frame = 0;
      const parent = node.offsetParent as HTMLElement | null;
      if (!parent) return;
      const bounds = sheet.getBoundingClientRect(),
        base = parent.getBoundingClientRect(),
        p = paragraph.getBoundingClientRect();
      const scale = bounds.width / (page?.widthPt ?? 595.3);
      if (!scale) return;
      const column =
        node.closest<HTMLElement>("[data-docx-column]") ??
        node.closest<HTMLElement>("[data-docx-body]");
      const c = column?.getBoundingClientRect() ?? p;
      const margins = page?.margins ?? {},
        left = (margins.left ?? 72) * scale,
        right = (margins.right ?? 72) * scale,
        top = (margins.top ?? 72) * scale,
        bottom = (margins.bottom ?? 72) * scale;
      const horizontal = anchor.horizontal,
        vertical = anchor.vertical;
      const h =
        horizontal.relativeTo === "page"
          ? { origin: bounds.left, size: bounds.width }
          : ["margin", "leftMargin", "rightMargin"].includes(
                horizontal.relativeTo,
              )
            ? {
                origin:
                  horizontal.relativeTo === "rightMargin"
                    ? bounds.right - right
                    : bounds.left +
                      (horizontal.relativeTo === "leftMargin" ? 0 : left),
                size:
                  horizontal.relativeTo === "leftMargin"
                    ? left
                    : horizontal.relativeTo === "rightMargin"
                      ? right
                      : bounds.width - left - right,
              }
            : horizontal.relativeTo === "column"
              ? { origin: c.left, size: c.width }
              : { origin: p.left, size: p.width };
      const v =
        vertical.relativeTo === "page"
          ? { origin: bounds.top, size: bounds.height }
          : vertical.relativeTo === "margin"
            ? { origin: bounds.top + top, size: bounds.height - top - bottom }
            : { origin: p.top, size: p.height };
      const coordinate = (
        axis: DrawingPosition,
        area: { origin: number; size: number },
        size: number,
      ) =>
        area.origin +
        (axis.alignment === "center" || axis.alignment === "middle"
          ? (area.size - size * scale) / 2
          : axis.alignment === "right" || axis.alignment === "bottom"
            ? area.size - size * scale
            : (axis.offsetPt ?? 0) * scale);
      const computed = getComputedStyle(parent),
        zoom = scale / (4 / 3);
      const x =
        (coordinate(horizontal, h, drawing.widthPt) -
          base.left -
          parseFloat(computed.borderLeftWidth) * zoom) /
        scale;
      const y =
        (coordinate(vertical, v, drawing.heightPt) -
          base.top -
          parseFloat(computed.borderTopWidth) * zoom) /
        scale;
      const apply = (key: "left" | "top", value: number) => {
        const css = `${Math.round(value * 1000) / 1000}pt`;
        if (node.style[key] !== css) node.style[key] = css;
      };
      node.style.right = "";
      node.style.marginLeft = "";
      apply("left", x);
      apply("top", y);
    };
    const schedule = () => {
      if (!frame) frame = requestAnimationFrame(measure);
    };
    measure();
    const resize = new ResizeObserver(schedule);
    resize.observe(sheet);
    const observer = new MutationObserver(schedule);
    observer.observe(sheet, {
      subtree: true,
      childList: true,
      characterData: true,
      attributes: true,
      attributeFilter: ["data-layout-pending"],
    });
    return () => {
      resize.disconnect();
      observer.disconnect();
      cancelAnimationFrame(frame);
    };
  }, [drawing.position, drawing.widthPt, drawing.heightPt, page, enabled]);
  return ref;
}
