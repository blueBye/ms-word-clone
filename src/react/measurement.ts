import type { BlockMeasurement } from "./pagination.js";

const PT_PER_PX = 0.75;
const number = (value: string) => parseFloat(value) || 0;

export function measureBlocks(body: Element): BlockMeasurement[] {
  const elements = Array.from(body.children);
  const measurements = elements.map((element) => {
    const style = getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    const rows =
      element.tagName === "TABLE"
        ? Array.from(
            element.querySelectorAll<HTMLTableRowElement>(
              ":scope > tbody > tr",
            ),
          )
        : undefined;
    return {
      heightPt:
        (rect.height + number(style.marginTop) + number(style.marginBottom)) *
        PT_PER_PX,
      tableWidthPt: rows ? rect.width * PT_PER_PX : undefined,
      columnWidthsPt: rows
        ? Array.from(element.querySelectorAll(":scope > colgroup > col")).map(
            (col) => col.getBoundingClientRect().width * PT_PER_PX,
          )
        : undefined,
      rowHeightsPt: rows?.map(
        (row) => row.getBoundingClientRect().height * PT_PER_PX,
      ),
      cells: rows?.map((row) =>
        Array.from(row.cells).map((cell) => {
          const style = getComputedStyle(cell);
          const cellTop = cell.getBoundingClientRect().top;
          return {
            insetPt:
              (number(style.paddingTop) +
                number(style.paddingBottom) +
                (number(style.borderTopWidth) +
                  number(style.borderBottomWidth)) /
                  2) *
              PT_PER_PX,
            blocks: Array.from(cell.children).map((child) => {
              const rect = child.getBoundingClientRect();
              const style = getComputedStyle(child);
              return {
                startPt:
                  (rect.top - cellTop - number(style.marginTop)) * PT_PER_PX,
                endPt:
                  (rect.bottom - cellTop + number(style.marginBottom)) *
                  PT_PER_PX,
              };
            }),
          };
        }),
      ),
    };
  });
  elements.forEach((element, i) => {
    if (getComputedStyle(element).float !== "none" && measurements[i + 1]) {
      measurements[i + 1].heightPt = Math.max(
        measurements[i].heightPt,
        measurements[i + 1].heightPt,
      );
      measurements[i].heightPt = 0;
    }
  });
  return measurements;
}

/** Floating drawings/textboxes can extend beyond their story's flow box. */
export function measureStoryInset(
  story: Element | null | undefined,
  kind: "header" | "footer",
): number {
  if (!story?.parentElement) return 0;
  const page = story.parentElement.getBoundingClientRect();
  const flow = story.getBoundingClientRect();
  let top = flow.top;
  let bottom = flow.bottom;
  for (const element of Array.from(
    story.querySelectorAll("[data-docx-paragraph], [data-docx-drawing], table"),
  )) {
    if (element.closest('[data-docx-behind-text="true"]')) continue;
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    if (style.display === "none" || (!rect.height && !rect.width)) continue;
    top = Math.min(top, rect.top - number(style.marginTop));
    bottom = Math.max(bottom, rect.bottom + number(style.marginBottom));
  }
  return (
    Math.max(0, kind === "header" ? bottom - page.top : page.bottom - top) *
    PT_PER_PX
  );
}
