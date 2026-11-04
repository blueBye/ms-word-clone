import {
  getParagraphs,
  type DocumentIR,
  type Paragraph,
  type DocxStyle,
  type EditCommand,
  type RunFormat,
  type ParagraphFormat,
} from "@local/docx-ir-engine";

export function paragraphs(document: DocumentIR | null): Paragraph[] {
  if (!document) return [];
  return [
    ...getParagraphs(document.blocks),
    ...Object.values(document.headers ?? {}).flatMap((b) =>
      getParagraphs(b ?? []),
    ),
    ...Object.values(document.footers ?? {}).flatMap((b) =>
      getParagraphs(b ?? []),
    ),
  ].filter(
    (p, index, all) => all.findIndex((other) => other.id === p.id) === index,
  );
}
export function paragraphText(p: Paragraph) {
  return p.children
    .filter((c) => c.type === "text")
    .map((c) => c.text)
    .join("");
}

/** Write only edited fields so inherited settings remain linked to the base style. */
export function styleFormCommand(
  style: DocxStyle,
  data: FormData,
): EditCommand {
  const run: RunFormat = {
    fontFamily: String(data.get("font")),
    fontSizePt: Number(data.get("size")),
    color: String(data.get("color")),
    bold: data.has("bold"),
    italic: data.has("italic"),
  };
  const initialRun: RunFormat = {
    fontFamily: style.run.fontFamily ?? "Calibri",
    fontSizePt: style.run.fontSizePt ?? 12,
    color: style.run.color ?? "#000000",
    bold: !!style.run.bold,
    italic: !!style.run.italic,
  };
  const paragraph: ParagraphFormat = {
    alignment: String(data.get("align")) as ParagraphFormat["alignment"],
    spaceBeforePt: Number(data.get("before")),
    spaceAfterPt: Number(data.get("after")),
  };
  const initialParagraph: ParagraphFormat = {
    alignment: (style.paragraph.alignment ??
      "left") as ParagraphFormat["alignment"],
    spaceBeforePt: style.paragraph.spaceBeforePt ?? 0,
    spaceAfterPt: style.paragraph.spaceAfterPt ?? 0,
  };
  const changed = (value: object, initial: object) =>
    Object.fromEntries(
      Object.entries(value).filter(
        ([key, item]) =>
          String(item).toLowerCase() !==
          String((initial as Record<string, unknown>)[key]).toLowerCase(),
      ),
    );
  return {
    type: "style",
    styleId: style.id,
    name: String(data.get("name")),
    ...(String(data.get("basedOn")) !== (style.basedOn ?? "")
      ? { basedOn: String(data.get("basedOn")) }
      : {}),
    gallery: data.has("gallery"),
    ...(Object.keys(changed(run, initialRun)).length
      ? { run: changed(run, initialRun) }
      : {}),
    ...(Object.keys(changed(paragraph, initialParagraph)).length
      ? { paragraph: changed(paragraph, initialParagraph) }
      : {}),
  };
}
/** Expand text changes to grapheme boundaries and leave unaffected runs intact. */
export function textChange(before: string, after: string) {
  const segmenter = new Intl.Segmenter(undefined, { granularity: "grapheme" });
  const a = Array.from(segmenter.segment(before), (s) => s.segment),
    b = Array.from(segmenter.segment(after), (s) => s.segment);
  let prefix = 0,
    suffix = 0;
  while (prefix < a.length && prefix < b.length && a[prefix] === b[prefix])
    prefix++;
  while (
    suffix < a.length - prefix &&
    suffix < b.length - prefix &&
    a[a.length - 1 - suffix] === b[b.length - 1 - suffix]
  )
    suffix++;
  return {
    start: a.slice(0, prefix).join("").length,
    end: before.length - a.slice(a.length - suffix).join("").length,
    text: b.slice(prefix, b.length - suffix).join(""),
  };
}
export function download(bytes: Uint8Array, filename: string) {
  const blob = new Blob([new Uint8Array(bytes)], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  const url = URL.createObjectURL(blob),
    anchor = window.document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

export function readSelection(container: HTMLElement) {
  const selection = window.getSelection();
  if (!selection?.rangeCount) return null;
  const range = selection.getRangeAt(0);
  const parent = (node: Node) =>
    (node.nodeType === 1
      ? (node as Element)
      : node.parentElement
    )?.closest<HTMLElement>("[data-paragraph-id]");
  const start = parent(range.startContainer),
    end = parent(range.endContainer);
  if (
    !start ||
    start !== end ||
    !container.contains(start) ||
    !start.closest("[data-docx-page]")
  )
    return null;
  const offset = (node: Node, index: number) => {
    const prefix = window.document.createRange();
    prefix.selectNodeContents(start);
    prefix.setEnd(node, index);
    const fragment = prefix.cloneContents();
    fragment.querySelectorAll("[data-docx-marker]").forEach((e) => e.remove());
    return Math.min(
      editableText(start).length,
      editableText(fragment, false).length,
    );
  };
  return {
    id: start.dataset.paragraphId!,
    start: offset(range.startContainer, range.startOffset),
    end: offset(range.endContainer, range.endOffset),
  };
}

/** Read the editing island's logical text; view-only list markers never enter the model. */
export function editableText(root: Node, trimFiller = true): string {
  const read = (node: Node): string => {
    if (node.nodeType === 3) return node.textContent ?? "";
    if (node.nodeType !== 1 && node.nodeType !== 11) return "";
    const el = node as Element;
    if (
      el.nodeType === 1 &&
      (el.hasAttribute("data-docx-marker") ||
        el.hasAttribute("data-docx-placeholder") ||
        el.hasAttribute("data-docx-mark"))
    )
      return "";
    if (el.nodeName === "BR") return "\n";
    return Array.from(node.childNodes)
      .map((child, index) => {
        const value = read(child),
          block = child.nodeType === 1 && ["DIV", "P"].includes(child.nodeName);
        return (block && index > 0 ? "\n" : "") + value;
      })
      .join("");
  };
  let result = read(root);
  if (
    trimFiller &&
    root.lastChild?.nodeName === "BR" &&
    !(root.lastChild as Element).hasAttribute("data-docx-placeholder")
  )
    result = result.slice(0, -1);
  return result;
}

export function textRange(element: HTMLElement, start: number, end: number) {
  const point = (position: number): { node: Node; offset: number } => {
    const walker = window.document.createTreeWalker(
      element,
      NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT,
    );
    let node: Node | null,
      remaining = position;
    while ((node = walker.nextNode())) {
      const parent =
        node.nodeType === 1 ? (node as Element) : node.parentElement;
      if (
        parent?.closest(
          "[data-docx-marker],[data-docx-placeholder],[data-docx-mark]",
        )
      )
        continue;
      if (node.nodeType === 3) {
        const count = node.textContent?.length ?? 0;
        if (remaining <= count) return { node, offset: remaining };
        remaining -= count;
      } else if (node.nodeName === "BR") {
        if (remaining === 0)
          return {
            node: node.parentNode!,
            offset: Array.from(node.parentNode!.childNodes).indexOf(
              node as ChildNode,
            ),
          };
        remaining--;
      }
    }
    return { node: element, offset: element.childNodes.length };
  };
  const a = point(start),
    b = point(end),
    range = window.document.createRange();
  range.setStart(a.node, a.offset);
  range.setEnd(b.node, b.offset);
  return range;
}

export function restoreSelection(
  element: HTMLElement,
  start: number,
  end: number,
) {
  const range = textRange(element, start, end);
  const selection = window.getSelection();
  selection?.removeAllRanges();
  selection?.addRange(range);
}

export function selectionRanges(
  container: HTMLElement,
  document: DocumentIR | null,
) {
  const single = readSelection(container);
  if (single) return [single];
  const selection = window.getSelection();
  if (!selection?.rangeCount) return [];
  const native = selection.getRangeAt(0);
  const host = (node: Node) =>
    (node.nodeType === 1
      ? (node as Element)
      : node.parentElement
    )?.closest<HTMLElement>("[data-paragraph-id]");
  const first = host(native.startContainer),
    last = host(native.endContainer);
  if (
    !first ||
    !last ||
    !container.contains(first) ||
    !container.contains(last)
  )
    return [];
  const unique = new Map<string, HTMLElement>();
  container
    .querySelectorAll<HTMLElement>("[data-docx-page] [data-paragraph-id]")
    .forEach((node) => {
      if (
        node.style.display !== "none" &&
        !unique.has(node.dataset.paragraphId!)
      )
        unique.set(node.dataset.paragraphId!, node);
    });
  const ids = [...unique.keys()],
    startIndex = ids.indexOf(first.dataset.paragraphId!),
    endIndex = ids.indexOf(last.dataset.paragraphId!);
  if (startIndex < 0 || endIndex < startIndex) return [];
  const model = new Map(
    paragraphs(document).map((p) => [p.id, paragraphText(p)]),
  );
  const offset = (element: HTMLElement, node: Node, index: number) => {
    const r = window.document.createRange();
    r.selectNodeContents(element);
    r.setEnd(node, index);
    return editableText(r.cloneContents(), false).length;
  };
  return ids.slice(startIndex, endIndex + 1).map((id, index, list) => ({
    id,
    start:
      index === 0
        ? offset(first, native.startContainer, native.startOffset)
        : 0,
    end:
      index === list.length - 1
        ? offset(last, native.endContainer, native.endOffset)
        : (model.get(id)?.length ?? 0),
  }));
}

export function changeCase(text: string, mode: string, locale?: string) {
  if (mode === "upper") return text.toLocaleUpperCase(locale);
  if (mode === "lower") return text.toLocaleLowerCase(locale);
  if (mode === "toggle")
    return Array.from(text, (c) =>
      c === c.toLocaleUpperCase(locale)
        ? c.toLocaleLowerCase(locale)
        : c.toLocaleUpperCase(locale),
    ).join("");
  if (mode === "title")
    return Array.from(
      new Intl.Segmenter(locale, { granularity: "word" }).segment(text),
      (part) =>
        part.isWordLike
          ? part.segment[0].toLocaleUpperCase(locale) +
            part.segment.slice(1).toLocaleLowerCase(locale)
          : part.segment,
    ).join("");
  if (mode === "sentence")
    return text
      .toLocaleLowerCase(locale)
      .replace(
        /(^|[.!?؟]\s+)(\p{L})/gu,
        (_, prefix: string, letter: string) =>
          prefix + letter.toLocaleUpperCase(locale),
      );
  return text;
}
