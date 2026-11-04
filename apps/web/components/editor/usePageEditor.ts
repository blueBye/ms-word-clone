import {
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type RefObject,
} from "react";
import type {
  DocumentSession,
  DocumentIR,
  EditCommand,
  RunFormat,
} from "@local/docx-ir-engine/core";
import {
  editableText,
  paragraphs,
  paragraphText,
  readSelection,
  restoreSelection,
  textChange,
} from "../../lib/editor";

export interface PageSelection {
  id: string;
  start: number;
  end: number;
}
interface Options {
  container: RefObject<HTMLElement | null>;
  session: DocumentSession | null;
  document: DocumentIR | null;
  isBusy: () => boolean;
  onCommit: (document: DocumentIR) => void;
  onSelect: (selection: PageSelection, text: string) => void;
  onCommand: (command: EditCommand) => Promise<boolean>;
  onHistory: (direction: "undo" | "redo") => void;
  onError: (message: string) => void;
  onNotice: (message: string) => void;
}
/** One native editing island. React owns the preview; it never reconciles browser-edited children. */
export function usePageEditor(options: Options) {
  const latest = useRef(options);
  latest.current = options;
  const selection = useRef<PageSelection | null>(null),
    island = useRef<HTMLElement | null>(null),
    source = useRef<HTMLElement | null>(null);
  const buffer = useRef<{
    id: string;
    text: string;
    version: number;
    format: RunFormat;
  } | null>(null);
  const typing = useRef<RunFormat>({}),
    composing = useRef(false),
    timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const committing = useRef<Promise<void> | null>(null),
    restoring = useRef(false),
    focused = useRef(false),
    version = useRef(0);
  const [dirty, setDirty] = useState(false),
    [typingFormat, setTypingFormat] = useState<RunFormat>({});

  function clearTimer() {
    if (timer.current) clearTimeout(timer.current);
    timer.current = null;
  }
  function detach() {
    const current = island.current;
    island.current = null;
    if (source.current) {
      source.current.style.display = source.current.dataset.editorDisplay ?? "";
      delete source.current.dataset.editorDisplay;
    }
    source.current = null;
    current?.remove();
  }
  function capture(manual = false) {
    const picked =
      latest.current.container.current &&
      readSelection(latest.current.container.current);
    if (!picked) return selection.current;
    if (
      manual &&
      (picked.id !== selection.current?.id ||
        picked.start !== selection.current?.start ||
        picked.end !== selection.current?.end)
    )
      setTyping({});
    selection.current = picked;
    const p = paragraphs(latest.current.document).find(
      (p) => p.id === picked.id,
    );
    latest.current.onSelect(
      picked,
      island.current?.dataset.paragraphId === picked.id
        ? editableText(island.current)
        : p
          ? paragraphText(p)
          : "",
    );
    return picked;
  }
  function setTyping(format: RunFormat) {
    typing.current = format;
    setTypingFormat(format);
  }
  function schedule() {
    clearTimer();
    timer.current = setTimeout(() => {
      void flush(false).catch(report);
    }, 450);
  }
  function report(error: unknown) {
    latest.current.onError(
      error instanceof Error ? error.message : "The edit could not be applied.",
    );
  }
  function input() {
    const el = island.current,
      id = el?.dataset.paragraphId;
    if (!el || !id) return;
    capture();
    buffer.current = {
      id,
      text: editableText(el),
      version: ++version.current,
      format: { ...typing.current },
    };
    setDirty(true);
    if (!composing.current) schedule();
  }
  async function flush(force = true): Promise<void> {
    clearTimer();
    if (committing.current) {
      await committing.current;
      if (buffer.current && !composing.current) return flush(force);
      return;
    }
    if (composing.current) return;
    if (!force && latest.current.isBusy()) {
      schedule();
      return;
    }
    const active = latest.current.session;
    if (!active || !buffer.current) return;
    const work = (async () => {
      while (buffer.current && !composing.current) {
        const snapshot = buffer.current,
          p = paragraphs(active.document).find((p) => p.id === snapshot.id);
        if (!p)
          throw new Error(
            "The active paragraph is no longer in this document.",
          );
        const change = textChange(paragraphText(p), snapshot.text);
        if (change.start !== change.end || change.text) {
          await active.execute({
            type: "text",
            id: snapshot.id,
            ...change,
            format: Object.keys(snapshot.format).length
              ? snapshot.format
              : undefined,
          });
          if (latest.current.session !== active) return;
          latest.current.onCommit(active.document);
        }
        if (buffer.current?.version === snapshot.version) buffer.current = null;
        setDirty(!!buffer.current);
      }
    })();
    committing.current = work;
    try {
      await work;
    } catch (error) {
      setDirty(!!buffer.current);
      throw error;
    } finally {
      committing.current = null;
    }
  }
  function editableElement(id: string) {
    const container = latest.current.container.current;
    if (!container || container.querySelector('[data-layout-pending="true"]'))
      return null;
    return container.querySelector<HTMLElement>(
      `[data-docx-page] [data-paragraph-id="${CSS.escape(id)}"]:not([data-editor-island])`,
    );
  }
  function ensureIsland(focus = false) {
    const selected = selection.current,
      active = latest.current.session;
    if (!selected || !active) return null;
    if (!active.capability(selected.id).text) return null;
    if (
      island.current?.isConnected &&
      island.current.dataset.paragraphId === selected.id
    ) {
      if (focus) island.current.focus({ preventScroll: true });
      return island.current;
    }
    const original = editableElement(selected.id);
    if (!original) return null;
    detach();
    const el = original.cloneNode(true) as HTMLElement;
    el.dataset.editorIsland = "";
    el.setAttribute("contenteditable", "plaintext-only");
    el.setAttribute("role", "textbox");
    el.setAttribute("aria-label", "Document paragraph");
    el.setAttribute("aria-multiline", "true");
    el.setAttribute("spellcheck", "false");
    el.tabIndex = 0;
    original.dataset.editorDisplay = original.style.display;
    original.style.display = "none";
    original.before(el);
    source.current = original;
    island.current = el;
    if (
      buffer.current?.id === selected.id &&
      editableText(el) !== buffer.current.text
    )
      el.textContent = buffer.current.text;
    el.addEventListener("input", input);
    el.addEventListener("compositionstart", () => {
      composing.current = true;
      clearTimer();
    });
    el.addEventListener("compositionend", () => {
      composing.current = false;
      input();
    });
    el.addEventListener("focus", () => {
      focused.current = true;
    });
    el.addEventListener("blur", () => {
      if (island.current !== el) return;
      focused.current = false;
      capture();
      void flush().catch(report);
    });
    el.addEventListener("keydown", (event) => {
      if (event.isComposing) return;
      if (event.key === "Enter" && !event.shiftKey) {
        event.preventDefault();
        void structural("splitParagraph");
      } else if (event.key === "Tab") {
        event.preventDefault();
        const picked = capture(),
          p = paragraphs(latest.current.document).find(
            (p) => p.id === picked?.id,
          );
        if (p?.properties.list && (picked?.start === 0 || event.shiftKey)) {
          void flush()
            .then(() =>
              latest.current.onCommand({
                type: "list",
                id: p.id!,
                kind: p.properties.list!.format as Extract<
                  EditCommand,
                  { type: "list" }
                >["kind"],
                level: Math.max(
                  0,
                  Math.min(
                    8,
                    (p.properties.numbering?.level ?? 0) +
                      (event.shiftKey ? -1 : 1),
                  ),
                ),
                preserveDefinition: true,
              }),
            )
            .catch(report);
        } else insert("\t");
      } else if (event.key === "Backspace" || event.key === "Delete") {
        const picked = capture(),
          value = editableText(el);
        if (
          picked &&
          picked.start === picked.end &&
          ((event.key === "Backspace" && picked.start === 0) ||
            (event.key === "Delete" && picked.end === value.length))
        ) {
          event.preventDefault();
          const p = paragraphs(latest.current.document).find(
            (p) => p.id === picked.id,
          );
          if (event.key === "Backspace" && p?.properties.list) {
            void flush()
              .then(() =>
                latest.current.onCommand({
                  type: "list",
                  id: picked.id,
                  kind: "none",
                }),
              )
              .catch(report);
            return;
          }
          void structural(
            "joinParagraph",
            event.key === "Backspace" ? "backward" : "forward",
          );
        }
      } else if (
        (event.ctrlKey || event.metaKey) &&
        event.key.toLowerCase() === "z"
      ) {
        event.preventDefault();
        void flush()
          .then(() =>
            latest.current.onHistory(event.shiftKey ? "redo" : "undo"),
          )
          .catch(report);
      }
      if (
        [
          "ArrowLeft",
          "ArrowRight",
          "ArrowUp",
          "ArrowDown",
          "Home",
          "End",
        ].includes(event.key)
      )
        queueMicrotask(() => capture(true));
    });
    if (focus) {
      el.focus({ preventScroll: true });
      restoreSelection(el, selected.start, selected.end);
    }
    return el;
  }
  async function structural(
    type: "splitParagraph" | "joinParagraph",
    direction?: "backward" | "forward",
  ) {
    const picked = capture();
    if (!picked) return;
    try {
      await flush();
      const command: EditCommand =
        type === "splitParagraph"
          ? { type, id: picked.id, start: picked.start, end: picked.end }
          : { type, id: picked.id, direction: direction! };
      if (await latest.current.onCommand(command)) {
        const target = latest.current.session?.selection;
        if (target) {
          selection.current = { ...target };
          setTyping({});
          focused.current = true;
          detach();
          restoring.current = true;
        }
      }
    } catch (error) {
      report(error);
    }
  }
  function select(value: PageSelection, focus = false) {
    if (value.id !== selection.current?.id) {
      detach();
      setTyping({});
    }
    selection.current = value;
    restoring.current = focus;
    if (focus) {
      focused.current = true;
      const el = ensureIsland(true);
      if (el) restoring.current = false;
    }
  }
  function insert(text: string) {
    const picked = capture() ?? selection.current;
    if (!picked) return;
    const el = ensureIsland(true);
    if (!el) return;
    restoreSelection(el, picked.start, picked.end);
    const range = window.getSelection()?.getRangeAt(0);
    if (!range) return;
    range.deleteContents();
    const node = window.document.createTextNode(text);
    range.insertNode(node);
    range.setStartAfter(node);
    range.collapse(true);
    const native = window.getSelection();
    native?.removeAllRanges();
    native?.addRange(range);
    input();
  }
  function focus() {
    if (window.document.querySelector("dialog[open], [data-home-menu-root]"))
      return;
    focused.current = true;
    restoring.current = true;
    const el = ensureIsland(true);
    if (el) restoring.current = false;
  }
  function refresh() {
    const shouldRestore = focused.current || restoring.current;
    detach();
    restoring.current = shouldRestore;
  }
  useLayoutEffect(() => {
    refresh();
  }, [options.document]);
  useEffect(() => {
    const container = options.container.current;
    if (!container) return;
    const observer = new MutationObserver(() => {
      if (!restoring.current) return;
      const active = window.document.activeElement;
      if (
        active?.closest("dialog[open]") ||
        active?.matches("input,select,textarea")
      )
        return;
      const el = ensureIsland(true);
      if (el) restoring.current = false;
    });
    observer.observe(container, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["data-layout-pending"],
    });
    const down = (event: MouseEvent) => {
      const target = (event.target as Element).closest<HTMLElement>(
        "[data-paragraph-id]",
      );
      if (
        !target?.closest("[data-docx-page]") ||
        target.hasAttribute("data-editor-island") ||
        target.closest("[data-editor-island]")
      )
        return;
      const id = target.dataset.paragraphId!;
      if (!latest.current.session?.capability(id).text) {
        latest.current.onNotice(
          latest.current.session?.capability(id).reason ?? "Read-only content.",
        );
        return;
      }
      if (buffer.current && buffer.current.id !== id) {
        event.preventDefault();
        void flush()
          .then(() => {
            select({ id, start: 0, end: 0 }, true);
            capture();
          })
          .catch(report);
        return;
      }
      event.preventDefault();
      select({ id, start: 0, end: 0 });
      const el = ensureIsland(true);
      if (!el) return;
      const doc = window.document as Document & {
        caretRangeFromPoint?: (x: number, y: number) => Range | null;
        caretPositionFromPoint?: (
          x: number,
          y: number,
        ) => { offsetNode: Node; offset: number } | null;
      };
      const hit = doc.caretRangeFromPoint?.(event.clientX, event.clientY),
        position = doc.caretPositionFromPoint?.(event.clientX, event.clientY);
      if (hit && el.contains(hit.startContainer)) {
        const s = window.getSelection();
        s?.removeAllRanges();
        s?.addRange(hit);
      } else if (position && el.contains(position.offsetNode)) {
        const r = window.document.createRange();
        r.setStart(position.offsetNode, position.offset);
        r.collapse(true);
        const s = window.getSelection();
        s?.removeAllRanges();
        s?.addRange(r);
      }
      capture(true);
    };
    container.addEventListener("mousedown", down);
    return () => {
      observer.disconnect();
      container.removeEventListener("mousedown", down);
      clearTimer();
      detach();
    };
  }, []);
  useEffect(() => {
    buffer.current = null;
    selection.current = null;
    setDirty(false);
    setTyping({});
    detach();
  }, [options.session]);
  return {
    dirty,
    typingFormat,
    setTyping,
    capture,
    select,
    focus,
    insert,
    flush,
    refresh,
    getSelection: () => selection.current,
    hasBuffer: () => !!buffer.current,
    isComposing: () => composing.current,
  };
}
