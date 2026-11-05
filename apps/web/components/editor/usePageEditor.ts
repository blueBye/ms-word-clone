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
  selectionRanges,
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
  onCommands: (commands: EditCommand[]) => Promise<boolean>;
  onHistory: (direction: "undo" | "redo") => void;
  onError: (message: string) => void;
  onNotice: (message: string) => void;
}
/** A single editing host permits native selection across paragraphs; draft clones isolate React from browser edits. */
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
    startHint: number;
  } | null>(null);
  const typing = useRef<RunFormat>({}),
    composing = useRef(false),
    timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const inputStart = useRef(0);
  const structuralBusy = useRef(false),
    draining = useRef(false);
  const queued = useRef<
    { kind: "text" | "enter" | "backward" | "forward"; text?: string }[]
  >([]);
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
      startHint: Math.min(
        buffer.current?.id === id
          ? buffer.current.startHint
          : inputStart.current,
        inputStart.current,
      ),
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
        const change = textChange(
          paragraphText(p),
          snapshot.text,
          snapshot.startHint,
        );
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
      if (focus)
        latest.current.container.current?.focus({ preventScroll: true });
      return island.current;
    }
    const original = editableElement(selected.id);
    if (!original) return null;
    detach();
    const el = original.cloneNode(true) as HTMLElement;
    el.dataset.editorIsland = "";
    // Inherit the document's editing host. Nested editing hosts trap mouse selection.
    el.setAttribute("spellcheck", "false");
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
    if (focus) {
      latest.current.container.current?.focus({ preventScroll: true });
      restoreSelection(el, selected.start, selected.end);
    }
    return el;
  }
  function keyDown(event: KeyboardEvent) {
    if (waiting() && ["Enter", "Backspace", "Delete"].includes(event.key)) {
      event.preventDefault();
      queued.current.push({
        kind:
          event.key === "Enter"
            ? "enter"
            : event.key === "Backspace"
              ? "backward"
              : "forward",
      });
      return;
    }
    const multi = selectionRanges(
      latest.current.container.current!,
      latest.current.document,
    );
    if (
      multi.length > 1 &&
      ["Enter", "Backspace", "Delete"].includes(event.key)
    ) {
      event.preventDefault();
      void replaceAcross(multi, "")
        .then(() =>
          event.key === "Enter" ? structural("splitParagraph") : undefined,
        )
        .catch(report);
      return;
    }
    const el = island.current;
    if (!el) return;
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
        value =
          picked?.id === el.dataset.paragraphId
            ? editableText(el)
            : paragraphText(
                paragraphs(latest.current.document).find(
                  (p) => p.id === picked?.id,
                )!,
              );
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
        .then(() => latest.current.onHistory(event.shiftKey ? "redo" : "undo"))
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
  }
  async function structural(
    type: "splitParagraph" | "joinParagraph",
    direction?: "backward" | "forward",
  ) {
    structuralBusy.current = true;
    const picked = capture();
    if (!picked) {
      structuralBusy.current = false;
      return;
    }
    const continued = { ...typing.current };
    try {
      await flush();
      const command: EditCommand =
        type === "splitParagraph"
          ? {
              type,
              id: picked.id,
              start: picked.start,
              end: picked.end,
              format: continued,
            }
          : { type, id: picked.id, direction: direction! };
      if (await latest.current.onCommand(command)) {
        const target = latest.current.session?.selection;
        if (target) {
          selection.current = { ...target };
          setTyping(continued);
          focused.current = true;
          detach();
          restoring.current = true;
        }
      }
    } catch (error) {
      report(error);
    } finally {
      structuralBusy.current = false;
      requestAnimationFrame(() => void drain());
    }
  }
  function waiting() {
    return (
      structuralBusy.current ||
      latest.current.isBusy() ||
      !!latest.current.container.current?.querySelector(
        '[data-layout-pending="true"]',
      )
    );
  }
  async function drain() {
    if (draining.current || waiting()) return;
    draining.current = true;
    try {
      while (queued.current.length && !waiting()) {
        const next = queued.current.shift()!;
        if (next.kind === "enter") {
          await structural("splitParagraph");
          continue;
        }
        if (next.kind === "text") {
          insert(next.text ?? "");
          continue;
        }
        const picked = selection.current;
        if (!picked) continue;
        const p = paragraphs(latest.current.document).find(
          (p) => p.id === picked.id,
        );
        if (!p) continue;
        const value =
          island.current?.dataset.paragraphId === picked.id
            ? editableText(island.current)
            : paragraphText(p);
        let { start, end } = picked;
        if (start === end) {
          const edges = [
            ...Array.from(
              new Intl.Segmenter(undefined, {
                granularity: "grapheme",
              }).segment(value),
              (s) => s.index,
            ),
            value.length,
          ];
          if (next.kind === "backward")
            start = edges.filter((edge) => edge < start).at(-1) ?? start;
          else end = edges.find((edge) => edge > end) ?? end;
        }
        if (start === end) await structural("joinParagraph", next.kind);
        else {
          selection.current = { id: picked.id, start, end };
          focus();
          insert("");
        }
      }
    } catch (error) {
      report(error);
    } finally {
      draining.current = false;
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
    if (waiting()) {
      queued.current.push({ kind: "text", text });
      return;
    }
    const ranges = selectionRanges(
      latest.current.container.current!,
      latest.current.document,
    );
    if (ranges.length > 1) {
      void replaceAcross(ranges, text).catch(report);
      return;
    }
    const picked = capture() ?? selection.current;
    if (!picked) return;
    const el = ensureIsland(true);
    if (!el) return;
    inputStart.current = picked.start;
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
  async function replaceAcross(ranges: PageSelection[], text: string) {
    structuralBusy.current = true;
    try {
      await flush();
      const host = latest.current.container.current!;
      const elements = ranges.map((r) => editableElement(r.id));
      const commands: EditCommand[] = ranges.map((r, i) => ({
        type: "text",
        ...r,
        text: i === 0 ? text : "",
        format: i === 0 ? typing.current : undefined,
      }));
      // Merge only neighbours in the same body/cell/story; keep table boundaries intact.
      for (let i = 1; i < ranges.length; i++)
        if (
          elements[0] &&
          elements[i] &&
          elements[i]!.parentElement === elements[0]!.parentElement
        )
          commands.push({
            type: "joinParagraph",
            id: ranges[0].id,
            direction: "forward",
          });
      if (await latest.current.onCommands(commands)) {
        detach();
        select(
          {
            id: ranges[0].id,
            start: ranges[0].start + text.length,
            end: ranges[0].start + text.length,
          },
          true,
        );
        latest.current.onSelect(
          selection.current!,
          paragraphText(
            paragraphs(latest.current.session!.document).find(
              (p) => p.id === ranges[0].id,
            )!,
          ),
        );
        restoring.current = true;
        host.focus({ preventScroll: true });
      }
    } finally {
      structuralBusy.current = false;
      requestAnimationFrame(() => void drain());
    }
  }
  function focus() {
    if (window.document.querySelector("dialog[open], [data-home-menu-root]"))
      return;
    if (
      window.document.activeElement?.closest(
        ".task-pane input,.task-pane textarea,.task-pane select",
      )
    )
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
      if (el) {
        restoring.current = false;
        void drain();
      }
    });
    observer.observe(container, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["data-layout-pending"],
    });
    container.setAttribute("contenteditable", "plaintext-only");
    container.setAttribute("role", "textbox");
    container.setAttribute("aria-multiline", "true");
    container.tabIndex = 0;
    const activate = () => {
      const picked = readSelection(container);
      if (!picked) return; // Leave a multi-paragraph native selection intact.
      if (!latest.current.session?.capability(picked.id).text) return;
      const apply = () => {
        select(picked);
        const el = ensureIsland();
        if (el) restoreSelection(el, picked.start, picked.end);
        focused.current = true;
        capture(true);
      };
      if (buffer.current && buffer.current.id !== picked.id)
        void flush().then(apply).catch(report);
      else apply();
    };
    const down = (event: MouseEvent) => {
      if (event.defaultPrevented) return;
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
      // Let the browser place and extend the range. Activate the draft after mouseup.
    };
    const before = (event: InputEvent) => {
      if (waiting()) {
        event.preventDefault();
        if (event.inputType === "insertText" && event.data)
          queued.current.push({ kind: "text", text: event.data });
        else if (
          event.inputType === "deleteContentBackward" ||
          event.inputType === "deleteContentForward"
        )
          queued.current.push({
            kind: event.inputType.endsWith("Backward") ? "backward" : "forward",
          });
        return;
      }
      const ranges = selectionRanges(container, latest.current.document);
      if (ranges.length > 1 && !event.isComposing) {
        event.preventDefault();
        if (
          event.inputType.startsWith("delete") ||
          event.inputType === "insertText"
        )
          void replaceAcross(ranges, event.data ?? "").catch(report);
        return;
      }
      const picked = readSelection(container);
      if (!picked || !latest.current.session?.capability(picked.id).text) {
        event.preventDefault();
        return;
      }
      inputStart.current = picked.start;
      if (
        event.inputType === "deleteContentBackward" &&
        picked.start === picked.end
      ) {
        const p = paragraphs(latest.current.document).find(
          (p) => p.id === picked.id,
        );
        const value =
          island.current?.dataset.paragraphId === picked.id
            ? editableText(island.current)
            : p
              ? paragraphText(p)
              : "";
        inputStart.current =
          Array.from(
            new Intl.Segmenter(undefined, { granularity: "grapheme" }).segment(
              value.slice(0, picked.start),
            ),
          ).at(-1)?.index ?? 0;
      }
      if (buffer.current && buffer.current.id !== picked.id) {
        event.preventDefault();
        const text = event.data;
        void flush()
          .then(() => {
            select(picked, true);
            if (text) insert(text);
          })
          .catch(report);
        return;
      }
      if (island.current?.dataset.paragraphId !== picked.id) {
        select(picked);
        const el = ensureIsland();
        if (el) restoreSelection(el, picked.start, picked.end);
      }
      if (event.inputType === "insertParagraph") {
        event.preventDefault();
        void structural("splitParagraph");
      }
    };
    const pasted = (event: ClipboardEvent) => {
      const text = event.clipboardData?.getData("text/plain");
      if (text === undefined) return;
      event.preventDefault();
      const ranges = selectionRanges(container, latest.current.document);
      if (ranges.length > 1) void replaceAcross(ranges, text).catch(report);
      else insert(text);
    };
    const focusIn = () => {
      focused.current = true;
    };
    const focusOut = () => {
      focused.current = false;
      capture();
      void flush().catch(report);
    };
    const composeStart = () => {
      composing.current = true;
      clearTimer();
    };
    const composeEnd = () => {
      composing.current = false;
      input();
    };
    container.addEventListener("compositionstart", composeStart);
    container.addEventListener("compositionend", composeEnd);
    container.addEventListener("keydown", keyDown);
    container.addEventListener("mousedown", down);
    container.addEventListener("mouseup", activate);
    container.addEventListener("beforeinput", before);
    container.addEventListener("input", input);
    container.addEventListener("paste", pasted);
    container.addEventListener("focus", focusIn);
    container.addEventListener("blur", focusOut);
    return () => {
      observer.disconnect();
      container.removeEventListener("compositionstart", composeStart);
      container.removeEventListener("compositionend", composeEnd);
      container.removeEventListener("keydown", keyDown);
      container.removeEventListener("mousedown", down);
      container.removeEventListener("mouseup", activate);
      container.removeEventListener("beforeinput", before);
      container.removeEventListener("input", input);
      container.removeEventListener("paste", pasted);
      container.removeEventListener("focus", focusIn);
      container.removeEventListener("blur", focusOut);
      container.removeAttribute("contenteditable");
      clearTimer();
      detach();
    };
  }, []);
  useEffect(() => {
    queued.current = [];
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
