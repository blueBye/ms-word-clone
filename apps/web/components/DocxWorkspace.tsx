"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import {
  type DocumentStyle,
  type DocumentIR,
  type EditCommand,
  type ParagraphFormat,
  type RunFormat,
} from "@local/docx-ir-engine/core";
import {
  openDocxSession,
  createBlankDocx,
  type DocxSession,
} from "@local/docx-ir-engine/docx";
import { PaginatedDocument } from "@local/docx-ir-engine/react";
import {
  paragraphs,
  paragraphText,
  textChange,
  download,
  selectionRanges,
  changeCase,
  restoreSelection,
} from "../lib/editor";
import {
  searchMatches,
  nextMatchIndex,
  type SearchHit,
  type SearchOptions,
  type SearchDirection,
} from "../lib/search";
import { useSearchHighlights } from "./editor/useSearchHighlights";
import { Ribbon, type RibbonTab } from "./editor/Ribbon";
import { Tool } from "./editor/Controls";
import { Icon } from "./editor/Icon";
import { FontPanel, useDocumentFonts } from "./editor/FontPanel";
import { WorkspaceDialogs, type EditorModal } from "./editor/WorkspaceDialogs";
import { Backstage } from "./editor/Backstage";
import { usePageEditor, type PageSelection } from "./editor/usePageEditor";
import { HomeRibbon } from "./editor/HomeRibbon";
import { HomeDialogs } from "./editor/HomeDialogs";
import type { MenuAnchor } from "./editor/HomeMenu";

type Pane =
  | "find"
  | "replace"
  | "edit"
  | "fonts"
  | "compatibility"
  | "styles"
  | "outline"
  | null;
const MAX_BYTES = 50 * 1024 * 1024;
const titles: Record<Exclude<Pane, null>, string> = {
  find: "Navigation",
  replace: "Find and replace",
  edit: "Paragraph editor",
  fonts: "Document fonts",
  compatibility: "Document compatibility",
  styles: "Styles",
  outline: "Document outline",
};

export function DocxWorkspace() {
  const [document, setDocument] = useState<DocumentIR | null>(null),
    [session, setSession] = useState<DocxSession | null>(null);
  const [filename, setFilename] = useState("Sample document.docx"),
    [busy, setBusy] = useState(true),
    [error, setError] = useState("");
  const [tab, setTab] = useState<RibbonTab>("Home"),
    [pane, setPane] = useState<Pane>(null),
    [modal, setModal] = useState<EditorModal>(null);
  const [fileView, setFileView] = useState<string | null>(null),
    [selectedId, setSelectedId] = useState("");
  const [range, setRange] = useState<{ start: number; end: number } | null>(
      null,
    ),
    [draft, setDraft] = useState("");
  const [pageCount, setPageCount] = useState(0),
    [currentPage, setCurrentPage] = useState(1),
    [zoom, setZoom] = useState(1);
  const [query, setQuery] = useState(""),
    [replacement, setReplacement] = useState(""),
    [searchOptions, setSearchOptions] = useState<SearchOptions>({});
  const [activeHit, setActiveHit] = useState<SearchHit | null>(null);
  const matchCase = !!searchOptions.matchCase;
  const setMatchCase = (checked: boolean) =>
    setSearchOptions((previous) => ({ ...previous, matchCase: checked }));
  const findInput = useRef<HTMLInputElement>(null);
  const [notice, setNotice] = useState(
    "Click on the page and type. Use Home to format your text.",
  );
  const [revision, setRevision] = useState(0),
    [downloaded, setDownloaded] = useState(-1);
  const [showMarks, setShowMarks] = useState(false),
    [painter, setPainter] = useState<RunFormat | null>(null);
  const [repeatFormat, setRepeatFormat] = useState<RunFormat>({
    underline: "single",
    highlight: "yellow",
    color: "#c42b1c",
  });
  const [homeDialog, setHomeDialog] = useState<string | null>(null),
    [styleList, setStyleList] = useState<DocumentStyle[]>([]);
  const [homeAnchor, setHomeAnchor] = useState<MenuAnchor | null>(null),
    [contextStyle, setContextStyle] = useState("");
  const openHome = (name: string, anchor?: MenuAnchor) => {
    if (anchor) setHomeAnchor(anchor);
    setHomeDialog(name);
  };
  const lastList = useRef<
    Partial<
      Record<"bullet" | "decimal", Extract<EditCommand, { type: "list" }>>
    >
  >({});
  const [clipboard, setClipboard] = useState<string[]>([]),
    [scope, setScope] = useState<PageSelection[]>([]);
  const fileInput = useRef<HTMLInputElement>(null),
    canvas = useRef<HTMLDivElement>(null),
    textInput = useRef<HTMLTextAreaElement>(null);
  const request = useRef(0),
    changing = useRef(false),
    activeSession = useRef<DocxSession | null>(null);
  const all = useMemo(() => paragraphs(document), [document]),
    selected = all.find((p) => p.id === selectedId);
  const text = selected ? paragraphText(selected) : "",
    draftDirty = !!selected && draft !== text;
  const unsaved = draftDirty || (!!session?.dirty && downloaded !== revision);
  const fonts = useDocumentFonts(document),
    capability = session?.capability(selectedId),
    canFormat = !!capability?.format;
  const words = session?.statistics.words ?? 0;
  const results = useMemo(() => {
    return session?.find(query, searchOptions) ?? [];
  }, [session, document, query, searchOptions]);
  const searchVisible =
    pane === "find" ||
    pane === "replace" ||
    homeDialog === "findDialog" ||
    homeDialog === "replace";
  useSearchHighlights(canvas, results, activeHit, searchVisible);
  useEffect(() => {
    if (pane === "find") findInput.current?.focus();
  }, [pane]);
  const editor = usePageEditor({
    container: canvas,
    session,
    document,
    isBusy: () => changing.current,
    onCommit: (next) => {
      setDocument(next);
      setRevision((r) => r + 1);
    },
    onSelect: (picked, value) => {
      setSelectedId(picked.id);
      setRange({ start: picked.start, end: picked.end });
      setDraft(value);
      setScope([]);
    },
    onCommand: execute,
    onHistory: (direction) => void history(direction),
    onError: setError,
    onNotice: setNotice,
  });
  useEffect(() => {
    let current = true;
    session?.styles().then((items) => {
      if (current) setStyleList(items);
    });
    return () => {
      current = false;
    };
  }, [session, document, revision]);
  useEffect(() => {
    editor.refresh();
    editor.focus();
  }, [showMarks]);

  async function open(bytes: Uint8Array, name: string) {
    const token = ++request.current;
    setBusy(true);
    setError("");
    try {
      if (bytes.byteLength > MAX_BYTES)
        throw new Error("Choose a DOCX smaller than 50 MB.");
      const next = await openDocxSession(bytes, {
        maxUncompressedBytes: MAX_BYTES,
      });
      if (token !== request.current) return;
      activeSession.current = next;
      setSession(next);
      setDocument(next.document);
      setFilename(name);
      const first = paragraphs(next.document)[0];
      setSelectedId(first?.id ?? "");
      setDraft(first ? paragraphText(first) : "");
      setRange(null);
      setRevision(0);
      setDownloaded(-1);
      setPageCount(0);
      setCurrentPage(1);
      setFileView(null);
      setPane(null);
      setScope([]);
      setPainter(null);
      setHomeDialog(null);
      setActiveHit(null);
      setNotice(
        next.lockedReason ??
          "Click on the page and type. Use Home to format your text.",
      );
    } catch (cause) {
      if (token === request.current)
        setError(
          cause instanceof Error
            ? cause.message
            : "Could not open this document.",
        );
    } finally {
      if (token === request.current) setBusy(false);
    }
  }
  function discardOK() {
    return (
      !unsaved ||
      window.confirm(
        "Some changes have not been downloaded. Discard them and open another document?",
      )
    );
  }
  async function sample(initial = false) {
    if (!initial && !discardOK()) return;
    try {
      const response = await fetch("/sample.docx");
      if (!response.ok)
        throw new Error(
          "Sample unavailable. Open a DOCX or start a blank document.",
        );
      await open(
        new Uint8Array(await response.arrayBuffer()),
        "Sample document.docx",
      );
    } catch (cause) {
      setBusy(false);
      setError(
        cause instanceof Error ? cause.message : "Could not load sample.",
      );
    }
  }
  useEffect(() => {
    void sample(true);
    return () => {
      request.current++;
    };
  }, []);
  useEffect(() => {
    const warn = (event: BeforeUnloadEvent) => {
      if (unsaved) event.preventDefault();
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [unsaved]);
  useEffect(() => {
    canvas.current
      ?.querySelectorAll<HTMLElement>("[data-docx-page] [data-paragraph-id]")
      .forEach((node) =>
        node.toggleAttribute(
          "data-selected",
          scope.length
            ? scope.some((target) => target.id === node.dataset.paragraphId)
            : node.dataset.paragraphId === selectedId,
        ),
      );
  }, [selectedId, document, pageCount, zoom, scope]);
  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (event.defaultPrevented) return;
      if (event.key === "Escape") {
        setFileView(null);
        window.document
          .querySelectorAll<HTMLDetailsElement>(".tool-menu[open]")
          .forEach((e) => (e.open = false));
      }
      if (event.ctrlKey || event.metaKey) {
        if (event.key.toLowerCase() === "s") {
          event.preventDefault();
          void save();
        }
        if (event.key.toLowerCase() === "f") {
          event.preventDefault();
          setPane("find");
          findInput.current?.focus();
        }
        if (event.key.toLowerCase() === "h") {
          event.preventDefault();
          setHomeDialog("replace");
        }
        if (event.key.toLowerCase() === "o") {
          event.preventDefault();
          fileInput.current?.click();
        }
        const field = (event.target as Element)?.closest(
          "input,textarea,select",
        );
        if (
          !field &&
          event.key.toLowerCase() === "a" &&
          canvas.current?.contains(window.document.activeElement)
        ) {
          event.preventDefault();
          void action("selectAll");
        }
        if (!field && event.key.toLowerCase() === "d" && selected) {
          event.preventDefault();
          setHomeDialog("font");
        }
        if (!field && ["b", "i", "u"].includes(event.key.toLowerCase())) {
          event.preventDefault();
          const key = event.key.toLowerCase();
          run(
            key === "b"
              ? { bold: !activeRun().bold }
              : key === "i"
                ? { italic: !activeRun().italic }
                : {
                    underline:
                      activeRun().underline && activeRun().underline !== "none"
                        ? "none"
                        : "single",
                  },
          );
        }
        if (!field && event.key.toLowerCase() === "z") {
          event.preventDefault();
          void history(event.shiftKey ? "redo" : "undo");
        }
        if (!field && event.key.toLowerCase() === "y") {
          event.preventDefault();
          void history("redo");
        }
      }
      if (
        event.key === "F3" &&
        !event.ctrlKey &&
        !event.altKey &&
        !event.metaKey
      ) {
        event.preventDefault();
        void findText(query, searchOptions, event.shiftKey ? "up" : "down");
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  });
  useEffect(() => {
    const container = canvas.current;
    if (!container) return;
    const observer = new IntersectionObserver(
      (entries) => {
        const visible = entries
          .filter((e) => e.isIntersecting)
          .sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0];
        if (visible)
          setCurrentPage(
            Number((visible.target as HTMLElement).dataset.docxPage),
          );
      },
      { root: container, threshold: [0.05, 0.25, 0.5] },
    );
    container
      .querySelectorAll("[data-docx-page]")
      .forEach((e) => observer.observe(e));
    return () => observer.disconnect();
  }, [document, pageCount]);

  async function flushDraft(active: DocxSession) {
    if (editor.hasBuffer()) {
      await editor.flush();
      return;
    }
    const current = paragraphs(active.document).find(
      (p) => p.id === selectedId,
    );
    if (draftDirty && current && paragraphText(current) !== draft) {
      await active.execute({
        type: "text",
        id: selectedId,
        ...textChange(paragraphText(current), draft),
      });
      setRevision((r) => r + 1);
    }
  }
  function sync(active: DocxSession, target?: PageSelection) {
    if (activeSession.current !== active) return;
    const picked = target ?? editor.getSelection();
    setDocument(active.document);
    const p =
      paragraphs(active.document).find(
        (p) => p.id === (picked?.id ?? selectedId),
      ) ?? paragraphs(active.document)[0];
    setSelectedId(p?.id ?? "");
    setDraft(p ? paragraphText(p) : "");
    if (picked && p?.id === picked.id) {
      setRange({ start: picked.start, end: picked.end });
      editor.select(picked);
    }
  }
  async function execute(command: EditCommand) {
    const active = activeSession.current;
    if (!active || changing.current) return false;
    changing.current = true;
    setBusy(true);
    setError("");
    try {
      if (command.type !== "text") await flushDraft(active);
      await active.execute(command);
      sync(
        active,
        command.type === "splitParagraph" || command.type === "joinParagraph"
          ? active.selection
          : undefined,
      );
      setRevision((r) => r + 1);
      setNotice("Changes on this device. Download DOCX to save a copy.");
      return true;
    } catch (cause) {
      sync(active);
      setError(
        cause instanceof Error
          ? cause.message
          : "The change could not be applied.",
      );
      return false;
    } finally {
      changing.current = false;
      setBusy(false);
    }
  }
  async function executeMany(commands: EditCommand[]) {
    if (commands.length === 1) return execute(commands[0]);
    const active = activeSession.current;
    if (!active || changing.current || !commands.length) return false;
    changing.current = true;
    setBusy(true);
    setError("");
    try {
      await flushDraft(active);
      await active.executeBatch(commands);
      sync(active);
      setRevision((r) => r + 1);
      setNotice(`${commands.length} changes applied as one undo action.`);
      return true;
    } catch (cause) {
      sync(active);
      setError(
        cause instanceof Error
          ? cause.message
          : "Changes could not be applied.",
      );
      return false;
    } finally {
      changing.current = false;
      setBusy(false);
    }
  }
  async function history(direction: "undo" | "redo") {
    const active = activeSession.current;
    if (!active || changing.current) return;
    changing.current = true;
    setBusy(true);
    try {
      await flushDraft(active);
      await active[direction]();
      sync(active);
      setRevision((r) => r + 1);
      setNotice(direction === "undo" ? "Change undone." : "Change restored.");
      editor.focus();
    } catch (cause) {
      setError(
        cause instanceof Error ? cause.message : "Could not restore history.",
      );
    } finally {
      changing.current = false;
      setBusy(false);
    }
  }
  async function save() {
    const active = activeSession.current;
    if (!active || changing.current) return;
    changing.current = true;
    setBusy(true);
    setError("");
    try {
      await flushDraft(active);
      sync(active);
      download(await active.exportDocx(), filename);
      setDownloaded(revision + (draftDirty ? 1 : 0));
      setNotice(
        "DOCX download started. The original file has not been overwritten.",
      );
      setFileView(null);
    } catch (cause) {
      setError(
        cause instanceof Error ? cause.message : "Could not download DOCX.",
      );
    } finally {
      changing.current = false;
      setBusy(false);
    }
  }
  function select(id: string, edit = false) {
    if (busy) return;
    if (draftDirty && id !== selectedId) {
      setNotice(
        "Apply or cancel paragraph text changes before selecting another paragraph.",
      );
      return;
    }
    const p = all.find((p) => p.id === id);
    if (!p) return;
    if (id !== selectedId) {
      setSelectedId(id);
      setDraft(paragraphText(p));
      setRange(null);
    }
    if (edit) {
      setPane("edit");
      setTimeout(() => textInput.current?.focus(), 0);
    }
  }
  function jump(id: string) {
    select(id);
    editor.select({ id, start: 0, end: 0 });
    canvas.current
      ?.querySelector<HTMLElement>(
        `[data-docx-page] [data-paragraph-id="${CSS.escape(id)}"]`,
      )
      ?.scrollIntoView({ behavior: "smooth", block: "center" });
  }
  function chooseSearchHit(hit: SearchHit) {
    select(hit.id);
    editor.select(hit);
    setRange({ start: hit.start, end: hit.end });
    setActiveHit(hit);
    setScope([]);
    const node = Array.from(
      canvas.current?.querySelectorAll<HTMLElement>(
        `[data-docx-page] [data-paragraph-id="${CSS.escape(hit.id)}"]`,
      ) ?? [],
    ).find((el) => el.getClientRects().length);
    if (node) {
      restoreSelection(node, hit.start, hit.end);
      node.scrollIntoView({ behavior: "smooth", block: "center" });
    }
  }
  async function findText(
    value: string,
    options: SearchOptions,
    direction: SearchDirection = "down",
  ) {
    await editor.flush();
    setQuery(value);
    setSearchOptions(options);
    const items = paragraphs(activeSession.current?.document ?? document);
    const matches = activeSession.current?.find(value, options) ?? [];
    const index = nextMatchIndex(
      matches,
      items.map((p) => p.id!),
      editor.getSelection(),
      direction,
    );
    if (index < 0) {
      setActiveHit(null);
      return "No results found.";
    }
    chooseSearchHit(matches[index]);
    const message = `Result ${index + 1} of ${matches.length}.`;
    setNotice(message);
    return message;
  }
  function activeRun(): RunFormat {
    const picked = editor.getSelection(),
      p = all.find((p) => p.id === (picked?.id ?? selectedId));
    let offset = 0;
    const runs = p?.children.filter((c) => c.type === "text") ?? [];
    const current =
      runs.find((r) => {
        offset += r.text.length;
        return picked && picked.end > picked.start
          ? offset > picked.start
          : offset >= (picked?.start ?? 0);
      })?.properties ?? p?.markProperties;
    const char = p
      ? paragraphText(p).slice(
          Math.max(0, (picked?.start ?? 0) - 1),
          (picked?.start ?? 0) + 1,
        )
      : "";
    const cs =
      current?.forceComplexScript ||
      current?.direction === "rtl" ||
      /[\p{Script=Arabic}\p{Script=Hebrew}]/u.test(char);
    const base = editor.typingFormat.clear ? p?.markProperties : current;
    return {
      ...base,
      fontFamily: cs
        ? (base?.complexScript?.fontFamily ?? base?.fontFamily)
        : base?.fontFamily,
      fontSizePt: cs
        ? (base?.complexScript?.fontSizePt ?? base?.fontSizePt)
        : base?.fontSizePt,
      bold: cs ? base?.complexScript?.bold : base?.bold,
      italic: cs ? base?.complexScript?.italic : base?.italic,
      complexScriptFont: base?.complexScript?.fontFamily,
      complexScriptSizePt: base?.complexScript?.fontSizePt,
      underline: base?.underline as RunFormat["underline"],
      ...editor.typingFormat,
    };
  }
  function targets() {
    return scope.length
      ? scope
      : editor.getSelection()
        ? [editor.getSelection()!]
        : selected
          ? [
              {
                id: selectedId,
                start: range?.start ?? 0,
                end: range?.end ?? text.length,
              },
            ]
          : [];
  }
  function run(format: RunFormat) {
    setRepeatFormat((previous) => ({
      ...previous,
      ...(format.underline && format.underline !== "none"
        ? { underline: format.underline }
        : {}),
      ...(format.highlight !== undefined
        ? { highlight: format.highlight }
        : {}),
      ...(format.color !== undefined ? { color: format.color } : {}),
    }));
    const picked = editor.getSelection();
    if (!scope.length && picked && picked.start === picked.end) {
      editor.setTyping({ ...editor.typingFormat, ...format });
      editor.focus();
      return;
    }
    void executeMany(
      targets().map((target) => ({
        type: "run",
        id: target.id,
        format,
        start: target.start,
        end: target.end,
      })),
    );
  }
  function para(format: ParagraphFormat) {
    if (format.styleId) editor.setTyping({});
    void executeMany(
      targets().map((target) => ({ type: "paragraph", id: target.id, format })),
    );
  }
  function insertText(value: string) {
    if (!selected || !canFormat) return;
    if (!editor.getSelection())
      editor.select({
        id: selectedId,
        start: range?.start ?? text.length,
        end: range?.end ?? text.length,
      });
    editor.insert(value);
  }
  function page(
    widthPt: number,
    heightPt: number,
    margins?: { top: number; bottom: number; left: number; right: number },
  ) {
    void execute({ type: "page", widthPt, heightPt, margins });
  }
  function homeCommand(command: EditCommand) {
    if (command.type !== "list") {
      void execute(command);
      return;
    }
    if (
      command.kind !== "none" &&
      !command.preserveDefinition &&
      !command.restart &&
      command.start === undefined &&
      !command.continue
    )
      lastList.current[command.kind === "bullet" ? "bullet" : "decimal"] = {
        ...command,
      };
    const chosen = targets(),
      picked = editor.getSelection();
    let ids = chosen.map((t) => t.id);
    const instance = selected?.properties.numbering?.id;
    if (
      !scope.length &&
      picked?.start === picked?.end &&
      instance &&
      instance !== "0" &&
      command.kind !== "none" &&
      !command.preserveDefinition
    ) {
      // A library choice restyles the current list; restart/value affects the tail only.
      let list = all.filter((p) => p.properties.numbering?.id === instance);
      if (command.restart || command.start !== undefined || command.continue)
        list = list.slice(
          Math.max(
            0,
            list.findIndex((p) => p.id === selectedId),
          ),
        );
      ids = list.map((p) => p.id!);
    }
    void executeMany(
      ids.map((id, index) => ({
        ...command,
        id,
        level:
          command.level ??
          all.find((p) => p.id === id)?.properties.numbering?.level ??
          0,
        ...(index > 0
          ? {
              restart: false,
              start: undefined,
              continue:
                command.restart ||
                command.start !== undefined ||
                command.continue,
            }
          : {}),
      })),
    );
  }
  async function styleAction(name: string, id: string) {
    if (name === "update") {
      const picked = editor.getSelection();
      if (!picked) {
        setNotice(
          "Select the text whose formatting you want the style to use.",
        );
        return;
      }
      const before = activeSession.current?.document;
      const done = await execute({
        type: "updateStyleFromSelection",
        styleId: id,
        ...picked,
      });
      if (done) {
        editor.setTyping({});
        setNotice(
          activeSession.current?.document === before
            ? "No common formatting changes were found in the selection."
            : "Style updated from the selected text. Other text using the style now follows it.",
        );
      }
    } else if (name === "select") {
      const ranges = all
        .filter((p) => (p.properties.styleId ?? "Normal") === id)
        .map((p) => ({ id: p.id!, start: 0, end: paragraphText(p).length }));
      setScope(ranges);
      setNotice(`${ranges.length} paragraphs selected with this style.`);
    } else if (name === "remove")
      await execute({ type: "style", styleId: id, gallery: false });
  }
  async function action(name: string) {
    if (name === "marks") {
      setShowMarks((value) => !value);
      return;
    }
    if (name === "painter") {
      setPainter((value) => (value ? null : activeRun()));
      setNotice("Format Painter: select the destination text on the page.");
      return;
    }
    if (name === "clear") {
      run({ clear: true });
      return;
    }
    if (name === "find") {
      await editor.flush();
      setPane("find");
      findInput.current?.focus();
      return;
    }
    if (name.startsWith("list:")) {
      const kind = name.slice(5) as "bullet" | "decimal";
      const same = targets().every((t) => {
        const p = all.find((p) => p.id === t.id);
        return kind === "bullet"
          ? p?.properties.list?.format === "bullet"
          : !!p?.properties.list && p.properties.list.format !== "bullet";
      });
      const remembered = lastList.current[kind];
      homeCommand({
        type: "list",
        id: selectedId,
        kind: same ? "none" : (remembered?.kind ?? kind),
        ...(!same && remembered
          ? {
              bullet: remembered.bullet,
              pattern: remembered.pattern,
              multilevel: remembered.multilevel,
            }
          : {}),
      });
      return;
    }
    if (name.startsWith("level:")) {
      const level = Math.max(
        0,
        Math.min(
          8,
          (selected?.properties.numbering?.level ?? 0) +
            (name === "level:up" ? 1 : -1),
        ),
      );
      const kind =
        selected?.properties.list?.format === "bullet"
          ? "bullet"
          : selected?.properties.list?.format === "lowerLetter"
            ? "lowerLetter"
            : selected?.properties.list?.format === "upperRoman"
              ? "upperRoman"
              : "decimal";
      void executeMany(
        targets().map((target) => ({
          type: "list",
          id: target.id,
          kind,
          level,
          preserveDefinition: true,
        })),
      );
      return;
    }
    if (name === "selectAll" || name === "selectSimilar") {
      await editor.flush();
      const candidates =
        name === "selectAll"
          ? all
          : all.filter(
              (p) => p.properties.styleId === selected?.properties.styleId,
            );
      const ranges = candidates
        .filter((p) => session?.capability(p.id!).format)
        .map((p) => ({ id: p.id!, start: 0, end: paragraphText(p).length }));
      setScope(ranges);
      setNotice(`${ranges.length} editable paragraphs selected.`);
      return;
    }
    if (name.startsWith("goto:")) {
      canvas.current
        ?.querySelector(
          `[data-docx-page="${Math.max(1, Math.min(pageCount, Number(name.slice(5))))}"]`,
        )
        ?.scrollIntoView({ block: "start", behavior: "smooth" });
      return;
    }
    if (name.startsWith("case:")) {
      const commands: EditCommand[] = [];
      for (const target of targets()) {
        const p = all.find((p) => p.id === target.id);
        if (!p) continue;
        const value = paragraphText(p);
        let start = target.start,
          end = target.end;
        if (start === end) {
          const word = Array.from(
            new Intl.Segmenter(document?.language, {
              granularity: "word",
            }).segment(value),
          ).find(
            (part) =>
              part.isWordLike &&
              part.index <= start &&
              part.index + part.segment.length >= start,
          );
          if (!word) continue;
          start = word.index;
          end = start + word.segment.length;
        }
        commands.push({
          type: "text",
          id: target.id,
          start,
          end,
          text: changeCase(
            value.slice(start, end),
            name.slice(5),
            document?.language,
          ),
        });
      }
      await executeMany(commands);
      return;
    }
    if (name.startsWith("symbol:")) {
      insertText(name.slice(7));
      return;
    }
    const w = document?.page?.widthPt ?? 595.3,
      h = document?.page?.heightPt ?? 841.9;
    if (name.startsWith("margin:")) {
      const n = Number(name.slice(7));
      page(w, h, {
        top: n === 54 ? 72 : n,
        bottom: n === 54 ? 72 : n,
        left: n,
        right: n,
      });
      return;
    }
    if (name.startsWith("size:")) {
      const [width, height] = name.slice(5).split(",").map(Number);
      page(width, height);
      return;
    }
    if (name.startsWith("orientation:")) {
      const landscape = name.endsWith("Landscape");
      page(
        landscape ? Math.max(w, h) : Math.min(w, h),
        landscape ? Math.min(w, h) : Math.max(w, h),
      );
      return;
    }
    if (
      ["fonts", "find", "replace", "edit", "styles", "outline"].includes(name)
    ) {
      setPane(name as Pane);
      return;
    }
    if (name === "stories") {
      setPane("outline");
      return;
    }
    if (["page", "paragraph", "table"].includes(name)) {
      setModal(name as EditorModal);
      return;
    }
    if (name === "insertParagraph" || name === "blankPage") {
      await execute({
        type: "insertParagraph",
        id: selectedId || undefined,
        text: "",
        pageBreakBefore: name === "blankPage",
      });
      return;
    }
    if (name === "date") {
      insertText(
        new Intl.DateTimeFormat(document?.language ?? "en", {
          dateStyle: "long",
          calendar: "gregory",
        }).format(new Date()),
      );
      return;
    }
    if (name === "copy" || name === "cut") {
      try {
        const picked = targets(),
          value = picked
            .map((target) => {
              const p = all.find((p) => p.id === target.id);
              const value =
                target.id === selectedId ? draft : p ? paragraphText(p) : "";
              return value.slice(target.start, target.end);
            })
            .join("\n");
        await navigator.clipboard.writeText(value);
        setClipboard((items) =>
          [value, ...items.filter((item) => item !== value)].slice(0, 12),
        );
        if (name === "cut")
          await executeMany(
            picked.map((target) => ({
              type: "text",
              id: target.id,
              start: target.start,
              end: target.end,
              text: "",
            })),
          );
        else setNotice("Text copied.");
      } catch {
        setPane("edit");
        setNotice("Use Ctrl+C in the paragraph editor to copy text.");
      }
      return;
    }
    if (name === "paste") {
      try {
        insertText(await navigator.clipboard.readText());
      } catch {
        editor.focus();
        setNotice("Click the page and press Ctrl+V to paste.");
      }
    }
  }
  async function replaceText(
    find: string,
    value: string,
    options: SearchOptions,
    allMatches: boolean,
  ) {
    if (!find) return "Enter text to find.";
    setQuery(find);
    setReplacement(value);
    setSearchOptions(options);
    await editor.flush();
    const commands: EditCommand[] = [],
      active = activeSession.current;
    const items = paragraphs(active?.document ?? document);
    const selectedMatch = editor.getSelection();
    let skipped = 0;
    for (const p of items) {
      const matches = searchMatches(paragraphText(p), find, options);
      if (!matches.length) continue;
      if (!active?.capability(p.id!).text) {
        skipped += matches.length;
        continue;
      }
      const hits = allMatches
        ? matches
        : matches.filter(
            (hit) =>
              selectedMatch &&
              selectedMatch.id === p.id &&
              hit.start === selectedMatch.start &&
              hit.end === selectedMatch.end,
          );
      for (const hit of [...hits].reverse())
        commands.push({ type: "text", id: p.id!, ...hit, text: value });
      if (!allMatches && hits.length) break;
    }
    if (!allMatches && !commands.length) return findText(find, options);
    const done = await executeMany(commands);
    const message = done
      ? `${commands.length} replacement(s). ${skipped ? `${skipped} protected match(es) skipped.` : ""}`
      : commands.length
        ? "Replacement could not be completed."
        : "No editable matches found.";
    setActiveHit(null);
    if (done && !allMatches) {
      const command = commands[0] as Extract<EditCommand, { type: "text" }>;
      editor.select({
        id: command.id,
        start: command.start + value.length,
        end: command.start + value.length,
      });
      await findText(find, options);
    }
    setNotice(message);
    return message;
  }
  function capturePageSelection() {
    if (!canvas.current) return;
    const picked = selectionRanges(canvas.current, document);
    setScope(picked.length > 1 ? picked : []);
    editor.capture(true);
    if (painter && picked.length) {
      void executeMany(
        picked.map((target) => ({
          type: "run",
          id: target.id,
          format: painter,
          start: target.start,
          end:
            target.end === target.start
              ? all.find((p) => p.id === target.id)
                ? paragraphText(all.find((p) => p.id === target.id)!).length
                : target.end
              : target.end,
        })),
      );
      setPainter(null);
    }
  }
  function fit(mode: "width" | "page") {
    if (!canvas.current) return;
    const w = ((document?.page?.widthPt ?? 595.3) * 4) / 3,
      h = ((document?.page?.heightPt ?? 841.9) * 4) / 3;
    setZoom(
      Math.max(
        0.25,
        Math.min(
          2,
          (canvas.current.clientWidth - 80) / w,
          mode === "page" ? (canvas.current.clientHeight - 55) / h : 2,
        ),
      ),
    );
  }
  async function print() {
    if (draftDirty && session)
      await execute({
        type: "text",
        id: selectedId,
        ...textChange(text, draft),
      });
    await window.document.fonts.ready;
    setFileView(null);
    requestAnimationFrame(() => requestAnimationFrame(() => window.print()));
  }

  return (
    <div className="word-app">
      <header className="titlebar">
        <span className="app-mark">D</span>
        <button
          className="filename"
          onClick={() => setModal("rename")}
          title="Rename download"
        >
          <bdi>{filename}</bdi>
          <Icon name="chevron" size={13} />
        </button>
        <span className="save-state">
          <span className="state-dot" />
          {busy
            ? "Working…"
            : draftDirty
              ? "Text changes pending"
              : unsaved
                ? "Changes on this device"
                : session?.dirty
                  ? "DOCX downloaded"
                  : "Stored on this device"}
        </span>
        <div className="quick-access">
          <Tool
            label="Download DOCX (Ctrl+S)"
            icon="save"
            onClick={() => void save()}
            disabled={busy || !session}
          />
          <Tool
            label="Undo (Ctrl+Z)"
            icon="undo"
            onClick={() => void history("undo")}
            disabled={busy || (!session?.canUndo && !draftDirty)}
          />
          <Tool
            label="Redo"
            icon="redo"
            onClick={() => void history("redo")}
            disabled={busy || !session?.canRedo}
          />
        </div>
        <button className="command-search" onClick={() => setPane("find")}>
          <Icon name="search" size={16} />
          <span>Find in document</span>
          <kbd>Ctrl F</kbd>
        </button>
        <span className="local-badge">Local workspace</span>
      </header>
      <nav className="tabbar" aria-label="Document ribbon">
        <button
          className="file-tab"
          aria-expanded={!!fileView}
          onClick={() => setFileView("Home")}
        >
          File
        </button>
        <div role="tablist" aria-label="Ribbon tabs">
          {(["Home", "Insert", "Layout"] as RibbonTab[]).map((name, index) => (
            <button
              id={`tab-${name}`}
              key={name}
              role="tab"
              aria-selected={tab === name}
              aria-controls={`ribbon-${name}`}
              tabIndex={tab === name ? 0 : -1}
              className={tab === name ? "selected-tab" : ""}
              onClick={() => setTab(name)}
              onKeyDown={(e) => {
                if (e.key === "ArrowRight" || e.key === "ArrowLeft") {
                  e.preventDefault();
                  const names: RibbonTab[] = ["Home", "Insert", "Layout"],
                    next =
                      names[(index + (e.key === "ArrowRight" ? 1 : 2)) % 3];
                  setTab(next);
                  window.document.getElementById(`tab-${next}`)?.focus();
                }
              }}
            >
              {name}
            </button>
          ))}
        </div>
        <div className="tabbar-right">
          <button onClick={() => setPane("fonts")}>
            Fonts
            {fonts.missing.length ? (
              <span className="count-dot">{fonts.missing.length}</span>
            ) : null}
          </button>
          <button onClick={() => setPane("compatibility")}>
            Compatibility
          </button>
        </div>
      </nav>
      {tab === "Home" ? (
        <HomeRibbon
          selected={selected}
          format={activeRun()}
          repeat={repeatFormat}
          disabled={busy || !!session?.lockedReason}
          canFormat={canFormat}
          hasSelection={
            scope.length > 0 || (!!range && range.end > range.start)
          }
          styles={styleList}
          onStyleContext={(id, anchor) => {
            setContextStyle(id);
            openHome("styleContext", anchor);
          }}
          showMarks={showMarks}
          painter={!!painter}
          onRun={run}
          onParagraph={para}
          onAction={(name) => void action(name)}
          onDialog={openHome}
        />
      ) : (
        <Ribbon
          tab={tab}
          document={document}
          selected={selected}
          range={range}
          disabled={busy || !!session?.lockedReason}
          canFormat={canFormat}
          canPage={!!session?.canSetPage}
          families={fonts.required.map((f) => f.family)}
          onRun={run}
          onParagraph={para}
          onAction={(name) => void action(name)}
        />
      )}
      <div className="context-strip">
        <span>
          <span className="mode-dot" />
          Document editing
        </span>
        <span>
          {selected
            ? `${range && range.end > range.start ? `${range.end - range.start} characters selected` : "Paragraph selected"} · ${selected.properties.direction?.toUpperCase() ?? "LTR"}`
            : "Select a paragraph"}
        </span>
        <button onClick={() => setPane("edit")}>
          Open paragraph editor ↗
        </button>
        {fonts.missing.length > 0 && (
          <button className="missing-pill" onClick={() => setPane("fonts")}>
            {fonts.missing.length} missing fonts
          </button>
        )}
      </div>
      {error && (
        <div className="error-banner" role="alert">
          <Icon name="info" size={17} />
          <span>{error}</span>
          <button aria-label="Dismiss error" onClick={() => setError("")}>
            ×
          </button>
        </div>
      )}
      <div className="workspace-body">
        <main
          ref={canvas}
          className={`page-workspace ${painter ? "format-painter-active" : ""}`}
          aria-label="Document pages"
          aria-busy={busy}
          onMouseUp={capturePageSelection}
        >
          <div className="page-ruler" aria-hidden="true">
            {[1, 2, 3, 4, 5, 6, 7, 8].map((n) => (
              <span key={n}>{n}</span>
            ))}
          </div>
          {document ? (
            <PaginatedDocument
              document={document}
              fontRevision={fonts.revision}
              onPageCountChange={setPageCount}
              zoom={zoom}
              showFormattingMarks={showMarks}
            />
          ) : (
            <div className="welcome-sheet">
              <Icon name="document" size={48} />
              <h1>Your next document starts here</h1>
              <p>Open a Word file or start with a blank page.</p>
              <button
                className="primary-button"
                disabled={busy}
                onClick={() =>
                  void open(createBlankDocx(), "Untitled document.docx")
                }
              >
                Create blank document
              </button>
              <button
                className="secondary-button"
                onClick={() => fileInput.current?.click()}
              >
                Open DOCX
              </button>
            </div>
          )}
        </main>
        {pane && (
          <aside className="task-pane" aria-label={titles[pane]}>
            <div className="pane-title">
              <h2>{titles[pane]}</h2>
              <button aria-label="Close pane" onClick={() => setPane(null)}>
                <Icon name="close" size={18} />
              </button>
            </div>
            <div className="pane-content">
              {pane === "fonts" && <FontPanel fonts={fonts} />}
              {pane === "edit" &&
                (selected ? (
                  <>
                    <p className="pane-description">
                      Edit this paragraph here. Select text here or on the page
                      to format a range. With no range, formatting applies to
                      all its text.
                    </p>
                    {!canFormat && (
                      <div className="pane-callout">{capability?.reason}</div>
                    )}
                    <label className="field">
                      Paragraph text
                      <textarea
                        ref={textInput}
                        dir={selected.properties.direction ?? "ltr"}
                        className="paragraph-editor"
                        value={draft}
                        disabled={!canFormat || busy}
                        onChange={(e) => {
                          setDraft(e.target.value);
                          setRange(null);
                        }}
                        onSelect={(e) =>
                          setRange({
                            start: e.currentTarget.selectionStart,
                            end: e.currentTarget.selectionEnd,
                          })
                        }
                      />
                    </label>
                    <div className="pane-actions">
                      <button
                        className="primary-button"
                        disabled={!draftDirty || !canFormat || busy}
                        onClick={() =>
                          void execute({
                            type: "text",
                            id: selectedId,
                            ...textChange(text, draft),
                          })
                        }
                      >
                        Apply text changes
                      </button>
                      <button
                        className="secondary-button"
                        disabled={!draftDirty || busy}
                        onClick={() => {
                          setDraft(text);
                          setRange(null);
                        }}
                      >
                        Cancel
                      </button>
                    </div>
                    <div className="selection-summary">
                      {range && range.end > range.start
                        ? `${range.end - range.start} selected characters`
                        : "Formatting scope: whole paragraph"}
                    </div>
                  </>
                ) : (
                  <p>Select a paragraph on the page.</p>
                ))}
              {(pane === "find" || pane === "replace") && (
                <>
                  <label className="field">
                    Find
                    <input
                      ref={findInput}
                      autoFocus
                      value={query}
                      onChange={(e) => {
                        setQuery(e.target.value);
                        setActiveHit(null);
                      }}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") {
                          e.preventDefault();
                          void findText(
                            query,
                            searchOptions,
                            e.shiftKey ? "up" : "down",
                          );
                        }
                      }}
                      placeholder="Search document text"
                    />
                  </label>
                  <label className="check-field">
                    <input
                      type="checkbox"
                      checked={!!searchOptions.wholeWords}
                      onChange={(e) =>
                        setSearchOptions({
                          ...searchOptions,
                          wholeWords: e.target.checked,
                        })
                      }
                    />
                    Find whole words only
                  </label>
                  <div className="pane-actions">
                    <button
                      className="secondary-button"
                      disabled={!results.length}
                      onClick={() => void findText(query, searchOptions, "up")}
                    >
                      Previous
                    </button>
                    <button
                      className="secondary-button"
                      disabled={!results.length}
                      onClick={() => void findText(query, searchOptions)}
                    >
                      Next
                    </button>
                    <button
                      className="secondary-button"
                      onClick={() => openHome("findDialog")}
                    >
                      Advanced Find…
                    </button>
                  </div>
                  <label className="check-field">
                    <input
                      type="checkbox"
                      checked={matchCase}
                      onChange={(e) => setMatchCase(e.target.checked)}
                    />
                    Match case
                  </label>
                  {pane === "replace" && (
                    <>
                      <label className="field">
                        Replace with
                        <input
                          value={replacement}
                          onChange={(e) => setReplacement(e.target.value)}
                        />
                      </label>
                      <p className="pane-description">
                        Choose a result, then replace the selected match.
                        Structured text stays locked.
                      </p>
                      <button
                        className="primary-button"
                        disabled={
                          busy ||
                          !query ||
                          !canFormat ||
                          draftDirty ||
                          !results.length
                        }
                        onClick={() =>
                          void replaceText(
                            query,
                            replacement,
                            searchOptions,
                            false,
                          )
                        }
                      >
                        Replace
                      </button>
                    </>
                  )}
                  <p className="result-count">
                    {query
                      ? `${results.length} results${results.length > 100 ? " (first 100 shown; Next reaches all results)" : ""}`
                      : "Search body text, tables, headers and footers."}
                  </p>
                  <div className="search-results">
                    {results.slice(0, 100).map((result, index) => (
                      <button
                        key={`${result.id}:${result.start}`}
                        className={
                          activeHit?.id === result.id &&
                          activeHit.start === result.start
                            ? "active"
                            : ""
                        }
                        onClick={() => chooseSearchHit(result)}
                      >
                        <small>Result {index + 1}</small>
                        <span dir="auto">
                          {result.content.slice(
                            Math.max(0, result.start - 35),
                            result.start,
                          )}
                          <mark>
                            {result.content.slice(result.start, result.end)}
                          </mark>
                          {result.content.slice(result.end, result.end + 90)}
                        </span>
                      </button>
                    ))}
                  </div>
                </>
              )}
              {pane === "compatibility" && (
                <>
                  <div className="pane-callout">
                    Source-retaining export keeps untouched package parts.
                    Supported edits include simple paragraph text and
                    formatting, paragraph properties, body paragraph/table
                    insertion and single-section page setup.
                  </div>
                  <p className="pane-description">
                    Structured text stays locked where it cannot be rewritten
                    safely. Browser pagination can differ from Word. Export has
                    automated round-trip checks; Word reopen fidelity is not yet
                    certified.
                  </p>
                  {session?.lockedReason && <p>{session.lockedReason}</p>}
                  <h3>Import notes</h3>
                  {document?.diagnostics?.length ? (
                    <ul className="diagnostic-list">
                      {document.diagnostics.map((note, index) => (
                        <li key={index}>{note.message}</li>
                      ))}
                    </ul>
                  ) : (
                    <p>
                      No import diagnostics. This does not certify full DOCX
                      compatibility.
                    </p>
                  )}
                  <details className="pane-help">
                    <summary>Developer: document IR</summary>
                    <pre>{JSON.stringify(document, null, 2)}</pre>
                  </details>
                </>
              )}
              {pane === "styles" && (
                <>
                  <p className="pane-description">
                    Source style names are retained. Use direct formatting for
                    edits; shared style-definition editing is not available yet.
                  </p>
                  {[
                    ...new Set(
                      all.map(
                        (p) => p.properties.styleId ?? "Default paragraph",
                      ),
                    ),
                  ].map((style) => (
                    <button
                      className="pane-style"
                      key={style}
                      onClick={() => {
                        const p = all.find(
                          (p) =>
                            (p.properties.styleId ?? "Default paragraph") ===
                            style,
                        );
                        if (p?.id) jump(p.id);
                      }}
                    >
                      <span>AaBbCc</span>
                      <b>{style}</b>
                      <small>Find a paragraph using this style</small>
                    </button>
                  ))}
                </>
              )}
              {pane === "outline" && (
                <>
                  <p className="pane-description">
                    Body, table, header and footer paragraphs. Choose one to
                    navigate or inspect.
                  </p>
                  <div className="outline-list">
                    {all.map((p, index) => (
                      <button
                        key={p.id}
                        className={selectedId === p.id ? "active" : ""}
                        onClick={() => jump(p.id!)}
                      >
                        <small>
                          {/header|footer/i.test(p.id ?? "")
                            ? "Header / footer"
                            : (p.properties.styleId ??
                              `Paragraph ${index + 1}`)}
                        </small>
                        <span dir="auto">
                          {paragraphText(p).slice(0, 90) || "(Empty paragraph)"}
                        </span>
                      </button>
                    ))}
                  </div>
                </>
              )}
            </div>
          </aside>
        )}
      </div>
      <footer className="statusbar">
        <div>
          <span>
            Page {currentPage} of {pageCount || "…"}
          </span>
          <span>{words.toLocaleString()} words</span>
          <span>
            {selected?.properties.direction === "rtl"
              ? "Right to left"
              : "Left to right"}
          </span>
        </div>
        <span className="status-message" role="status">
          {notice}
        </span>
        <div className="zoom-controls">
          <button onClick={() => fit("page")} title="Fit page">
            <Icon name="page" size={15} />
          </button>
          <button onClick={() => fit("width")}>Fit width</button>
          <button
            aria-label="Zoom out"
            onClick={() => setZoom((z) => Math.max(0.25, z - 0.1))}
          >
            −
          </button>
          <input
            aria-label="Zoom"
            type="range"
            min="25"
            max="200"
            step="5"
            value={Math.round(zoom * 100)}
            onChange={(e) => setZoom(Number(e.target.value) / 100)}
          />
          <button
            aria-label="Zoom in"
            onClick={() => setZoom((z) => Math.min(2, z + 0.1))}
          >
            +
          </button>
          <button onClick={() => setZoom(1)}>{Math.round(zoom * 100)}%</button>
        </div>
      </footer>
      <input
        ref={fileInput}
        hidden
        aria-label="Open DOCX"
        type="file"
        accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        onChange={async (e) => {
          const file = e.target.files?.[0];
          e.target.value = "";
          if (file && discardOK()) {
            if (!/\.docx$/i.test(file.name)) {
              setError("Choose a .docx file.");
              return;
            }
            if (file.size > MAX_BYTES) {
              setError("Choose a DOCX smaller than 50 MB.");
              return;
            }
            await open(new Uint8Array(await file.arrayBuffer()), file.name);
          }
        }}
      />
      {fileView && (
        <Backstage
          view={fileView}
          onView={setFileView}
          filename={filename}
          pageCount={pageCount}
          words={words}
          fontCount={fonts.required.length}
          document={document}
          busy={busy}
          onNew={() => {
            if (discardOK())
              void open(createBlankDocx(), "Untitled document.docx");
          }}
          onSample={() => void sample()}
          onOpen={() => fileInput.current?.click()}
          onSave={() => void save()}
          onOriginal={() => {
            if (session)
              download(
                session.original(),
                filename.replace(/\.docx$/i, "-original.docx"),
              );
          }}
          onPrint={() => void print()}
          onRename={() => setModal("rename")}
          onPane={(value) => {
            setFileView(null);
            setPane(value);
          }}
        />
      )}
      <WorkspaceDialogs
        modal={modal}
        onClose={() => setModal(null)}
        document={document}
        selected={selected}
        filename={filename}
        onFilename={setFilename}
        busy={busy}
        canPage={!!session?.canSetPage}
        execute={execute}
      />
      <HomeDialogs
        anchor={homeAnchor}
        contextStyle={contextStyle}
        onStyleContext={(id, anchor) => {
          setContextStyle(id);
          openHome("styleContext", anchor);
        }}
        onStyleAction={(name, id) => void styleAction(name, id)}
        dialog={homeDialog}
        onClose={() => {
          setHomeDialog(null);
          setTimeout(() => editor.focus(), 0);
        }}
        onDialog={openHome}
        selected={selected}
        format={activeRun()}
        families={fonts.required.map((font) => font.family)}
        styles={styleList}
        clipboard={clipboard}
        query={query}
        replacement={replacement}
        options={searchOptions}
        onRun={run}
        onParagraph={para}
        onAction={(name) => void action(name)}
        onText={insertText}
        onCommand={homeCommand}
        onFind={findText}
        onReplace={replaceText}
        onSort={(count, descending, numeric) => {
          const start = all.findIndex((p) => p.id === selectedId);
          void execute({
            type: "sort",
            ids: all.slice(start, start + count).map((p) => p.id!),
            descending,
            numeric,
            locale: document?.language,
          });
        }}
      />
    </div>
  );
}
