"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import {
  type DocumentStyle,
  type DocumentIR,
  type EditCommand,
  type ParagraphFormat,
  type RunFormat,
  similarFormatting,
  pageForParagraph,
  sectionTarget,
  physicalAlignment,
  tableContext,
  getParagraphs,
  columnWidth,
  type DrawingInline,
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
import { useCommentHighlight } from "./editor/useCommentHighlight";
import { Ribbon, type RibbonTab } from "./editor/Ribbon";
import { ContextRibbon, type ContextTab } from "./editor/ContextRibbon";
import {
  ObjectFormatDialog,
  TableFormatDialog,
} from "./editor/ObjectFormatDialog";
import { ObjectSelectionOverlay } from "./editor/ObjectSelectionOverlay";
import { IconsDialog } from "./editor/GraphicsGallery";
import { SymbolDialog } from "./editor/SymbolDialog";
import { Tool } from "./editor/Controls";
import { Icon } from "./editor/Icon";
import { FontPanel, useDocumentFonts } from "./editor/FontPanel";
import { WorkspaceDialogs, type EditorModal } from "./editor/WorkspaceDialogs";
import { Backstage } from "./editor/Backstage";
import { usePageEditor, type PageSelection } from "./editor/usePageEditor";
import { HomeRibbon } from "./editor/HomeRibbon";
import { HomeDialogs } from "./editor/HomeDialogs";
import { ClipboardPane } from "./editor/ClipboardPane";
import { useClipboard } from "./editor/useClipboard";
import { ParagraphRuler } from "./editor/ParagraphRuler";
import { InsertLinkDialog } from "./editor/InsertLinkDialog";
import { AuthoringDialog } from "./editor/AuthoringDialog";
import {
  ObjectPanel,
  type ObjectRef,
  documentObjects,
} from "./editor/ObjectPanel";
import { AssetDialog } from "./editor/AssetDialog";
import { StylesPane } from "./editor/StylesPane";
import type { MenuAnchor } from "./editor/HomeMenu";

type Pane =
  | "objects"
  | "comments"
  | "clipboard"
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
  objects: "Selection",
  comments: "Comments",
  clipboard: "Clipboard",
  find: "Navigation",
  replace: "Find and replace",
  edit: "Paragraph editor",
  fonts: "Document fonts",
  compatibility: "Document compatibility",
  styles: "Styles",
  outline: "Document outline",
};

export function DocxWorkspace() {
  const [assetDialog, setAssetDialog] = useState<
    "model3d" | "object" | "spreadsheet" | null
  >(null);
  const [authoring, setAuthoring] = useState<string | null>(null),
    [selectedObjects, setSelectedObjects] = useState<ObjectRef[]>([]);
  const [document, setDocument] = useState<DocumentIR | null>(null),
    [session, setSession] = useState<DocxSession | null>(null);
  const [filename, setFilename] = useState("Sample document.docx"),
    [busy, setBusy] = useState(true),
    [error, setError] = useState("");
  const [tab, setTab] = useState<RibbonTab | ContextTab>("Home"),
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
  const [editingStory, setEditingStory] = useState<"header" | "footer" | null>(
    null,
  );
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
  const {
    items: clipboard,
    setItems: setClipboard,
    read: readClipboard,
    error: clipboardError,
  } = useClipboard(pane === "clipboard");
  const [scope, setScope] = useState<PageSelection[]>([]);
  const fileInput = useRef<HTMLInputElement>(null),
    canvas = useRef<HTMLDivElement>(null),
    textInput = useRef<HTMLTextAreaElement>(null);
  const pictureInput = useRef<HTMLInputElement>(null);
  const pictureReplace = useRef<ObjectRef | null>(null);
  const equationDraft = useRef<(() => Promise<boolean>) | null>(null);
  const [objectFormat, setObjectFormat] = useState<string | null>(null),
    [tableFormat, setTableFormat] = useState<
      "properties" | "split" | "deleteCells" | null
    >(null),
    [equationSnippet, setEquationSnippet] = useState({ text: "", serial: 0 });
  const [hoveredComment, setHoveredComment] = useState<{
    id: string;
    start: number;
    end: number;
  } | null>(null);
  useCommentHighlight(canvas, hoveredComment);
  const textFileInput = useRef<HTMLInputElement>(null);
  const request = useRef(0),
    changing = useRef(false),
    activeSession = useRef<DocxSession | null>(null);
  const all = useMemo(() => paragraphs(document), [document]);
  const selected = useMemo(() => {
    const p = all.find((p) => p.id === selectedId);
    return p
      ? { ...p, properties: physicalAlignment(p.properties) }
      : undefined;
  }, [all, selectedId]);
  const objects = documentObjects(all);
  const pickedObject = objects.find((o) =>
    selectedObjects.some((ref) => ref.id === o.id && ref.index === o.index),
  );
  const pickedTable = useMemo(
    () => (document ? tableContext(document, selectedId) : undefined),
    [document, selectedId],
  );
  const contextualTabs: ContextTab[] = pickedObject?.drawing.equation
    ? ["Equation"]
    : pickedObject?.drawing.image &&
        !pickedObject.drawing.attachment &&
        !pickedObject.drawing.video
      ? ["Picture Format"]
      : pickedObject?.drawing.shape || pickedObject?.drawing.children
        ? ["Shape Format"]
        : [];
  if (pickedTable) contextualTabs.push("Table Design", "Table Layout");
  const ribbonTabs: (RibbonTab | ContextTab)[] = [
    "Home",
    "Insert",
    "Layout",
    ...contextualTabs,
  ];
  const shownTab = ribbonTabs.includes(tab) ? tab : "Home";
  const activeDocument = useMemo(() => {
    if (!document) return null;
    const target = sectionTarget(document, selectedId);
    return {
      ...document,
      page: pageForParagraph(document, selectedId),
      headers: target?.section?.headers ?? document.headers,
      footers: target?.section?.footers ?? document.footers,
    };
  }, [document, selectedId]);
  useEffect(() => {
    const root = canvas.current;
    if (!root) return;
    const paint = () =>
      root
        .querySelectorAll<HTMLElement>(
          "[data-drawing-owner][data-drawing-index]",
        )
        .forEach((node) =>
          node.toggleAttribute(
            "data-object-selected",
            selectedObjects.some(
              (o) =>
                o.id === node.dataset.drawingOwner &&
                o.index === Number(node.dataset.drawingIndex),
            ),
          ),
        );
    paint();
    const observer = new MutationObserver(paint);
    observer.observe(root, {
      subtree: true,
      childList: true,
      attributes: true,
      attributeFilter: ["data-layout-pending"],
    });
    return () => observer.disconnect();
  }, [selectedObjects, document]);
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
  useSearchHighlights(canvas, results, activeHit, searchVisible, scope);
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
      setSelectedObjects([]);
      setAuthoring(null);
      setAssetDialog(null);
      setEditingStory(null);
    },
    onCommand: execute,
    onCommands: executeMany,
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
      setSelectedObjects([]);
      setObjectFormat(null);
      setTableFormat(null);
      setAuthoring(null);
      setAssetDialog(null);
      pictureReplace.current = null;
      equationDraft.current = null;
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
      if (
        ["page", "columns", "layout", "story"].includes(command.type) &&
        !(command as { id?: string }).id
      )
        command = { ...command, id: selectedId || undefined } as EditCommand;
      if (command.type !== "text") await flushDraft(active);
      await active.execute(command);
      sync(
        active,
        [
          "splitParagraph",
          "joinParagraph",
          "insertDrawing",
          "insertImage",
          "insertField",
          "insertBreak",
          "sectionBreak",
          "insertTable",
          "table",
        ].includes(command.type)
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
      await active.executeBatch(
        commands.map((command) =>
          ["page", "columns", "layout", "story"].includes(command.type) &&
          !(command as { id?: string }).id
            ? ({ ...command, id: selectedId || undefined } as EditCommand)
            : command,
        ),
      );
      sync(
        active,
        commands.some((c) =>
          [
            "insertDrawing",
            "insertImage",
            "insertField",
            "insertBreak",
            "sectionBreak",
            "table",
          ].includes(c.type),
        )
          ? active.selection
          : undefined,
      );
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
    if (changing.current) return;
    if (equationDraft.current && !(await equationDraft.current())) return;
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
        normalizeParagraph: true,
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
    if (name.startsWith("shape:")) {
      const kind = name.slice(6);
      await insertGraphic({
        type: "drawing",
        widthPt: kind === "line" ? 144 : 108,
        heightPt: kind === "line" ? 2 : 72,
        alt: kind,
        shape: {
          kind,
          fill: kind === "line" ? undefined : "#4472C4",
          stroke: "#2F528F",
          strokeWidthPt: 1,
        },
      });
      return;
    }
    if (name.startsWith("equation:")) {
      if (name === "equation:new") {
        const previous = new Set(objects.map((o) => `${o.id}-${o.index}`));
        if (
          await execute({
            type: "insertDrawing",
            id: selectedId || undefined,
            drawing: {
              type: "drawing",
              widthPt: 120,
              heightPt: 36,
              alt: "Equation",
              equation: { expression: "x" },
            },
          })
        ) {
          const o = documentObjects(
            paragraphs(activeSession.current!.document),
          ).find(
            (o) => o.drawing.equation && !previous.has(`${o.id}-${o.index}`),
          );
          if (o) {
            setSelectedObjects([{ id: o.id, index: o.index }]);
            setSelectedId(o.id);
            setTab("Equation");
            setEquationSnippet((p) => ({
              text: "__focus__",
              serial: p.serial + 1,
            }));
          }
        }
        return;
      }
      setEquationSnippet((p) => ({
        text: name === "equation:focus" ? "__focus__" : name.slice(16),
        serial: p.serial + 1,
      }));
      return;
    }
    if (name === "picture:replace") {
      pictureReplace.current = pickedObject
        ? { id: pickedObject.id, index: pickedObject.index }
        : null;
      pictureInput.current?.click();
      return;
    }
    if (name === "picture:options" || name === "picture:crop") {
      setObjectFormat(name.endsWith("crop") ? "Crop" : "Size");
      return;
    }
    if (name.startsWith("table:") && pickedTable && !/^table:\d/.test(name)) {
      const operation = name.slice(6);
      if (
        operation === "properties" ||
        operation === "split" ||
        operation === "deleteCells"
      ) {
        setTableFormat(operation);
        return;
      }
      if (operation.startsWith("select:")) {
        const what = operation.slice(7),
          rows = pickedTable.table.rows;
        const cells = rows.flatMap((r, i) =>
          r.cells.filter(
            (c) =>
              !c.continuation &&
              (what === "table" ||
                (what === "row" && i === pickedTable.row) ||
                (what === "column" &&
                  c.column <= pickedTable.cell.column &&
                  c.column + c.columnSpan > pickedTable.cell.column) ||
                (what === "cell" && c === pickedTable.cell)),
          ),
        );
        const selected = cells
          .flatMap((c) => getParagraphs(c.blocks))
          .map((p) => ({ id: p.id!, start: 0, end: paragraphText(p).length }));
        setScope(selected);
        return;
      }
      if (operation.startsWith("align:")) {
        const [, verticalAlign, alignment] = operation.split(":");
        const command = makeTableCommand({
          format: {
            verticalAlign: verticalAlign as "top" | "center" | "bottom",
          },
        });
        const r = command.range!;
        await executeMany([
          command,
          ...pickedTable.table.rows.flatMap((row, i) =>
            i >= r.top && i <= r.bottom
              ? row.cells
                  .filter((c) => c.column >= r.left && c.column <= r.right)
                  .flatMap((c) => getParagraphs(c.blocks))
                  .map((p) => ({
                    type: "paragraph" as const,
                    id: p.id!,
                    format: {
                      alignment: alignment as "left" | "center" | "right",
                    },
                  }))
              : [],
          ),
        ]);
        return;
      }
      await tableCommand({
        action: operation as Extract<EditCommand, { type: "table" }>["action"],
      });
      return;
    }
    if (name === "textFromFile") {
      textFileInput.current?.click();
      return;
    }
    if (name.startsWith("section:")) {
      const picked = editor.getSelection() ?? {
        id: selectedId,
        start: range?.start ?? 0,
        end: range?.end ?? 0,
      };
      await execute({
        type: "sectionBreak",
        ...picked,
        kind: name.slice(8) as
          | "nextPage"
          | "continuous"
          | "evenPage"
          | "oddPage",
      });
      return;
    }
    if (name === "object:edit") {
      const d = objects.find((o) =>
        selectedObjects.some((ref) => ref.id === o.id && ref.index === o.index),
      )?.drawing;
      if (d?.chart) setAuthoring("chart");
      else if (d?.equation) {
        setTab("Equation");
        setEquationSnippet((p) => ({
          text: "__focus__",
          serial: p.serial + 1,
        }));
      } else if (d?.spreadsheet) setAssetDialog("spreadsheet");
      return;
    }
    if (name.startsWith("asset:")) {
      setAssetDialog(name.slice(6) as "model3d" | "object" | "spreadsheet");
      return;
    }
    if (name.startsWith("break:")) {
      const picked = editor.getSelection() ?? {
        id: selectedId,
        start: range?.start ?? 0,
        end: range?.end ?? 0,
      };
      if (name === "break:wrap") insertText("\n");
      else {
        if (
          await execute({
            type: "insertBreak",
            ...picked,
            kind: name.slice(6) as "page" | "column",
          })
        ) {
          const target = activeSession.current?.selection;
          if (target) editor.select(target, true);
        }
      }
      return;
    }
    if (name === "cover:remove") {
      await execute({
        type: "cover",
        title: "",
        subtitle: "",
        author: "",
        date: "",
        design: "simple",
        remove: true,
      });
      return;
    }
    if (name === "screenshot") {
      let stream: MediaStream | undefined;
      try {
        stream = await navigator.mediaDevices.getDisplayMedia({ video: true });
        const video = window.document.createElement("video");
        video.srcObject = stream;
        await video.play();
        const canvas = window.document.createElement("canvas");
        canvas.width = video.videoWidth;
        canvas.height = video.videoHeight;
        canvas.getContext("2d")!.drawImage(video, 0, 0);
        const available = columnWidth(
            activeDocument ? pageForParagraph(activeDocument, selectedId) : {},
          ),
          scale = Math.min(0.75, available / canvas.width);
        await execute({
          type: "insertImage",
          id: selectedId,
          dataUrl: canvas.toDataURL("image/png"),
          widthPt: canvas.width * scale,
          heightPt: canvas.height * scale,
          alt: "Screen capture",
        });
      } catch (error) {
        setNotice(
          error instanceof Error ? error.message : "Screen capture cancelled.",
        );
      } finally {
        stream?.getTracks().forEach((track) => track.stop());
      }
      return;
    }
    if (name === "video") {
      setAuthoring("video");
      return;
    }
    if (name.startsWith("dialog:")) {
      setAuthoring(name.slice(7));
      return;
    }
    if (name === "objects" || name === "comments") {
      setPane(name);
      return;
    }
    if (name === "object:options") {
      if (pickedObject) setObjectFormat("Position");
      else setPane("objects");
      return;
    }
    if (name.startsWith("object:")) {
      if (
        await execute({
          type: "arrange",
          objects: selectedObjects,
          action: name.slice(7),
        })
      ) {
        if (
          ["object:group", "object:ungroup", "object:delete"].includes(name)
        ) {
          setSelectedObjects([]);
          if (canvas.current) delete canvas.current.dataset.objectSelection;
        }
      }
      return;
    }
    if (name === "layout:mirror") {
      await executeMany([
        {
          type: "page",
          widthPt: activeDocument?.page?.widthPt ?? 595.3,
          heightPt: activeDocument?.page?.heightPt ?? 841.9,
          margins: { top: 72, bottom: 72, left: 90, right: 72 },
        },
        { type: "layout", format: { mirrorMargins: true } },
      ]);
      return;
    }
    if (name.startsWith("lineNumbers:")) {
      const mode = name.slice(12);
      await execute({
        type: "layout",
        format: {
          lineNumbers:
            mode === "none"
              ? undefined
              : {
                  mode: mode as "page" | "section" | "continuous",
                  start: 1,
                  countBy: 1,
                  distancePt: 18,
                },
        },
      });
      return;
    }
    if (name.startsWith("hyphenation:")) {
      await execute({
        type: "layout",
        format: { hyphenation: name.slice(12) as "auto" | "none" },
      });
      return;
    }
    if (name.startsWith("columns:")) {
      await execute({ type: "columns", count: Number(name.slice(8)) });
      return;
    }
    if (name === "picture") {
      pictureInput.current?.click();
      return;
    }
    if (name === "link") {
      setModal("link");
      return;
    }
    if (name === "clipboard") {
      setPane("clipboard");
      void readClipboard();
      return;
    }
    if (name.startsWith("story:")) {
      const [, kind, operation] = name.split(":");
      await editStory(kind as "header" | "footer", operation);
      return;
    }
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
      const ranges = (
        name === "selectAll"
          ? all.map((p) => ({
              id: p.id!,
              start: 0,
              end: paragraphText(p).length,
            }))
          : similarFormatting(
              all,
              selectedId,
              editor.getSelection()?.start ?? 0,
            )
      ).filter((r) => session?.capability(r.id).format);
      setScope(ranges);
      setNotice(`${ranges.length} text ranges selected.`);
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
    const w = activeDocument?.page?.widthPt ?? 595.3,
      h = activeDocument?.page?.heightPt ?? 841.9;
    if (name.startsWith("margins:")) {
      const [top, right, bottom, left] = name.slice(8).split(",").map(Number);
      page(w, h, { top, right, bottom, left });
      return;
    }
    if (name.startsWith("table:")) {
      const [rows, columns] = name.slice(6).split(",").map(Number);
      await execute({
        type: "insertTable",
        id: selectedId || undefined,
        rows,
        columns,
      });
      return;
    }
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
      [
        "fonts",
        "find",
        "replace",
        "edit",
        "styles",
        "outline",
        "clipboard",
      ].includes(name)
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
    if (!canvas.current || canvas.current.dataset.objectSelection) return;
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
  function makeTableCommand(
    patch: Omit<Extract<EditCommand, { type: "table" }>, "type" | "id">,
  ): Extract<EditCommand, { type: "table" }> {
    const context = pickedTable!;
    const selected = scope
      .map((s) => document && tableContext(document, s.id))
      .filter(
        (c): c is NonNullable<typeof c> => !!c && c.table === context.table,
      );
    const range =
      patch.range ??
      (selected.length
        ? {
            top: Math.min(...selected.map((c) => c.row)),
            bottom: Math.max(
              ...selected.map((c) => c.row + c.cell.rowSpan - 1),
            ),
            left: Math.min(...selected.map((c) => c.cell.column)),
            right: Math.max(
              ...selected.map((c) => c.cell.column + c.cell.columnSpan - 1),
            ),
          }
        : {
            top: context.row,
            bottom: context.row + context.cell.rowSpan - 1,
            left: context.cell.column,
            right: context.cell.column + context.cell.columnSpan - 1,
          });
    if (
      !scope.length &&
      ["distributeRows", "distributeColumns"].includes(patch.action ?? "")
    ) {
      range.top = 0;
      range.bottom = context.table.rows.length - 1;
      range.left = 0;
      range.right =
        Math.max(
          context.table.columnsPt.length,
          ...context.table.rows.flatMap((r) =>
            r.cells.map((c) => c.column + c.columnSpan),
          ),
        ) - 1;
    }
    return { type: "table", id: selectedId, ...patch, range };
  }
  async function tableCommand(
    patch: Omit<Extract<EditCommand, { type: "table" }>, "type" | "id">,
  ) {
    return pickedTable ? execute(makeTableCommand(patch)) : false;
  }
  async function changeObject(
    drawing: import("@local/docx-ir-engine/core").DrawingInline,
  ) {
    return pickedObject
      ? execute({
          type: "drawing",
          id: pickedObject.id,
          index: pickedObject.index,
          drawing,
        })
      : false;
  }
  async function insertGraphic(drawing: DrawingInline) {
    const previous = new Set(objects.map((o) => `${o.id}-${o.index}`));
    if (
      !(await execute({
        type: "insertDrawing",
        id: selectedId || undefined,
        drawing,
      }))
    )
      return false;
    const o = documentObjects(paragraphs(activeSession.current!.document)).find(
      (o) => !previous.has(`${o.id}-${o.index}`),
    );
    if (o) {
      setSelectedObjects([{ id: o.id, index: o.index }]);
      setSelectedId(o.id);
      setScope([]);
      if (canvas.current) canvas.current.dataset.objectSelection = "true";
      setTab(o.drawing.image ? "Picture Format" : "Shape Format");
    }
    return true;
  }
  async function editStory(
    kind: "header" | "footer",
    operation = "edit",
    pageNumber = currentPage,
    scopeId = selectedId,
  ) {
    const active = activeSession.current;
    if (!active) return;
    const page = pageForParagraph(active.document, scopeId);
    const storyBlocks = () => {
      const target = sectionTarget(active.document, scopeId);
      return target?.section?.[stories] ?? active.document[stories];
    };
    const variant =
      page.differentFirstPage && pageNumber === 1
        ? "first"
        : page.evenAndOddHeaders && pageNumber % 2 === 0
          ? "even"
          : "default";
    const stories = kind === "header" ? "headers" : "footers";
    if (operation === "remove") {
      await execute({
        type: "story",
        id: scopeId,
        kind,
        variant,
        remove: true,
      });
      setEditingStory(null);
      return;
    }
    if (operation === "number" || !storyBlocks()?.[variant]?.length) {
      if (
        !(await execute({
          type: "story",
          id: scopeId,
          kind,
          variant,
          pageNumber: operation === "number",
          alignment: "center",
        }))
      )
        return;
    }
    setEditingStory(kind);
    const blocks = storyBlocks()?.[variant];
    let p = blocks
      ? paragraphs({
          ...active.document,
          blocks,
          headers: undefined,
          footers: undefined,
        }).find((p) => active.capability(p.id!).text)
      : undefined;
    if (!p && operation === "edit") {
      if (
        !(await execute({
          type: "story",
          id: scopeId,
          kind,
          variant,
          appendText: true,
        }))
      )
        return;
      p = paragraphs({
        ...active.document,
        blocks: storyBlocks()?.[variant] ?? [],
        headers: undefined,
        footers: undefined,
      }).find((p) => active.capability(p.id!).text);
    }
    if (p) {
      setSelectedId(p.id!);
      setDraft(paragraphText(p));
      setRange({ start: 0, end: 0 });
      editor.select({ id: p.id!, start: 0, end: 0 }, true);
    }
  }
  function fit(mode: "width" | "page") {
    if (!canvas.current) return;
    const w = ((activeDocument?.page?.widthPt ?? 595.3) * 4) / 3,
      h = ((activeDocument?.page?.heightPt ?? 841.9) * 4) / 3;
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
          {ribbonTabs.map((name, index) => (
            <button
              id={`tab-${name.replaceAll(" ", "-")}`}
              key={name}
              role="tab"
              aria-selected={shownTab === name}
              aria-controls={`ribbon-${name.replaceAll(" ", "-")}`}
              tabIndex={shownTab === name ? 0 : -1}
              className={`${shownTab === name ? "selected-tab" : ""} ${contextualTabs.includes(name as ContextTab) ? "context-tab" : ""}`}
              onClick={() => setTab(name)}
              onKeyDown={(e) => {
                if (e.key === "ArrowRight" || e.key === "ArrowLeft") {
                  e.preventDefault();
                  const names = ribbonTabs,
                    next =
                      names[
                        (index +
                          (e.key === "ArrowRight" ? 1 : names.length - 1)) %
                          names.length
                      ];
                  setTab(next);
                  window.document
                    .getElementById(`tab-${next.replaceAll(" ", "-")}`)
                    ?.focus();
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
      {shownTab === "Home" ? (
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
          openMenu={homeDialog}
          onDismissMenu={() => setHomeDialog(null)}
        />
      ) : contextualTabs.includes(shownTab as ContextTab) ? (
        <ContextRibbon
          tab={shownTab as ContextTab}
          object={pickedObject?.drawing}
          table={pickedTable}
          disabled={busy || !!session?.lockedReason}
          onDrawing={(d) => void changeObject(d)}
          onTable={(p) => void tableCommand(p)}
          onAction={(name) => void action(name)}
        />
      ) : (
        <Ribbon
          objectCount={selectedObjects.length}
          hasGroup={objects.some(
            (o) =>
              selectedObjects.some(
                (ref) => ref.id === o.id && ref.index === o.index,
              ) && !!o.drawing.children,
          )}
          tab={shownTab as RibbonTab}
          document={activeDocument}
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
        {editingStory && (
          <button
            onClick={() => {
              setEditingStory(null);
              const p = all.find((p) => document?.blocks.includes(p));
              if (p) jump(p.id!);
            }}
          >
            Close Header and Footer
          </button>
        )}
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
          data-editing-story={editingStory ?? undefined}
          data-object-selection={selectedObjects.length ? "true" : undefined}
          aria-label="Document pages"
          aria-busy={busy}
          tabIndex={-1}
          onKeyDownCapture={(event) => {
            if (!selectedObjects.length) return;
            const input = event.target as HTMLElement;
            if (input.closest("input,textarea,button,select")) return;
            if (["Delete", "Backspace"].includes(event.key)) {
              event.preventDefault();
              event.stopPropagation();
              void action("object:delete");
            } else if (event.key === "Escape") {
              event.preventDefault();
              event.stopPropagation();
              setSelectedObjects([]);
            } else if (
              ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"].includes(
                event.key,
              )
            ) {
              event.preventDefault();
              event.stopPropagation();
              const step = event.shiftKey ? 10 : 1;
              void action(
                `object:move:${event.key === "ArrowLeft" ? -step : event.key === "ArrowRight" ? step : 0},${event.key === "ArrowUp" ? -step : event.key === "ArrowDown" ? step : 0}`,
              );
            }
          }}
          onMouseDownCapture={(event) => {
            const target = event.target as Element,
              object = target.closest<HTMLElement>(
                "[data-drawing-owner][data-drawing-index]",
              );
            if (target.closest("[data-editor-overlay]")) return;
            if (
              !object ||
              object.contains(target.closest("[data-paragraph-id]")) ||
              target.closest("a[download]")
            ) {
              if (equationDraft.current) void equationDraft.current();
              setSelectedObjects([]);
              if (canvas.current) delete canvas.current.dataset.objectSelection;
              const cell = target.closest<HTMLElement>("td");
              const p =
                cell &&
                (target.closest<HTMLElement>("[data-paragraph-id]") ??
                  cell.querySelector<HTMLElement>("[data-paragraph-id]"));
              if (p?.dataset.paragraphId) {
                setSelectedId(p.dataset.paragraphId);
                if (!pickedTable) setTab("Table Layout");
              }
              return;
            }
            event.preventDefault();
            if (canvas.current) canvas.current.dataset.objectSelection = "true";
            canvas.current?.focus({ preventScroll: true });
            const ref = {
              id: object.dataset.drawingOwner!,
              index: Number(object.dataset.drawingIndex),
            };
            setSelectedId(ref.id);
            setDraft(
              all.find((p) => p.id === ref.id)
                ? paragraphText(all.find((p) => p.id === ref.id)!)
                : "",
            );
            setScope([]);
            const drawing = objects.find(
              (o) => o.id === ref.id && o.index === ref.index,
            )?.drawing;
            if (drawing?.equation) {
              setTab("Equation");
              setEquationSnippet((p) => ({
                text: "__focus__",
                serial: p.serial + 1,
              }));
            } else if (drawing?.image && !drawing.attachment && !drawing.video)
              setTab("Picture Format");
            else if (drawing?.shape || drawing?.children)
              setTab("Shape Format");
            setSelectedObjects((previous) =>
              event.ctrlKey || event.metaKey
                ? previous.some((o) => o.id === ref.id && o.index === ref.index)
                  ? previous.filter(
                      (o) => o.id !== ref.id || o.index !== ref.index,
                    )
                  : [...previous, ref]
                : [ref],
            );
          }}
          onMouseUp={capturePageSelection}
          onClick={(event) => {
            if ((event.target as Element).closest("a:not([download])"))
              event.preventDefault();
          }}
          onDoubleClick={(event) => {
            const target = event.target as HTMLElement;
            const object = target.closest<HTMLElement>(
              "[data-drawing-owner][data-drawing-index]",
            );
            if (object) {
              const d = objects.find(
                (o) =>
                  o.id === object.dataset.drawingOwner &&
                  o.index === Number(object.dataset.drawingIndex),
              )?.drawing;
              if (d?.chart || d?.equation || d?.spreadsheet) {
                setSelectedObjects([
                  {
                    id: object.dataset.drawingOwner!,
                    index: Number(object.dataset.drawingIndex),
                  },
                ]);
                if (d.chart) setAuthoring("chart");
                else if (d.equation) {
                  setTab("Equation");
                  setEquationSnippet((p) => ({
                    text: "__focus__",
                    serial: p.serial + 1,
                  }));
                } else setAssetDialog("spreadsheet");
                return;
              }
            }
            const page = target.closest<HTMLElement>("[data-docx-page]");
            const sheet = target.closest<HTMLElement>("[data-docx-engine]");
            if (!page || !sheet || !document) return;
            const scopeId =
              page.querySelector<HTMLElement>(
                "[data-docx-body] [data-paragraph-id]",
              )?.dataset.paragraphId ?? selectedId;
            const layout = pageForParagraph(document, scopeId);
            const bounds = sheet.getBoundingClientRect(),
              y = event.clientY - bounds.top,
              scale = bounds.height / (layout.heightPt ?? 841.9);
            const story =
              target.closest<HTMLElement>("[data-docx-story]")?.dataset
                .docxStory;
            if (story === "header" || y < (layout.margins?.top ?? 72) * scale)
              void editStory(
                "header",
                "edit",
                Number(page.dataset.docxSectionPage ?? page.dataset.docxPage),
                scopeId,
              );
            else if (
              story === "footer" ||
              y > bounds.height - (layout.margins?.bottom ?? 72) * scale
            )
              void editStory(
                "footer",
                "edit",
                Number(page.dataset.docxSectionPage ?? page.dataset.docxPage),
                scopeId,
              );
            else setEditingStory(null);
          }}
        >
          {document && (
            <ParagraphRuler
              container={canvas}
              document={activeDocument!}
              selected={selected}
              zoom={zoom}
              disabled={busy}
              onChange={para}
            />
          )}
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
          {selectedObjects.length === 1 && pickedObject && (
            <ObjectSelectionOverlay
              pendingDraft={equationDraft}
              container={canvas}
              object={pickedObject}
              onChange={changeObject}
              onAction={(name) => void action(name)}
              snippet={equationSnippet}
            />
          )}
        </main>
        {pane && (
          <aside
            className={`task-pane ${pane === "find" || pane === "clipboard" ? "task-pane-leading" : ""}`}
            aria-label={titles[pane]}
          >
            <div className="pane-title">
              <h2>{titles[pane]}</h2>
              <button aria-label="Close pane" onClick={() => setPane(null)}>
                <Icon name="close" size={18} />
              </button>
            </div>
            <div className="pane-content">
              {pane === "fonts" && <FontPanel fonts={fonts} />}
              {pane === "objects" && (
                <ObjectPanel
                  all={all}
                  selected={selectedObjects}
                  onSelect={(ref, multi) =>
                    setSelectedObjects((previous) =>
                      multi
                        ? [
                            ...previous.filter(
                              (o) => o.id !== ref.id || o.index !== ref.index,
                            ),
                            ref,
                          ]
                        : [ref],
                    )
                  }
                  onChange={(ref, drawing) =>
                    void execute({ type: "drawing", ...ref, drawing })
                  }
                  onAction={(name) => void action(name)}
                />
              )}
              {pane === "comments" &&
                all.flatMap((p) =>
                  (p.comments ?? []).map((c) => (
                    <article
                      className="comment-card"
                      key={`${p.id}-${c.key}`}
                      tabIndex={0}
                      onMouseEnter={() =>
                        setHoveredComment({
                          id: p.id!,
                          start: c.start,
                          end: c.end,
                        })
                      }
                      onMouseLeave={() => setHoveredComment(null)}
                      onFocus={() =>
                        setHoveredComment({
                          id: p.id!,
                          start: c.start,
                          end: c.end,
                        })
                      }
                      onBlur={() => setHoveredComment(null)}
                    >
                      <b>{c.author}</b>
                      <small>{new Date(c.date).toLocaleString()}</small>
                      <textarea
                        defaultValue={c.text}
                        onBlur={(e) =>
                          void execute({
                            type: "comment",
                            id: p.id!,
                            ...c,
                            text: e.target.value,
                          })
                        }
                      />
                      <button onClick={() => jump(p.id!)}>Go to text</button>
                      <button
                        onClick={() =>
                          void execute({
                            type: "comment",
                            id: p.id!,
                            ...c,
                            remove: true,
                          })
                        }
                      >
                        Delete
                      </button>
                    </article>
                  )),
                )}
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
                    Supported edits include text and styles, tables, pictures,
                    shapes, charts, equations, bookmarks, comments, embedded
                    files, headers/footers and section page setup.
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
              {pane === "clipboard" && (
                <ClipboardPane
                  items={clipboard}
                  disabled={busy || !canFormat}
                  onPaste={insertText}
                  onClear={() => setClipboard([])}
                  onRead={() => void readClipboard()}
                  error={clipboardError}
                />
              )}
              {pane === "styles" && (
                <StylesPane
                  styles={styleList}
                  current={selected?.properties.styleId}
                  disabled={busy || !canFormat}
                  onApply={(id) => para({ styleId: id })}
                  onContext={(id, anchor) => {
                    setContextStyle(id);
                    openHome("styleContext", anchor);
                  }}
                />
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
        ref={pictureInput}
        type="file"
        accept="image/png,image/jpeg"
        aria-label="Insert Picture"
        hidden
        onChange={async (event) => {
          const file = event.target.files?.[0];
          event.target.value = "";
          if (!file) return;
          try {
            if (file.size > 20_000_000)
              throw new Error("Choose an image smaller than 20 MB.");
            const bitmap = await createImageBitmap(file),
              available = columnWidth(
                document ? pageForParagraph(document, selectedId) : {},
              );
            const scale = Math.min(0.75, available / bitmap.width),
              widthPt = bitmap.width * scale,
              heightPt = bitmap.height * scale;
            bitmap.close();
            const dataUrl = await new Promise<string>((resolve, reject) => {
              const reader = new FileReader();
              reader.onload = () => resolve(String(reader.result));
              reader.onerror = () =>
                reject(new Error("Could not read the image."));
              reader.readAsDataURL(file);
            });
            const replace = pictureReplace.current;
            pictureReplace.current = null;
            if (replace) {
              const original = documentObjects(
                paragraphs(activeSession.current!.document),
              ).find((o) => o.id === replace.id && o.index === replace.index);
              if (original?.drawing.image)
                await execute({
                  type: "drawing",
                  ...replace,
                  drawing: {
                    ...original.drawing,
                    alt: file.name,
                    image: { ...original.drawing.image, dataUrl },
                  },
                });
              return;
            }
            if (
              await execute({
                type: "insertImage",
                id: selectedId || undefined,
                dataUrl,
                widthPt,
                heightPt,
                alt: file.name,
              })
            ) {
              const target = activeSession.current?.selection;
              if (target) {
                setSelectedId(target.id);
                setDraft("");
                editor.select(target, true);
              }
            }
          } catch (error) {
            setError(
              error instanceof Error
                ? error.message
                : "Could not insert the picture.",
            );
          }
        }}
      />
      {modal === "link" && selected && (
        <InsertLinkDialog
          selected={selected}
          range={range ?? { start: 0, end: 0 }}
          execute={execute}
          onClose={() => {
            setModal(null);
            setTimeout(() => editor.focus(), 0);
          }}
        />
      )}
      <input
        ref={textFileInput}
        hidden
        type="file"
        aria-label="Insert text file"
        accept=".txt,.md,.csv,text/plain,text/csv"
        onChange={async (event) => {
          const file = event.target.files?.[0];
          event.target.value = "";
          if (!file) return;
          try {
            if (file.size > 2_000_000)
              throw new Error("Choose a text file smaller than 2 MB.");
            const lines = (await file.text())
              .replace(/\r\n?/g, "\n")
              .split("\n");
            await executeMany(
              lines.reverse().map((text) => ({
                type: "insertParagraph",
                id: selectedId || undefined,
                text,
              })),
            );
          } catch (error) {
            setError(
              error instanceof Error ? error.message : "Could not insert file.",
            );
          }
        }}
      />
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
        document={activeDocument}
        selected={selected}
        filename={filename}
        onFilename={setFilename}
        busy={busy}
        canPage={!!session?.canSetPage}
        execute={execute}
      />
      {objectFormat && pickedObject && (
        <ObjectFormatDialog
          drawing={pickedObject.drawing}
          initialTab={objectFormat}
          onApply={changeObject}
          onClose={() => setObjectFormat(null)}
        />
      )}
      {tableFormat && pickedTable && (
        <TableFormatDialog
          context={pickedTable}
          mode={tableFormat}
          onApply={tableCommand}
          onClose={() => setTableFormat(null)}
        />
      )}
      {assetDialog && (
        <AssetDialog
          mode={assetDialog}
          id={selectedId || undefined}
          object={objects.find((o) =>
            selectedObjects.some(
              (ref) => ref.id === o.id && ref.index === o.index,
            ),
          )}
          execute={executeMany}
          onClose={() => setAssetDialog(null)}
        />
      )}
      {authoring === "icons" && (
        <IconsDialog
          onClose={() => setAuthoring(null)}
          onInsert={async (ids) => {
            const commands: EditCommand[] = ids.map((id) => ({
              type: "insertDrawing",
              id: selectedId || undefined,
              drawing: {
                type: "drawing",
                widthPt: 54,
                heightPt: 54,
                alt: id,
                shape: {
                  kind: `icon_${id}`,
                  fill: "#000000",
                  strokeWidthPt: 0,
                },
              },
            }));
            return executeMany(commands);
          }}
        />
      )}
      {authoring === "symbols" && (
        <SymbolDialog
          families={fonts.required.map((f) => f.family)}
          onClose={() => setAuthoring(null)}
          onInsert={async (text, font) => {
            if (!selected) return false;
            const picked = editor.getSelection() ?? {
              id: selectedId,
              start: range?.start ?? 0,
              end: range?.end ?? 0,
            };
            const done = await execute({
              type: "text",
              ...picked,
              text,
              format: { fontFamily: font, complexScriptFont: font },
            });
            if (done)
              editor.select({
                id: picked.id,
                start: picked.start + text.length,
                end: picked.start + text.length,
              });
            return done;
          }}
        />
      )}
      {authoring && !["icons", "symbols"].includes(authoring) && document && (
        <AuthoringDialog
          mode={authoring}
          object={objects.find((o) =>
            selectedObjects.some(
              (ref) => ref.id === o.id && ref.index === o.index,
            ),
          )}
          document={activeDocument!}
          all={all}
          selected={selected}
          range={range ?? { start: 0, end: 0 }}
          execute={async (commands) => {
            const changed = await executeMany(commands);
            if (changed && authoring === "comment") setPane("comments");
            return changed;
          }}
          onText={insertText}
          onClose={() => {
            setAuthoring(null);
            setTimeout(() => editor.focus(), 0);
          }}
          onNavigate={(id, start, end) => {
            setAuthoring(null);
            jump(id);
            editor.select({ id, start, end }, true);
          }}
        />
      )}
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
