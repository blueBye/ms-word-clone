import {
  useLayoutEffect,
  useRef,
  useState,
  useEffect,
  type RefObject,
} from "react";
import { createPortal } from "react-dom";
import { parseEquation, type DrawingInline } from "@local/docx-ir-engine/core";
import { EquationView } from "@local/docx-ir-engine/react";
import type { ObjectRef } from "./ObjectPanel";
type Bounds = {
  left: number;
  top: number;
  width: number;
  height: number;
  scale: number;
  pageLeft: number;
  pageTop: number;
};
export function ObjectSelectionOverlay({
  container,
  object,
  onChange,
  onAction,
  snippet,
  pendingDraft,
}: {
  container: RefObject<HTMLElement | null>;
  object?: ObjectRef & { drawing: DrawingInline };
  onChange: (d: DrawingInline) => Promise<boolean>;
  onAction: (name: string) => void;
  snippet: { text: string; serial: number };
  pendingDraft: RefObject<(() => Promise<boolean>) | null>;
}) {
  const [bounds, setBounds] = useState<Bounds | null>(null),
    [draft, setDraft] = useState<Bounds | null>(null);
  const drag = useRef<{
    x: number;
    y: number;
    handle: string;
    box: Bounds;
    drawing: DrawingInline;
  } | null>(null);
  useLayoutEffect(() => {
    const host = container.current;
    if (!host || !object) {
      setBounds(null);
      return;
    }
    let frame = 0;
    const measure = () => {
      frame = 0;
      const node = [
          ...host.querySelectorAll<HTMLElement>(
            `[data-drawing-owner="${CSS.escape(object.id)}"][data-drawing-index="${object.index}"]`,
          ),
        ].find(
          (n) =>
            n.closest("[data-docx-page]") &&
            n.style.display !== "none" &&
            !n.closest("[data-editor-island]"),
        ),
        page = node?.closest<HTMLElement>("[data-docx-engine]");
      if (!node || !page) {
        setBounds(null);
        return;
      }
      const b = node.getBoundingClientRect(),
        p = page.getBoundingClientRect(),
        h = host.getBoundingClientRect();
      const next =
        b.bottom < h.top + 32 ||
        b.top > h.bottom ||
        b.right < h.left ||
        b.left > h.right
          ? null
          : {
              left: b.left,
              top: b.top,
              width: b.width,
              height: b.height,
              scale: b.width / object.drawing.widthPt,
              pageLeft: p.left,
              pageTop: p.top,
            };
      setBounds((previous) =>
        JSON.stringify(previous) === JSON.stringify(next) ? previous : next,
      );
    };
    const schedule = () => {
      if (!frame) frame = requestAnimationFrame(measure);
    };
    measure();
    const observer = new MutationObserver(schedule);
    observer.observe(host, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["style", "data-layout-pending"],
    });
    host.addEventListener("scroll", schedule);
    window.addEventListener("resize", schedule);
    return () => {
      observer.disconnect();
      host.removeEventListener("scroll", schedule);
      window.removeEventListener("resize", schedule);
      cancelAnimationFrame(frame);
    };
  }, [container, object?.id, object?.index, object?.drawing]);
  if (!bounds || !object || typeof document === "undefined") return null;
  const box = draft ?? bounds,
    d = object.drawing;
  const resized = (x: number, y: number, shift: boolean) => {
    const current = drag.current!;
    const deltaX = x - current.x,
      deltaY = y - current.y,
      h = current.handle;
    if (h === "move")
      return {
        ...current.box,
        left: current.box.left + deltaX,
        top: current.box.top + deltaY,
      };
    let width =
        current.box.width +
        (h.includes("e") ? deltaX : h.includes("w") ? -deltaX : 0),
      height =
        current.box.height +
        (h.includes("s") ? deltaY : h.includes("n") ? -deltaY : 0);
    width = Math.max(12, width);
    height = Math.max(12, height);
    if (h.length === 2 && !shift) {
      const ratio = current.box.width / current.box.height;
      if (Math.abs(deltaX) > Math.abs(deltaY)) height = width / ratio;
      else width = height * ratio;
    }
    return {
      ...current.box,
      width,
      height,
      left:
        current.box.left + (h.includes("w") ? current.box.width - width : 0),
      top:
        current.box.top + (h.includes("n") ? current.box.height - height : 0),
    };
  };
  return createPortal(
    <div
      data-editor-overlay=""
      className="object-selection-overlay"
      style={{
        position: "fixed",
        left: box.left,
        top: box.top,
        width: box.width,
        height: box.height,
      }}
    >
      {!d.equation &&
        ["nw", "n", "ne", "e", "se", "s", "sw", "w"].map((handle) => (
          <button
            key={handle}
            type="button"
            className={`object-resize object-resize-${handle}`}
            aria-label={`Resize object ${handle}`}
            title="Drag to resize; hold Shift to change proportions"
            onPointerDown={(e) => {
              e.preventDefault();
              e.stopPropagation();
              e.currentTarget.setPointerCapture(e.pointerId);
              drag.current = {
                x: e.clientX,
                y: e.clientY,
                handle,
                box: bounds,
                drawing: structuredClone(d),
              };
            }}
            onPointerMove={(e) => {
              if (drag.current)
                setDraft(resized(e.clientX, e.clientY, e.shiftKey));
            }}
            onPointerUp={async (e) => {
              if (!drag.current) return;
              const result = resized(e.clientX, e.clientY, e.shiftKey),
                current = drag.current;
              drag.current = null;
              e.currentTarget.releasePointerCapture(e.pointerId);
              const updated = {
                ...current.drawing,
                widthPt: result.width / current.box.scale,
                heightPt: result.height / current.box.scale,
              };
              if (updated.position) {
                updated.position.horizontal.offsetPt =
                  (updated.position.horizontal.offsetPt ?? 0) +
                  (result.left - current.box.left) / current.box.scale;
                updated.position.vertical.offsetPt =
                  (updated.position.vertical.offsetPt ?? 0) +
                  (result.top - current.box.top) / current.box.scale;
              }
              await onChange(updated);
              setDraft(null);
            }}
            onPointerCancel={() => {
              drag.current = null;
              setDraft(null);
            }}
          />
        ))}
      <div className="object-mini-toolbar">
        {!d.equation && (
          <button
            type="button"
            aria-label="Move object"
            title="Drag to move object"
            style={{ cursor: "move", touchAction: "none" }}
            onPointerDown={(e) => {
              e.preventDefault();
              e.currentTarget.setPointerCapture(e.pointerId);
              drag.current = {
                x: e.clientX,
                y: e.clientY,
                handle: "move",
                box: bounds,
                drawing: structuredClone(d),
              };
            }}
            onPointerMove={(e) => {
              if (drag.current) setDraft(resized(e.clientX, e.clientY, false));
            }}
            onPointerUp={async (e) => {
              const current = drag.current;
              if (!current) return;
              drag.current = null;
              e.currentTarget.releasePointerCapture(e.pointerId);
              const dx = (e.clientX - current.x) / current.box.scale,
                dy = (e.clientY - current.y) / current.box.scale;
              await onChange({
                ...current.drawing,
                position: {
                  ...current.drawing.position,
                  horizontal: {
                    relativeTo:
                      current.drawing.position?.horizontal.relativeTo ??
                      "column",
                    ...current.drawing.position?.horizontal,
                    alignment: undefined,
                    offsetPt:
                      (current.drawing.position?.horizontal.offsetPt ?? 0) + dx,
                  },
                  vertical: {
                    relativeTo:
                      current.drawing.position?.vertical.relativeTo ??
                      "paragraph",
                    ...current.drawing.position?.vertical,
                    alignment: undefined,
                    offsetPt:
                      (current.drawing.position?.vertical.offsetPt ?? 0) + dy,
                  },
                  wrap: current.drawing.position?.wrap ?? "wrapNone",
                  behindText: current.drawing.position?.behindText ?? false,
                  zIndex: current.drawing.position?.zIndex ?? 2,
                },
              });
              setDraft(null);
            }}
            onPointerCancel={() => {
              drag.current = null;
              setDraft(null);
            }}
          >
            ✥
          </button>
        )}
        <button
          type="button"
          onClick={() =>
            onAction(d.equation ? "equation:focus" : "picture:options")
          }
        >
          {d.equation ? "Equation" : "Layout Options"}
        </button>
        <button
          type="button"
          onClick={() => onAction("object:delete")}
          aria-label={d.equation ? "Delete equation" : "Delete object"}
        >
          ×
        </button>
      </div>
      {d.equation && (
        <EquationEditor
          key={`${object.id}-${object.index}`}
          drawing={d}
          snippet={snippet}
          onApply={onChange}
          pendingDraft={pendingDraft}
          onDelete={() => onAction("object:delete")}
          onExit={() => container.current?.focus({ preventScroll: true })}
        />
      )}
    </div>,
    document.body,
  );
}
function EquationEditor({
  drawing,
  onApply,
  snippet,
  pendingDraft,
  onDelete,
  onExit,
}: {
  drawing: DrawingInline;
  onApply: (d: DrawingInline) => Promise<boolean>;
  snippet: { text: string; serial: number };
  pendingDraft: RefObject<(() => Promise<boolean>) | null>;
  onDelete: () => void;
  onExit: () => void;
}) {
  const [value, setValue] = useState(drawing.equation!.expression),
    [error, setError] = useState("");
  const input = useRef<HTMLTextAreaElement>(null),
    last = useRef(-1),
    latest = useRef(onApply),
    currentValue = useRef(value),
    saved = useRef(value),
    previous = useRef(value),
    saving = useRef<Promise<boolean> | null>(null),
    latestDrawing = useRef(drawing);
  latest.current = onApply;
  latestDrawing.current = drawing;
  currentValue.current = value;
  useEffect(() => {
    const next = drawing.equation!.expression;
    if (
      currentValue.current === previous.current ||
      currentValue.current === saved.current
    )
      setValue(next);
    previous.current = next;
  }, [drawing.equation?.expression]);
  async function save(expression: string): Promise<boolean> {
    if (saving.current) await saving.current;
    if (expression === latestDrawing.current.equation!.expression) return true;
    try {
      parseEquation(expression);
    } catch (error) {
      setError(error instanceof Error ? error.message : "Invalid equation.");
      return false;
    }
    saved.current = expression;
    const operation = latest.current({
      ...latestDrawing.current,
      equation: { expression },
    });
    saving.current = operation;
    try {
      const done = await operation;
      if (done) setError("");
      else setError("Could not update the equation. Choose Apply to retry.");
      return done;
    } finally {
      saving.current = null;
    }
  }
  useEffect(() => {
    const flush = () => save(currentValue.current);
    pendingDraft.current = flush;
    return () => {
      if (pendingDraft.current === flush) pendingDraft.current = null;
    };
  });
  useEffect(() => {
    if (value === drawing.equation!.expression) return;
    const timer = setTimeout(() => void save(value), 600);
    return () => clearTimeout(timer);
  }, [value, drawing.equation?.expression]);
  useEffect(() => {
    if (snippet.serial === last.current) return;
    last.current = snippet.serial;
    if (snippet.text === "__focus__") {
      input.current?.focus();
      if (value === "x") input.current?.select();
      return;
    }
    const node = input.current,
      start = node?.selectionStart ?? value.length,
      end = node?.selectionEnd ?? start;
    setValue(value.slice(0, start) + snippet.text + value.slice(end));
    requestAnimationFrame(() => {
      input.current?.focus();
      input.current?.setSelectionRange(
        start + snippet.text.length,
        start + snippet.text.length,
      );
    });
  }, [snippet]);
  return (
    <form
      className="inline-equation-editor"
      onSubmit={async (e) => {
        e.preventDefault();
        await save(value);
      }}
    >
      <label>
        Equation <span>LaTeX input</span>
        <textarea
          id="selected-equation-input"
          ref={input}
          value={value}
          aria-label="Equation source"
          onChange={(e) => setValue(e.target.value)}
          onBlur={(e) => {
            const target = e.relatedTarget as HTMLElement | null;
            if (target?.closest("button")) return;
            if (value !== drawing.equation!.expression) void save(value);
          }}
          onKeyDown={(e) => {
            if (e.key === "Escape") {
              e.preventDefault();
              e.stopPropagation();
              void save(value).then((done) => {
                if (done) onExit();
              });
            } else if (!value && ["Backspace", "Delete"].includes(e.key)) {
              e.preventDefault();
              onDelete();
            } else if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              e.currentTarget.form?.requestSubmit();
            }
          }}
        />
      </label>
      <div className="equation-preview">
        <EquationView expression={value} />
      </div>
      {error && <p role="alert">{error}</p>}
      <button className="primary-button">Apply</button>
    </form>
  );
}
