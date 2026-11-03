import { useLayoutEffect, useRef, useState, type ReactNode } from "react";
import { createPortal } from "react-dom";

export type MenuAnchor = {
  left: number;
  top: number;
  bottom: number;
  right: number;
};
export function MenuBranch({
  label,
  children,
  disabled = false,
}: {
  label: string;
  children: ReactNode;
  disabled?: boolean;
}) {
  const [open, setOpen] = useState(false),
    [left, setLeft] = useState(false),
    [top, setTop] = useState(-4);
  const ref = useRef<HTMLDivElement>(null),
    trigger = useRef<HTMLButtonElement>(null);
  useLayoutEffect(() => {
    if (open && ref.current) {
      const branch = ref.current.getBoundingClientRect();
      setLeft(branch.right + 285 > innerWidth);
      const content = ref.current.querySelector<HTMLElement>(
        ":scope > .home-submenu-content",
      );
      if (content)
        setTop(
          Math.max(
            8 - branch.top,
            -4 -
              Math.max(
                0,
                branch.top - 4 + content.offsetHeight - innerHeight + 8,
              ),
          ),
        );
    }
  }, [open]);
  return (
    <div
      ref={ref}
      className={`home-submenu ${left ? "opens-left" : ""}`}
      onMouseEnter={() => !disabled && setOpen(true)}
      onMouseLeave={() => setOpen(false)}
    >
      <button
        ref={trigger}
        className="home-choice"
        role="menuitem"
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={disabled}
        onClick={() => setOpen(true)}
        onKeyDown={(e) => {
          if (e.key === "ArrowRight") {
            e.preventDefault();
            e.stopPropagation();
            setOpen(true);
            requestAnimationFrame(() =>
              ref.current
                ?.querySelector<HTMLElement>(".home-submenu-content button")
                ?.focus(),
            );
          }
        }}
      >
        {label}
        <span className="submenu-chevron">›</span>
      </button>
      {open && (
        <div
          role="menu"
          aria-label={label}
          className="home-submenu-content"
          style={{ top }}
          onKeyDown={(e) => {
            if (e.key === "ArrowLeft" || e.key === "Escape") {
              e.preventDefault();
              e.stopPropagation();
              setOpen(false);
              trigger.current?.focus();
            }
          }}
        >
          {children}
        </div>
      )}
    </div>
  );
}
/** One anchored surface for ribbon galleries and right-click menus. */
export function HomeMenu({
  label,
  anchor,
  onClose,
  children,
}: {
  label: string;
  anchor: MenuAnchor | null;
  onClose: () => void;
  children: ReactNode;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState({ left: 8, top: 190 });
  const close = useRef(onClose);
  close.current = onClose;
  useLayoutEffect(() => {
    const el = ref.current!;
    const place = () => {
      const rect = el.getBoundingClientRect(),
        a = anchor ?? { left: 8, right: 8, top: 190, bottom: 190 };
      const left = Math.max(8, Math.min(a.left, innerWidth - rect.width - 8));
      const top =
        a.bottom + rect.height < innerHeight - 8
          ? a.bottom + 2
          : Math.max(
              8,
              Math.min(a.top - rect.height - 2, innerHeight - rect.height - 8),
            );
      setPosition({ left, top });
    };
    place();
    const observer = new ResizeObserver(place);
    observer.observe(el);
    const outside = (event: PointerEvent) => {
      if (!el.contains(event.target as Node)) close.current();
    };
    window.addEventListener("pointerdown", outside, true);
    window.addEventListener("resize", place);
    // Focus belongs to the menu while it is open; the model selection is retained separately.
    el.querySelector<HTMLElement>(
      "input:not(:disabled),button:not(:disabled)",
    )?.focus();
    return () => {
      observer.disconnect();
      window.removeEventListener("pointerdown", outside, true);
      window.removeEventListener("resize", place);
    };
  }, [anchor]);
  return createPortal(
    <div
      ref={ref}
      role="menu"
      aria-label={label}
      data-home-menu-root
      className="home-dropdown"
      style={position}
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          event.stopPropagation();
          onClose();
          return;
        }
        if (event.key === "Tab") {
          onClose();
          return;
        }
        if ((event.target as Element).matches("input,textarea,select")) return;
        const buttons = Array.from(
          ref.current!.querySelectorAll<HTMLButtonElement>(
            "button:not(:disabled)",
          ),
        ).filter((el) => el.getClientRects().length > 0);
        const index = buttons.indexOf(
          window.document.activeElement as HTMLButtonElement,
        );
        let next = index;
        if (event.key === "ArrowDown") next = (index + 1) % buttons.length;
        else if (event.key === "ArrowUp")
          next = (index - 1 + buttons.length) % buttons.length;
        else if (event.key === "Home") next = 0;
        else if (event.key === "End") next = buttons.length - 1;
        else return;
        event.preventDefault();
        buttons[next]?.focus();
      }}
    >
      {children}
    </div>,
    window.document.body,
  );
}
