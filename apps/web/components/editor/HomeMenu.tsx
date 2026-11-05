import { useLayoutEffect, useRef, useState, type ReactNode } from "react";
import { createPortal } from "react-dom";
import { menuPosition, type MenuBounds } from "../../lib/menu-position";

export type MenuAnchor = MenuBounds & { element?: HTMLElement };
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
  placement = "below",
  className = "",
}: {
  label: string;
  anchor: MenuAnchor | null;
  onClose: () => void;
  children: ReactNode;
  placement?: "below" | "overlap";
  className?: string;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState({
    left: anchor?.left ?? 8,
    top: (anchor?.bottom ?? 0) + 2,
    maxHeight: 0,
    clipped: false,
  });
  const close = useRef(onClose);
  close.current = onClose;
  useLayoutEffect(() => {
    const el = ref.current!;
    if (!anchor) {
      close.current();
      return;
    }
    // The native top layer keeps the menu above the document without a modal backdrop.
    el.showPopover?.();
    const place = () => {
      const a = anchor.element?.getBoundingClientRect() ?? anchor;
      setPosition(
        menuPosition(
          a,
          { width: el.getBoundingClientRect().width, height: el.scrollHeight },
          { width: innerWidth, height: innerHeight },
          placement,
        ),
      );
    };
    place();
    const observer = new ResizeObserver(place);
    observer.observe(el);
    const outside = (event: PointerEvent) => {
      if (
        !el.contains(event.target as Node) &&
        !anchor.element?.contains(event.target as Node)
      )
        close.current();
    };
    window.addEventListener("pointerdown", outside, true);
    window.addEventListener("resize", place);
    const scroll = (event: Event) => {
      if (!(event.target instanceof Node) || !el.contains(event.target))
        place();
    };
    window.addEventListener("scroll", scroll, true);
    // Focus belongs to the menu while it is open; the model selection is retained separately.
    el.querySelector<HTMLElement>(
      "input:not(:disabled),button:not(:disabled)",
    )?.focus({ preventScroll: true });
    return () => {
      observer.disconnect();
      window.removeEventListener("pointerdown", outside, true);
      window.removeEventListener("resize", place);
      window.removeEventListener("scroll", scroll, true);
      if (
        el.isConnected &&
        typeof el.hidePopover === "function" &&
        el.matches(":popover-open")
      )
        el.hidePopover();
    };
  }, [anchor, placement]);
  return createPortal(
    <div
      ref={ref}
      popover="manual"
      role="menu"
      aria-label={label}
      data-home-menu-root
      className={`home-dropdown ${className}`}
      style={{
        position: "fixed",
        inset: "auto",
        margin: 0,
        padding: 0,
        width: "max-content",
        maxWidth: "min(560px, calc(100vw - 16px))",
        zIndex: 1000,
        left: position.left,
        top: position.top,
        maxHeight: position.maxHeight || undefined,
        overflowY: position.clipped ? "auto" : "visible",
        overflowX: position.clipped ? "hidden" : "visible",
      }}
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
