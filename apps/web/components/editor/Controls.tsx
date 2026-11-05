import { useEffect, useId, useRef, type ReactNode } from "react";
import { Icon } from "./Icon";

export function Tool({
  label,
  icon,
  children,
  onClick,
  disabled,
  reason,
  active,
  large = false,
}: {
  label: string;
  icon?: string;
  children?: ReactNode;
  onClick?: () => void;
  disabled?: boolean;
  reason?: string;
  active?: boolean | "mixed";
  large?: boolean;
}) {
  return (
    <button
      type="button"
      className={`tool ${large ? "tool-large" : ""} ${active ? "is-active" : ""}`}
      title={
        disabled
          ? `${label}: ${reason ?? "Select an editable paragraph first."}`
          : label
      }
      aria-label={label}
      aria-pressed={active}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {icon ? <Icon name={icon} size={large ? 27 : 18} /> : children}
      <span className={large || (!children && !icon) ? "" : "tool-label"}>
        {label}
      </span>
    </button>
  );
}
export function Group({
  label,
  children,
  more,
  moreLabel,
  morePopup = "dialog",
}: {
  label: string;
  children: ReactNode;
  more?: () => void;
  moreLabel?: string;
  morePopup?: "dialog" | "menu" | false;
}) {
  return (
    <section className="ribbon-group" aria-label={label}>
      <div className="ribbon-group-tools">{children}</div>
      <div className="group-caption">
        {label}
        {more && (
          <button
            type="button"
            className="ribbon-dialog-launcher"
            title={moreLabel ?? `More ${label} options`}
            aria-label={moreLabel ?? `More ${label} options`}
            aria-haspopup={morePopup || undefined}
            onMouseDown={(e) => e.preventDefault()}
            onClick={more}
          >
            <svg
              width="12"
              height="12"
              viewBox="0 0 12 12"
              aria-hidden="true"
              fill="none"
              stroke="currentColor"
              strokeWidth="1"
            >
              <path d="M2 8V2h6M5 5l5 5M6 10h4V6" />
            </svg>
          </button>
        )}
      </div>
    </section>
  );
}
export function Menu({
  label,
  icon,
  children,
  large = false,
}: {
  label: string;
  icon?: string;
  children: ReactNode;
  large?: boolean;
}) {
  return (
    <details
      className={`tool-menu ${large ? "large-menu" : ""}`}
      onToggle={(event) => {
        if (event.currentTarget.open)
          window.document
            .querySelectorAll<HTMLDetailsElement>(".tool-menu[open]")
            .forEach((menu) => {
              if (menu !== event.currentTarget) menu.open = false;
            });
      }}
    >
      <summary title={label}>
        {icon && <Icon name={icon} size={large ? 27 : 18} />}
        <span>{label}</span>
        <Icon name="chevron" size={12} />
      </summary>
      <div
        className="tool-menu-content"
        onClick={(e) => {
          if ((e.target as Element).closest("button:not(:disabled)"))
            e.currentTarget.closest("details")?.removeAttribute("open");
        }}
      >
        {children}
      </div>
    </details>
  );
}
export function MenuItem({
  children,
  onClick,
  disabled,
  reason,
}: {
  children: ReactNode;
  onClick?: () => void;
  disabled?: boolean;
  reason?: string;
}) {
  return (
    <button
      type="button"
      className="menu-item"
      disabled={disabled}
      title={reason}
      onClick={onClick}
    >
      {children}
      {disabled && <small>{reason ?? "Not available yet"}</small>}
    </button>
  );
}
export function Dialog({
  title,
  children,
  onClose,
}: {
  title: string;
  children: ReactNode;
  onClose: () => void;
}) {
  const ref = useRef<HTMLDialogElement>(null);
  const titleId = useId();
  useEffect(() => {
    const el = ref.current;
    el?.showModal();
    return () => el?.close();
  }, []);
  return (
    <dialog
      ref={ref}
      className="editor-dialog"
      aria-labelledby={titleId}
      onCancel={onClose}
    >
      <div className="dialog-title">
        <h2 id={titleId}>{title}</h2>
        <button aria-label="Close dialog" onClick={onClose}>
          <Icon name="close" />
        </button>
      </div>
      {children}
    </dialog>
  );
}
