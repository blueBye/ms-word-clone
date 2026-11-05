import { useRef, useState, type ReactNode } from "react";
import { HomeMenu, type MenuAnchor } from "./HomeMenu";
import { Icon } from "./Icon";
import { RibbonIcon } from "./RibbonIcon";

export function RibbonCommand({
  label,
  icon,
  large = false,
  disabled = false,
  onClick,
  children,
  menu = false,
}: {
  label: string;
  icon: string;
  large?: boolean;
  disabled?: boolean;
  onClick?: () => void;
  children?: ReactNode;
  menu?: boolean;
}) {
  return (
    <button
      type="button"
      className={`office-command ${large ? "office-command-large" : ""}`}
      disabled={disabled}
      title={disabled ? `${label} — not available for editing yet` : label}
      aria-label={label}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      <RibbonIcon name={icon} size={large ? 32 : 16} />
      <span>{label}</span>
      {menu && <Icon name="chevron" size={10} />}
      {children}
    </button>
  );
}
export function RibbonDropdown({
  label,
  icon,
  large = false,
  disabled = false,
  children,
}: {
  label: string;
  icon: string;
  large?: boolean;
  disabled?: boolean;
  children: ReactNode;
}) {
  const button = useRef<HTMLButtonElement>(null);
  const [anchor, setAnchor] = useState<MenuAnchor | null>(null);
  const dismiss = () => {
    setAnchor(null);
    button.current?.focus({ preventScroll: true });
  };
  return (
    <>
      <button
        ref={button}
        type="button"
        className={`office-command office-dropdown ${large ? "office-command-large" : ""}`}
        aria-label={label}
        title={label}
        aria-haspopup="menu"
        aria-expanded={!!anchor}
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => {
          const element = button.current!;
          setAnchor(
            anchor
              ? null
              : { ...element.getBoundingClientRect().toJSON(), element },
          );
        }}
      >
        <RibbonIcon name={icon} size={large ? 32 : 16} />
        <span>{label}</span>
        <Icon name="chevron" size={10} />
      </button>
      {anchor && (
        <HomeMenu label={label} anchor={anchor} onClose={dismiss}>
          <div
            className="office-menu"
            onClick={(e) => {
              if ((e.target as Element).closest("button:not(:disabled)"))
                dismiss();
            }}
          >
            {children}
          </div>
        </HomeMenu>
      )}
    </>
  );
}
export function RibbonChoice({
  label,
  description,
  icon,
  disabled = false,
  checked,
  onClick,
}: {
  label: string;
  description?: string;
  icon?: string;
  disabled?: boolean;
  checked?: boolean;
  onClick?: () => void;
}) {
  return (
    <button
      type="button"
      role={checked === undefined ? "menuitem" : "menuitemradio"}
      aria-checked={checked}
      className={`office-choice ${checked ? "is-active" : ""}`}
      disabled={disabled}
      title={disabled ? "Not available for editing yet" : undefined}
      onClick={onClick}
    >
      {icon && <RibbonIcon name={icon} size={28} />}
      <span>
        {label}
        {description && <small>{description}</small>}
      </span>
      {checked && <Icon name="check" size={14} />}
    </button>
  );
}

export function TablePicker({
  disabled,
  onPick,
}: {
  disabled: boolean;
  onPick: (rows: number, columns: number) => void;
}) {
  const [size, setSize] = useState({ rows: 1, columns: 1 });
  return (
    <div className="table-picker">
      <p>
        {size.columns} × {size.rows} Table
      </p>
      <div
        className="table-picker-grid"
        role="group"
        aria-label="Choose table size"
      >
        {Array.from({ length: 80 }, (_, i) => {
          const rows = Math.floor(i / 10) + 1,
            columns = (i % 10) + 1;
          return (
            <button
              type="button"
              key={i}
              aria-label={`${columns} columns, ${rows} rows`}
              disabled={disabled}
              tabIndex={rows === size.rows && columns === size.columns ? 0 : -1}
              className={
                rows <= size.rows && columns <= size.columns ? "selected" : ""
              }
              onMouseEnter={() => setSize({ rows, columns })}
              onFocus={() => setSize({ rows, columns })}
              onClick={() => onPick(rows, columns)}
              onKeyDown={(e) => {
                const delta = {
                  ArrowRight: 1,
                  ArrowLeft: -1,
                  ArrowDown: 10,
                  ArrowUp: -10,
                }[e.key];
                if (delta === undefined) return;
                e.preventDefault();
                e.stopPropagation();
                const next = e.currentTarget.parentElement?.children[
                  Math.max(0, Math.min(79, i + delta))
                ] as HTMLButtonElement | undefined;
                next?.focus();
              }}
            />
          );
        })}
      </div>
    </div>
  );
}
