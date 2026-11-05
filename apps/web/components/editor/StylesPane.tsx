import type { DocumentStyle } from "@local/docx-ir-engine/core";
import type { MenuAnchor } from "./HomeMenu";
import { StylePreview } from "./StylePreview";

export function StylesPane({
  styles,
  current,
  disabled,
  onApply,
  onContext,
}: {
  styles: DocumentStyle[];
  current?: string;
  disabled: boolean;
  onApply: (id: string) => void;
  onContext: (id: string, anchor: MenuAnchor) => void;
}) {
  return (
    <div className="styles-pane-list">
      {styles.map((style) => (
        <button
          key={style.id}
          disabled={disabled}
          className={current === style.id ? "is-active" : ""}
          aria-pressed={current === style.id}
          title={`Apply ${style.name}`}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => onApply(style.id)}
          onContextMenu={(e) => {
            e.preventDefault();
            onContext(style.id, {
              left: e.clientX,
              right: e.clientX,
              top: e.clientY,
              bottom: e.clientY,
            });
          }}
          onKeyDown={(e) => {
            if (e.key === "ContextMenu" || (e.shiftKey && e.key === "F10")) {
              e.preventDefault();
              const element = e.currentTarget;
              onContext(style.id, {
                ...element.getBoundingClientRect().toJSON(),
                element,
              });
            }
          }}
        >
          <StylePreview style={style} compact />
        </button>
      ))}
    </div>
  );
}
