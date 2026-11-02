import { useEffect, useState } from "react";
import type {
  Paragraph,
  ParagraphFormat,
  RunFormat,
} from "@local/docx-ir-engine";
import { Group } from "./Controls";
import { Icon } from "./Icon";

export function HomeIcon({ name, size = 20 }: { name: string; size?: number }) {
  return (
    <img
      src={`/icons/fluent/${name}.svg`}
      alt=""
      aria-hidden="true"
      width={size}
      height={size}
      draggable={false}
    />
  );
}
function Button({
  label,
  icon,
  onClick,
  disabled,
  active,
  large = false,
}: {
  label: string;
  icon: string;
  onClick: () => void;
  disabled?: boolean;
  active?: boolean;
  large?: boolean;
}) {
  return (
    <button
      type="button"
      className={`home-tool ${large ? "home-tool-large" : ""} ${active ? "is-active" : ""}`}
      aria-label={label}
      title={label}
      aria-pressed={active}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      <HomeIcon name={icon} size={large ? 32 : 20} />
      {large && <span>{label}</span>}
    </button>
  );
}
function Split({
  label,
  icon,
  onClick,
  onArrow,
  disabled,
  large = false,
  color,
}: {
  label: string;
  icon: string;
  onClick: () => void;
  onArrow: () => void;
  disabled?: boolean;
  large?: boolean;
  color?: string;
}) {
  return (
    <div className={`home-split ${large ? "home-split-large" : ""}`}>
      <button
        type="button"
        className="home-tool"
        title={label}
        aria-label={label}
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={onClick}
      >
        <span className="home-icon-wrap">
          <HomeIcon name={icon} size={large ? 32 : 20} />
          {color && (
            <span className="home-color-line" style={{ background: color }} />
          )}
        </span>
        {large && <span>{label}</span>}
      </button>
      <button
        type="button"
        className="home-arrow"
        aria-label={`${label} options`}
        title={`${label} options`}
        aria-haspopup="dialog"
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={onArrow}
      >
        <Icon name="chevron" size={10} />
      </button>
    </div>
  );
}
function Combo({
  value,
  label,
  onCommit,
  onArrow,
  small = false,
  disabled,
}: {
  value: string;
  label: string;
  onCommit: (value: string) => void;
  onArrow: () => void;
  small?: boolean;
  disabled?: boolean;
}) {
  const [draft, setDraft] = useState(value),
    [focused, setFocused] = useState(false);
  useEffect(() => {
    if (!focused) setDraft(value);
  }, [value, focused]);
  return (
    <div className={`home-combo ${small ? "home-combo-size" : ""}`}>
      <input
        aria-label={label}
        value={draft}
        disabled={disabled}
        onFocus={() => setFocused(true)}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={() => {
          setFocused(false);
          if (draft.trim() && draft !== value) onCommit(draft.trim());
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            e.currentTarget.blur();
          }
          if (e.key === "Escape") {
            setDraft(value);
            e.currentTarget.blur();
          }
        }}
      />
      <button
        type="button"
        aria-label={`${label} options`}
        aria-haspopup="dialog"
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={onArrow}
      >
        <Icon name="chevron" size={10} />
      </button>
    </div>
  );
}
interface Props {
  selected?: Paragraph;
  format: RunFormat;
  disabled: boolean;
  canFormat: boolean;
  hasSelection: boolean;
  styles: { id: string; name: string }[];
  showMarks: boolean;
  painter: boolean;
  onRun: (format: RunFormat) => void;
  onParagraph: (format: ParagraphFormat) => void;
  onAction: (action: string) => void;
  onDialog: (dialog: string) => void;
}
export function HomeRibbon(p: Props) {
  const disabled = p.disabled || !p.canFormat,
    paragraphDisabled = p.disabled || !p.selected,
    format = p.format;
  const direction = p.selected?.properties.direction ?? "ltr";
  const indent =
    p.selected?.properties.indentStartPt ??
    p.selected?.properties.indentLeftPt ??
    0;
  return (
    <div
      className="ribbon home-ribbon"
      id="ribbon-Home"
      role="tabpanel"
      aria-labelledby="tab-Home"
    >
      <Group label="Clipboard" more={() => p.onDialog("clipboard")}>
        <Split
          label="Paste"
          icon="paste"
          large
          disabled={disabled}
          onClick={() => p.onAction("paste")}
          onArrow={() => p.onDialog("paste")}
        />
        <div className="home-clipboard-stack">
          <button
            disabled={disabled || !p.hasSelection}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => p.onAction("cut")}
          >
            <HomeIcon name="cut" size={16} />
            Cut
          </button>
          <button
            disabled={!p.hasSelection}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => p.onAction("copy")}
          >
            <HomeIcon name="copy" size={16} />
            Copy
          </button>
          <button
            className={p.painter ? "is-active" : ""}
            disabled={disabled}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => p.onAction("painter")}
          >
            <HomeIcon name="painter" size={16} />
            Format Painter
          </button>
        </div>
      </Group>
      <Group label="Font" more={() => p.onDialog("font")}>
        <div className="home-two-rows">
          <div className="home-row">
            <Combo
              label="Font family"
              value={format.fontFamily ?? "Calibri"}
              onCommit={(fontFamily) => p.onRun({ fontFamily })}
              onArrow={() => p.onDialog("fontFamily")}
              disabled={disabled}
            />
            <Combo
              label="Font size"
              value={String(format.fontSizePt ?? 12)}
              onCommit={(value) => p.onRun({ fontSizePt: Number(value) })}
              onArrow={() => p.onDialog("fontSize")}
              small
              disabled={disabled}
            />
            <button
              className="home-tool home-grow"
              aria-label="Grow Font"
              title="Grow Font"
              disabled={disabled}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() =>
                p.onRun({
                  fontSizePt: Math.min(400, (format.fontSizePt ?? 12) + 2),
                })
              }
            >
              A<sup>⌃</sup>
            </button>
            <button
              className="home-tool home-shrink"
              aria-label="Shrink Font"
              title="Shrink Font"
              disabled={disabled}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() =>
                p.onRun({
                  fontSizePt: Math.max(1, (format.fontSizePt ?? 12) - 2),
                })
              }
            >
              A<sup>⌄</sup>
            </button>
            <button
              className="home-tool home-case"
              aria-label="Change Case"
              title="Change Case"
              aria-haspopup="dialog"
              disabled={disabled}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => p.onDialog("case")}
            >
              Aa
              <Icon name="chevron" size={10} />
            </button>
            <Button
              label="Clear All Formatting"
              icon="clear"
              disabled={disabled}
              onClick={() => p.onAction("clear")}
            />
          </div>
          <div className="home-row">
            <Button
              label="Bold"
              icon="bold"
              active={!!format.bold}
              disabled={disabled}
              onClick={() => p.onRun({ bold: !format.bold })}
            />
            <Button
              label="Italic"
              icon="italic"
              active={!!format.italic}
              disabled={disabled}
              onClick={() => p.onRun({ italic: !format.italic })}
            />
            <Split
              label="Underline"
              icon="underline"
              disabled={disabled}
              onClick={() =>
                p.onRun({
                  underline:
                    format.underline && format.underline !== "none"
                      ? "none"
                      : "single",
                })
              }
              onArrow={() => p.onDialog("underline")}
            />
            <Button
              label="Strikethrough"
              icon="strike"
              active={!!format.strike}
              disabled={disabled}
              onClick={() => p.onRun({ strike: !format.strike })}
            />
            <Button
              label="Subscript"
              icon="subscript"
              active={format.verticalAlign === "subscript"}
              disabled={disabled}
              onClick={() =>
                p.onRun({
                  verticalAlign:
                    format.verticalAlign === "subscript"
                      ? "baseline"
                      : "subscript",
                })
              }
            />
            <Button
              label="Superscript"
              icon="superscript"
              active={format.verticalAlign === "superscript"}
              disabled={disabled}
              onClick={() =>
                p.onRun({
                  verticalAlign:
                    format.verticalAlign === "superscript"
                      ? "baseline"
                      : "superscript",
                })
              }
            />
            <span className="home-divider" />
            <Split
              label="Text Effects and Typography"
              icon="textEffects"
              disabled={disabled}
              onClick={() => p.onDialog("effects")}
              onArrow={() => p.onDialog("effects")}
            />
            <Split
              label="Text Highlight Color"
              icon="highlight"
              color="#fff200"
              disabled={disabled}
              onClick={() => p.onRun({ highlight: "yellow" })}
              onArrow={() => p.onDialog("highlight")}
            />
            <Split
              label="Font Color"
              icon="fontColor"
              color={
                format.color === "auto" ? "#000" : (format.color ?? "#c42b1c")
              }
              disabled={disabled}
              onClick={() => p.onRun({ color: format.color ?? "#c42b1c" })}
              onArrow={() => p.onDialog("color")}
            />
          </div>
        </div>
      </Group>
      <Group label="Paragraph" more={() => p.onAction("paragraph")}>
        <div className="home-two-rows">
          <div className="home-row">
            <Split
              label="Bullets"
              icon="bullets"
              disabled={paragraphDisabled}
              onClick={() => p.onAction("list:bullet")}
              onArrow={() => p.onDialog("bullets")}
            />
            <Split
              label="Numbering"
              icon="numbering"
              disabled={paragraphDisabled}
              onClick={() => p.onAction("list:decimal")}
              onArrow={() => p.onDialog("numbering")}
            />
            <Split
              label="Multilevel List"
              icon="multilevel"
              disabled={paragraphDisabled}
              onClick={() => p.onDialog("multilevel")}
              onArrow={() => p.onDialog("multilevel")}
            />
            <Button
              label="Decrease Indent"
              icon="indentDecrease"
              disabled={paragraphDisabled}
              onClick={() =>
                p.selected?.properties.numbering?.id &&
                p.selected.properties.numbering.id !== "0"
                  ? p.onAction("level:down")
                  : p.onParagraph({ indentStartPt: Math.max(0, indent - 18) })
              }
            />
            <Button
              label="Increase Indent"
              icon="indentIncrease"
              disabled={paragraphDisabled}
              onClick={() =>
                p.selected?.properties.numbering?.id &&
                p.selected.properties.numbering.id !== "0"
                  ? p.onAction("level:up")
                  : p.onParagraph({ indentStartPt: indent + 18 })
              }
            />
            <span className="home-divider" />
            <Button
              label="Left-to-Right Text Direction"
              icon="ltr"
              active={direction === "ltr"}
              disabled={paragraphDisabled}
              onClick={() => p.onParagraph({ direction: "ltr" })}
            />
            <Button
              label="Right-to-Left Text Direction"
              icon="rtl"
              active={direction === "rtl"}
              disabled={paragraphDisabled}
              onClick={() => p.onParagraph({ direction: "rtl" })}
            />
            <Button
              label="Sort"
              icon="sort"
              disabled={paragraphDisabled}
              onClick={() => p.onDialog("sort")}
            />
            <Button
              label="Show/Hide ¶"
              icon="marks"
              active={p.showMarks}
              onClick={() => p.onAction("marks")}
            />
          </div>
          <div className="home-row">
            {(["left", "center", "right"] as const).map((alignment) => (
              <Button
                key={alignment}
                label={`Align ${alignment}`}
                icon={`align${alignment[0].toUpperCase()}${alignment.slice(1)}`}
                active={p.selected?.properties.alignment === alignment}
                disabled={paragraphDisabled}
                onClick={() => p.onParagraph({ alignment })}
              />
            ))}
            <Split
              label="Justify"
              icon="justify"
              disabled={paragraphDisabled}
              onClick={() => p.onParagraph({ alignment: "justify" })}
              onArrow={() => p.onDialog("justify")}
            />
            <span className="home-divider" />
            <Split
              label="Line and Paragraph Spacing"
              icon="lineSpacing"
              disabled={paragraphDisabled}
              onClick={() => p.onDialog("spacing")}
              onArrow={() => p.onDialog("spacing")}
            />
            <Split
              label="Shading"
              icon="shading"
              disabled={paragraphDisabled}
              onClick={() => p.onParagraph({ shading: "#FFF2CC" })}
              onArrow={() => p.onDialog("shading")}
            />
            <Split
              label="Borders"
              icon="borders"
              disabled={paragraphDisabled}
              onClick={() => p.onParagraph({ borders: "bottom" })}
              onArrow={() => p.onDialog("borders")}
            />
          </div>
        </div>
      </Group>
      <Group label="Styles" more={() => p.onDialog("styles")}>
        <div className="home-style-gallery">
          {p.styles.slice(0, 6).map((style) => (
            <button
              key={style.id}
              className={
                p.selected?.properties.styleId === style.id ? "selected" : ""
              }
              disabled={paragraphDisabled}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => p.onParagraph({ styleId: style.id })}
            >
              <span
                style={{
                  fontSize:
                    style.id === "Title"
                      ? 23
                      : style.id.startsWith("Heading")
                        ? 19
                        : 18,
                  fontWeight: style.id.startsWith("Heading") ? 600 : 400,
                  color: style.id.startsWith("Heading") ? "#315a8c" : undefined,
                }}
              >
                AaBbCc
              </span>
              <small>{style.name}</small>
            </button>
          ))}
        </div>
        <button
          className="home-style-more"
          aria-label="More styles"
          aria-haspopup="dialog"
          onClick={() => p.onDialog("styles")}
        >
          <Icon name="chevron" size={12} />
        </button>
      </Group>
      <Group label="Editing">
        <div className="home-editing-stack">
          <button onClick={() => p.onAction("find")}>
            <HomeIcon name="find" size={17} />
            Find
          </button>
          <button onClick={() => p.onDialog("replace")}>
            <HomeIcon name="replace" size={17} />
            Replace
          </button>
          <button aria-haspopup="dialog" onClick={() => p.onDialog("select")}>
            <HomeIcon name="select" size={17} />
            Select
            <Icon name="chevron" size={10} />
          </button>
        </div>
      </Group>
    </div>
  );
}
