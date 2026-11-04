import { useEffect, useState, useRef } from "react";
import type {
  Paragraph,
  ParagraphFormat,
  RunFormat,
  DocxStyle,
} from "@local/docx-ir-engine";
import { Group } from "./Controls";
import { Icon } from "./Icon";
import { StylePreview } from "./StylePreview";
import type { MenuAnchor } from "./HomeMenu";

export function HomeIcon({ name, size = 20 }: { name: string; size?: number }) {
  const glyphs: Record<string, React.ReactNode> = {
    bold: (
      <text x="6" y="18" fontFamily="Arial" fontSize="18" fontWeight="700">
        B
      </text>
    ),
    italic: (
      <text
        x="8"
        y="18"
        fontFamily="Times New Roman"
        fontSize="20"
        fontStyle="italic"
      >
        I
      </text>
    ),
    underline: (
      <>
        <text x="6" y="16" fontFamily="Arial" fontSize="17">
          U
        </text>
        <path d="M5 20h14" />
      </>
    ),
    strike: (
      <>
        <text x="2" y="17" fontFamily="Arial" fontSize="14">
          abc
        </text>
        <path d="M2 12h21" />
      </>
    ),
    subscript: (
      <>
        <text x="3" y="15" fontFamily="Arial" fontSize="18">
          x
        </text>
        <text x="14" y="21" fontFamily="Arial" fontSize="10" fill="#2261a0">
          2
        </text>
      </>
    ),
    superscript: (
      <>
        <text x="3" y="20" fontFamily="Arial" fontSize="18">
          x
        </text>
        <text x="14" y="10" fontFamily="Arial" fontSize="10" fill="#2261a0">
          2
        </text>
      </>
    ),
    fontColor: (
      <>
        <text x="5" y="17" fontFamily="Times New Roman" fontSize="21">
          A
        </text>
        <path d="M3 22h18" stroke="#d32424" strokeWidth="3" />
      </>
    ),
    textEffects: (
      <text
        x="4"
        y="19"
        fontFamily="Arial"
        fontSize="21"
        fill="#c9e3f3"
        stroke="#2379a8"
        strokeWidth=".8"
      >
        A
      </text>
    ),
    highlight: (
      <>
        <path d="m9 16 9-12 4 3-9 12z" fill="#d2d5d8" />
        <path d="m6 16 3 3-4 1z" fill="#333" />
        <path d="M2 22h20" stroke="#fff200" strokeWidth="3" />
      </>
    ),
    clear: (
      <>
        <text x="2" y="17" fontFamily="Arial" fontSize="19" fill="#647485">
          A
        </text>
        <path d="m12 16 5-6 6 5-5 6z" fill="#c987ca" stroke="#996199" />
        <path d="M12 21h10" />
      </>
    ),
    ltr: (
      <>
        <text x="11" y="19" fontFamily="Arial" fontSize="20">
          ¶
        </text>
        <path d="M1 10h8M6 7l3 3-3 3" stroke="#3486a3" fill="none" />
      </>
    ),
    rtl: (
      <>
        <text x="11" y="19" fontFamily="Arial" fontSize="20">
          ¶
        </text>
        <path d="M9 10H1m3-3-3 3 3 3" stroke="#3486a3" fill="none" />
      </>
    ),
    multilevel: (
      <>
        <g fontFamily="Arial" fontSize="7" fill="#347da1" stroke="none">
          <text x="1" y="7">
            1
          </text>
          <text x="5" y="14">
            a
          </text>
          <text x="9" y="21">
            i
          </text>
        </g>
        <path d="M8 5h12M12 12h8M16 19h4" />
      </>
    ),
    painter: (
      <>
        <path d="M4 3h16v6H4z" fill="#f0b46b" stroke="#b57a3f" />
        <path d="M6 9h12l-4 4v8h-4v-8z" fill="#dae1e5" stroke="#788c99" />
      </>
    ),
  };
  if (glyphs[name])
    return (
      <svg
        width={size}
        height={size}
        viewBox="0 0 24 24"
        aria-hidden="true"
        fill="currentColor"
        stroke="none"
      >
        <g strokeWidth="1.2">{glyphs[name]}</g>
      </svg>
    );
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
  active,
}: {
  label: string;
  icon: string;
  onClick: () => void;
  onArrow: () => void;
  disabled?: boolean;
  large?: boolean;
  color?: string;
  active?: boolean;
}) {
  return (
    <div className={`home-split ${large ? "home-split-large" : ""}`}>
      <button
        type="button"
        className={`home-tool ${active ? "is-active" : ""}`}
        aria-pressed={active}
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
        aria-haspopup="menu"
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
        aria-haspopup="menu"
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
  repeat: RunFormat;
  disabled: boolean;
  canFormat: boolean;
  hasSelection: boolean;
  styles: DocxStyle[];
  showMarks: boolean;
  painter: boolean;
  onRun: (format: RunFormat) => void;
  onParagraph: (format: ParagraphFormat) => void;
  onAction: (action: string) => void;
  onDialog: (dialog: string, anchor?: MenuAnchor) => void;
  onStyleContext: (id: string, anchor: MenuAnchor) => void;
}
export function HomeRibbon(p: Props) {
  const trigger = useRef<HTMLElement | null>(null);
  const open = (name: string) => {
    const element = trigger.current;
    const rect = element?.getBoundingClientRect();
    p.onDialog(
      name,
      rect
        ? {
            left: rect.left,
            right: rect.right,
            top: rect.top,
            bottom: rect.bottom,
            element: element!,
          }
        : undefined,
    );
  };
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
      onClickCapture={(e) => {
        trigger.current = (e.target as Element).closest<HTMLElement>("button");
      }}
      id="ribbon-Home"
      role="tabpanel"
      aria-labelledby="tab-Home"
    >
      <Group label="Clipboard" more={() => open("clipboard")}>
        <Split
          label="Paste"
          icon="paste"
          large
          disabled={disabled}
          onClick={() => p.onAction("paste")}
          onArrow={() => open("paste")}
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
      <Group label="Font" more={() => open("font")}>
        <div className="home-two-rows">
          <div className="home-row">
            <Combo
              label="Font family"
              value={format.fontFamily ?? "Calibri"}
              onCommit={(fontFamily) => p.onRun({ fontFamily })}
              onArrow={() => open("fontFamily")}
              disabled={disabled}
            />
            <Combo
              label="Font size"
              value={String(format.fontSizePt ?? 12)}
              onCommit={(value) => p.onRun({ fontSizePt: Number(value) })}
              onArrow={() => open("fontSize")}
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
              aria-haspopup="menu"
              disabled={disabled}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => open("case")}
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
              active={!!format.underline && format.underline !== "none"}
              icon="underline"
              disabled={disabled}
              onClick={() =>
                p.onRun({
                  underline:
                    format.underline && format.underline !== "none"
                      ? "none"
                      : (p.repeat.underline ?? "single"),
                })
              }
              onArrow={() => open("underline")}
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
              onClick={() => open("effects")}
              onArrow={() => open("effects")}
            />
            <Split
              label="Text Highlight Color"
              icon="highlight"
              color="#fff200"
              disabled={disabled}
              onClick={() =>
                p.onRun({ highlight: p.repeat.highlight ?? "yellow" })
              }
              onArrow={() => open("highlight")}
            />
            <Split
              label="Font Color"
              icon="fontColor"
              color={
                format.color === "auto" ? "#000" : (p.repeat.color ?? "#c42b1c")
              }
              disabled={disabled}
              onClick={() => p.onRun({ color: p.repeat.color ?? "#c42b1c" })}
              onArrow={() => open("color")}
            />
          </div>
        </div>
      </Group>
      <Group label="Paragraph" more={() => p.onAction("paragraph")}>
        <div className="home-two-rows">
          <div className="home-row">
            <Split
              label="Bullets"
              active={p.selected?.properties.list?.format === "bullet"}
              icon="bullets"
              disabled={paragraphDisabled}
              onClick={() => p.onAction("list:bullet")}
              onArrow={() => open("bullets")}
            />
            <Split
              label="Numbering"
              active={
                !!p.selected?.properties.list &&
                p.selected.properties.list.format !== "bullet"
              }
              icon="numbering"
              disabled={paragraphDisabled}
              onClick={() => p.onAction("list:decimal")}
              onArrow={() => open("numbering")}
            />
            <Split
              label="Multilevel List"
              icon="multilevel"
              disabled={paragraphDisabled}
              onClick={() => open("multilevel")}
              onArrow={() => open("multilevel")}
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
              onClick={() => open("sort")}
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
              active={p.selected?.properties.alignment === "justify"}
              icon="justify"
              disabled={paragraphDisabled}
              onClick={() => p.onParagraph({ alignment: "justify" })}
              onArrow={() => open("justify")}
            />
            <span className="home-divider" />
            <Split
              label="Line and Paragraph Spacing"
              icon="lineSpacing"
              disabled={paragraphDisabled}
              onClick={() => open("spacing")}
              onArrow={() => open("spacing")}
            />
            <Split
              label="Shading"
              icon="shading"
              disabled={paragraphDisabled}
              onClick={() => p.onParagraph({ shading: "#FFF2CC" })}
              onArrow={() => open("shading")}
            />
            <Split
              label="Borders"
              icon="borders"
              disabled={paragraphDisabled}
              onClick={() => p.onParagraph({ borders: "bottom" })}
              onArrow={() => open("borders")}
            />
          </div>
        </div>
      </Group>
      <Group label="Styles" more={() => open("styles")}>
        <div className="home-style-gallery">
          {p.styles
            .filter((s) => s.gallery)
            .slice(0, 12)
            .map((style) => (
              <button
                key={style.id}
                className={
                  p.selected?.properties.styleId === style.id ? "selected" : ""
                }
                disabled={paragraphDisabled}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => p.onParagraph({ styleId: style.id })}
                onContextMenu={(e) => {
                  e.preventDefault();
                  const rect = e.currentTarget.getBoundingClientRect();
                  p.onStyleContext(
                    style.id,
                    e.clientX
                      ? {
                          left: e.clientX,
                          right: e.clientX,
                          top: e.clientY,
                          bottom: e.clientY,
                        }
                      : rect,
                  );
                }}
              >
                <StylePreview style={style} compact />
              </button>
            ))}
        </div>
        <button
          className="home-style-more"
          aria-label="More styles"
          aria-haspopup="menu"
          onClick={() => open("styles")}
        >
          <Icon name="chevron" size={12} />
        </button>
      </Group>
      <Group label="Editing">
        <div className="home-editing-stack">
          <div className="home-find-split">
            <button onClick={() => p.onAction("find")}>
              <HomeIcon name="find" size={17} />
              Find
            </button>
            <button
              aria-label="Find options"
              aria-haspopup="menu"
              onClick={() => open("findMenu")}
            >
              <Icon name="chevron" size={10} />
            </button>
          </div>
          <button onClick={() => open("replace")}>
            <HomeIcon name="replace" size={17} />
            Replace
          </button>
          <button aria-haspopup="menu" onClick={() => open("select")}>
            <HomeIcon name="select" size={17} />
            Select
            <Icon name="chevron" size={10} />
          </button>
        </div>
      </Group>
    </div>
  );
}
