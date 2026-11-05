import type {
  DocumentIR,
  Paragraph,
  ParagraphFormat,
  RunFormat,
} from "@local/docx-ir-engine/core";
import { Group } from "./Controls";
import {
  RibbonCommand as Command,
  RibbonDropdown as Dropdown,
  RibbonChoice as Choice,
  TablePicker,
} from "./RibbonControls";
import { RibbonIcon } from "./RibbonIcon";

export type RibbonTab = "Home" | "Insert" | "Layout";
export interface RibbonProps {
  tab: RibbonTab;
  document: DocumentIR | null;
  selected?: Paragraph;
  range?: { start: number; end: number } | null;
  disabled: boolean;
  canFormat: boolean;
  canPage: boolean;
  families: string[];
  onRun: (format: RunFormat) => void;
  onParagraph: (format: ParagraphFormat) => void;
  onAction: (name: string) => void;
}
export function Ribbon({
  tab,
  document,
  selected,
  disabled,
  canFormat,
  canPage,
  onParagraph,
  onAction,
}: RibbonProps) {
  const paragraphDisabled = disabled || !selected;
  const pageDisabled = disabled || !canPage;
  const command = (
    label: string,
    icon: string,
    large = true,
    action?: string,
  ) => (
    <Command
      key={label}
      label={label}
      icon={icon}
      large={large}
      menu={[
        "Cover Page",
        "Pictures",
        "Shapes",
        "3D Models",
        "Screenshot",
        "Page Number",
        "Text Box",
        "Quick Parts",
        "WordArt",
        "Drop Cap",
        "Signature Line",
        "Object",
        "Equation",
        "Columns",
        "Position",
        "Wrap Text",
        "Bring Forward",
        "Send Backward",
        "Align",
        "Group",
        "Rotate",
        "Line Numbers",
        "Hyphenation",
      ].includes(label)}
      disabled={!action || disabled}
      onClick={() => action && onAction(action)}
    />
  );
  const properties = selected?.properties;
  const rtl = properties?.direction === "rtl";
  const indents = [
    {
      label: "Before text",
      key: "indentStartPt",
      value:
        properties?.indentStartPt ??
        (rtl ? properties?.indentRightPt : properties?.indentLeftPt) ??
        0,
    },
    {
      label: "After text",
      key: "indentEndPt",
      value:
        properties?.indentEndPt ??
        (rtl ? properties?.indentLeftPt : properties?.indentRightPt) ??
        0,
    },
  ] as const;
  const spacings = [
    {
      label: "Before",
      key: "spaceBeforePt",
      value: properties?.spaceBeforePt ?? 0,
    },
    {
      label: "After",
      key: "spaceAfterPt",
      value: properties?.spaceAfterPt ?? 0,
    },
  ] as const;
  return (
    <div
      className="ribbon office-ribbon"
      id={`ribbon-${tab}`}
      role="tabpanel"
      aria-labelledby={`tab-${tab}`}
    >
      {tab === "Insert" && (
        <>
          <Group label="Pages">
            <div className="office-stack">
              {command("Cover Page", "cover", false)}
              <Command
                label="Blank Page"
                icon="blank"
                disabled={disabled || !document}
                onClick={() => onAction("blankPage")}
              />
              <Command
                label="Page Break"
                icon="pageBreak"
                disabled={paragraphDisabled}
                onClick={() => onParagraph({ pageBreakBefore: true })}
              />
            </div>
          </Group>
          <Group label="Tables">
            <Dropdown
              label="Table"
              icon="table"
              large
              disabled={disabled || !document}
            >
              <TablePicker
                disabled={disabled || !document}
                onPick={(rows, columns) => onAction(`table:${rows},${columns}`)}
              />
              <Choice
                label="Insert Table…"
                icon="table"
                onClick={() => onAction("table")}
              />
              <Choice label="Draw Table" disabled />
              <Choice label="Convert Text to Table…" disabled />
              <Choice label="Excel Spreadsheet" disabled />
              <Choice label="Quick Tables" disabled />
            </Dropdown>
          </Group>
          <Group label="Illustrations">
            {command("Pictures", "picture")}
            {command("Shapes", "shapes")}
            {command("Icons", "icons")}
            {command("3D Models", "model")}
            {command("SmartArt", "smartArt")}
            {command("Chart", "chart")}
            {command("Screenshot", "screenshot")}
          </Group>
          <Group label="Media">{command("Online Videos", "video")}</Group>
          <Group label="Links">
            <div className="office-stack">
              {command("Link", "link", false)}
              {command("Bookmark", "bookmark", false)}
              {command("Cross-reference", "reference", false)}
            </div>
          </Group>
          <Group label="Comments">{command("Comment", "comment")}</Group>
          <Group label="Header & Footer">
            {(["Header", "Footer"] as const).map((label) => (
              <Dropdown
                key={label}
                label={label}
                icon={label.toLowerCase()}
                large
                disabled={!document}
              >
                <Choice
                  label={`View ${label}`}
                  icon={label.toLowerCase()}
                  onClick={() => onAction("stories")}
                />
                <Choice label={`Edit ${label}`} disabled />
                <Choice label={`Remove ${label}`} disabled />
              </Dropdown>
            ))}
            {command("Page Number", "pageNumber")}
          </Group>
          <Group label="Text">
            {command("Text Box", "textBox")}
            {command("Quick Parts", "quickParts")}
            {command("WordArt", "wordArt")}
            {command("Drop Cap", "dropCap")}
            <div className="office-stack">
              {command("Signature Line", "signature", false)}
              <Command
                label="Date & Time"
                icon="date"
                disabled={disabled || !canFormat}
                onClick={() => onAction("date")}
              />
              {command("Object", "object", false)}
            </div>
          </Group>
          <Group label="Symbols">
            <div className="office-stack">
              {command("Equation", "equation", false)}
              <Dropdown
                label="Symbol"
                icon="symbol"
                disabled={disabled || !canFormat}
              >
                <div className="office-symbol-grid">
                  {[
                    "☐",
                    "☑",
                    "✓",
                    "©",
                    "®",
                    "™",
                    "±",
                    "×",
                    "÷",
                    "α",
                    "β",
                    "Ω",
                    "→",
                    "←",
                    "…",
                    "—",
                  ].map((symbol) => (
                    <button
                      type="button"
                      key={symbol}
                      title={symbol}
                      onClick={() => onAction(`symbol:${symbol}`)}
                    >
                      {symbol}
                    </button>
                  ))}
                </div>
                <Choice
                  label="Zero-width non-joiner"
                  description="ZWNJ"
                  onClick={() => onAction("symbol:\u200c")}
                />
                <Choice
                  label="Nonbreaking Space"
                  onClick={() => onAction("symbol:\u00a0")}
                />
              </Dropdown>
            </div>
          </Group>
        </>
      )}
      {tab === "Layout" && (
        <>
          <Group label="Page Setup" more={() => onAction("page")}>
            <Dropdown
              label="Margins"
              icon="margins"
              large
              disabled={pageDisabled}
            >
              <Choice
                label="Normal"
                icon="margins"
                description={"Top: 1″   Bottom: 1″   Left: 1″   Right: 1″"}
                onClick={() => onAction("margin:72")}
              />
              <Choice
                label="Narrow"
                icon="margins"
                description={
                  "Top: 0.5″   Bottom: 0.5″   Left: 0.5″   Right: 0.5″"
                }
                onClick={() => onAction("margin:36")}
              />
              <Choice
                label="Moderate"
                icon="margins"
                description={
                  "Top: 1″   Bottom: 1″   Left: 0.75″   Right: 0.75″"
                }
                onClick={() => onAction("margin:54")}
              />
              <Choice
                label="Wide"
                icon="margins"
                description={"Top: 1″   Bottom: 1″   Left: 2″   Right: 2″"}
                onClick={() => onAction("margins:72,144,72,144")}
              />
              <Choice label="Mirrored" disabled />
              <hr />
              <Choice
                label="Custom Margins…"
                onClick={() => onAction("page")}
              />
            </Dropdown>
            <Dropdown
              label="Orientation"
              icon="orientation"
              large
              disabled={pageDisabled}
            >
              {["Portrait", "Landscape"].map((label) => (
                <Choice
                  key={label}
                  label={label}
                  icon="orientation"
                  checked={
                    label ===
                    ((document?.page?.widthPt ?? 0) >
                    (document?.page?.heightPt ?? 0)
                      ? "Landscape"
                      : "Portrait")
                  }
                  onClick={() => onAction(`orientation:${label}`)}
                />
              ))}
            </Dropdown>
            <Dropdown label="Size" icon="size" large disabled={pageDisabled}>
              {[
                ["Letter", "612,792", "8.5″ × 11″"],
                ["Legal", "612,1008", "8.5″ × 14″"],
                ["A3", "841.9,1190.6", "29.7 cm × 42 cm"],
                ["A4", "595.3,841.9", "21 cm × 29.7 cm"],
                ["A5", "419.5,595.3", "14.8 cm × 21 cm"],
              ].map(([label, value, description]) => (
                <Choice
                  key={label}
                  label={label}
                  description={description}
                  icon="blank"
                  onClick={() => onAction(`size:${value}`)}
                />
              ))}
              <hr />
              <Choice
                label="More Paper Sizes…"
                onClick={() => onAction("page")}
              />
            </Dropdown>
            {command("Columns", "columns")}
            <div className="office-stack">
              <Dropdown
                label="Breaks"
                icon="pageBreak"
                disabled={paragraphDisabled}
              >
                <p className="office-menu-heading">Page Breaks</p>
                <Choice
                  label="Page"
                  icon="pageBreak"
                  description="Start the paragraph on the next page."
                  onClick={() => onParagraph({ pageBreakBefore: true })}
                />
                <Choice label="Column" disabled />
                <Choice label="Text Wrapping" disabled />
                <p className="office-menu-heading">Section Breaks</p>
                {["Next Page", "Continuous", "Even Page", "Odd Page"].map(
                  (label) => (
                    <Choice key={label} label={label} disabled />
                  ),
                )}
              </Dropdown>
              {command("Line Numbers", "lineNumbers", false)}
              {command("Hyphenation", "hyphenation", false)}
            </div>
          </Group>
          <Group label="Paragraph" more={() => onAction("paragraph")}>
            {[
              { heading: "Indent", fields: indents, icon: "columns" },
              { heading: "Spacing", fields: spacings, icon: "alignLeft" },
            ].map(({ heading, fields, icon }) => (
              <div className="office-paragraph-fields" key={heading}>
                <span className="office-field-heading">{heading}</span>
                {fields.map(({ label, key, value }) => (
                  <label key={key}>
                    <RibbonIcon name={icon} size={16} />
                    <span>{label}:</span>
                    <span className="office-number-field">
                      <input
                        aria-label={`${heading} ${label}`}
                        key={`${selected?.id}-${value}`}
                        type="number"
                        min={heading === "Spacing" ? 0 : -720}
                        max="1440"
                        step="6"
                        defaultValue={value}
                        disabled={paragraphDisabled}
                        onBlur={(e) => {
                          const n = e.currentTarget.valueAsNumber;
                          if (
                            Number.isFinite(n) &&
                            e.currentTarget.validity.valid &&
                            n !== value
                          )
                            onParagraph({ [key]: n });
                        }}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") e.currentTarget.blur();
                        }}
                      />
                      <span>pt</span>
                    </span>
                  </label>
                ))}
              </div>
            ))}
          </Group>
          <Group label="Arrange">
            {command("Position", "position")}
            {command("Wrap Text", "wrap")}
            {command("Bring Forward", "forward")}
            {command("Send Backward", "backward")}
            {command("Selection Pane", "selection", true, "outline")}
            <div className="office-stack">
              {command("Align", "alignLeft", false)}
              {command("Group", "group", false)}
              {command("Rotate", "rotate", false)}
            </div>
          </Group>
        </>
      )}
    </div>
  );
}
