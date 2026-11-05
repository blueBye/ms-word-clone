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
import { ArrangeControls } from "./ArrangeControls";
import { ShapesGallery } from "./GraphicsGallery";

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
  objectCount?: number;
  hasGroup?: boolean;
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
  objectCount = 0,
  hasGroup = false,
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
              <Dropdown label="Cover Page" icon="cover">
                <Choice
                  label="Insert Cover Page…"
                  onClick={() => onAction("dialog:cover")}
                />
                <Choice
                  label="Remove Current Cover Page"
                  onClick={() => onAction("cover:remove")}
                />
              </Dropdown>
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
                onClick={() => onAction("break:page")}
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
              <Choice
                label="Draw Table"
                onClick={() => onAction("dialog:drawTable")}
              />
              <Choice
                label="Convert Text to Table…"
                disabled={!canFormat}
                onClick={() => onAction("dialog:convertTable")}
              />
              <Choice
                label="Excel Spreadsheet"
                onClick={() => onAction("asset:spreadsheet")}
              />
              <Choice
                label="Quick Tables"
                onClick={() => onAction("dialog:quickTable")}
              />
            </Dropdown>
          </Group>
          <Group label="Illustrations">
            {command("Pictures", "picture", true, "picture")}
            <Dropdown
              label="Shapes"
              icon="shapes"
              large
              disabled={disabled || !document}
            >
              <ShapesGallery onPick={(kind) => onAction(`shape:${kind}`)} />
            </Dropdown>
            {command("Icons", "icons", true, "dialog:icons")}
            {command("3D Models", "model", true, "asset:model3d")}
            {command("SmartArt", "smartArt", true, "dialog:smartArt")}
            {command("Chart", "chart", true, "dialog:chart")}
            {command("Screenshot", "screenshot", true, "screenshot")}
          </Group>
          <Group label="Media">
            {command("Online Videos", "video", true, "video")}
          </Group>
          <Group label="Links">
            <div className="office-stack">
              <Command
                label="Link"
                icon="link"
                disabled={disabled || !canFormat}
                onClick={() => onAction("link")}
              />
              {command("Bookmark", "bookmark", false, "dialog:bookmark")}
              {command(
                "Cross-reference",
                "reference",
                false,
                "dialog:reference",
              )}
            </div>
          </Group>
          <Group label="Comments">
            {command("Comment", "comment", true, "dialog:comment")}
          </Group>
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
                  label={`Edit ${label}`}
                  icon={label.toLowerCase()}
                  onClick={() => onAction(`story:${label.toLowerCase()}:edit`)}
                />
                <Choice
                  label={`Remove ${label}`}
                  onClick={() =>
                    onAction(`story:${label.toLowerCase()}:remove`)
                  }
                />
              </Dropdown>
            ))}
            <Dropdown
              label="Page Number"
              icon="pageNumber"
              large
              disabled={disabled || !document}
            >
              <Choice
                label="Top of Page"
                icon="header"
                onClick={() => onAction("story:header:number")}
              />
              <Choice
                label="Bottom of Page"
                icon="footer"
                onClick={() => onAction("story:footer:number")}
              />
            </Dropdown>
          </Group>
          <Group label="Text">
            {command("Text Box", "textBox", true, "dialog:textBox")}
            {command("Quick Parts", "quickParts", true, "dialog:quickParts")}
            {command("WordArt", "wordArt", true, "dialog:wordArt")}
            {command("Drop Cap", "dropCap", true, "dialog:dropCap")}
            <div className="office-stack">
              {command(
                "Signature Line",
                "signature",
                false,
                "dialog:signature",
              )}
              <Command
                label="Date & Time"
                icon="date"
                disabled={disabled || !canFormat}
                onClick={() => onAction("dialog:dateOptions")}
              />
              <Dropdown label="Object" icon="object" disabled={disabled}>
                <Choice
                  label="Object…"
                  onClick={() => onAction("asset:object")}
                />
                <Choice
                  label="Text from File…"
                  onClick={() => onAction("textFromFile")}
                />
              </Dropdown>
            </div>
          </Group>
          <Group label="Symbols">
            <div className="office-stack">
              {command("Equation", "equation", false, "equation:new")}
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
                <Choice
                  label="More Symbols…"
                  onClick={() => onAction("dialog:symbols")}
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
              <Choice
                label="Mirrored"
                onClick={() => onAction("layout:mirror")}
              />
              <hr />
              <Choice
                label="Custom Margins…"
                onClick={() => onAction("page")}
              />
              <Choice
                label="Layout Options…"
                onClick={() => onAction("dialog:layoutOptions")}
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
            <Dropdown
              label="Columns"
              icon="columns"
              large
              disabled={pageDisabled}
            >
              {[
                ["One", 1],
                ["Two", 2],
                ["Three", 3],
              ].map(([label, count]) => (
                <Choice
                  key={label}
                  label={String(label)}
                  icon="columns"
                  checked={(document?.page?.columns?.count ?? 1) === count}
                  onClick={() => onAction(`columns:${count}`)}
                />
              ))}
              <Choice
                label="More Columns…"
                onClick={() => onAction("dialog:columnsOptions")}
              />
            </Dropdown>
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
                  onClick={() => onAction("break:page")}
                />
                <Choice
                  label="Column"
                  onClick={() => onAction("break:column")}
                />
                <Choice
                  label="Text Wrapping"
                  onClick={() => onAction("break:wrap")}
                />
                <p className="office-menu-heading">Section Breaks</p>
                {[
                  ["Next Page", "nextPage"],
                  ["Continuous", "continuous"],
                  ["Even Page", "evenPage"],
                  ["Odd Page", "oddPage"],
                ].map(([label, value]) => (
                  <Choice
                    key={label}
                    label={label}
                    onClick={() => onAction(`section:${value}`)}
                  />
                ))}
              </Dropdown>
              <Dropdown
                label="Line Numbers"
                icon="lineNumbers"
                disabled={pageDisabled}
              >
                {[
                  ["None", "none"],
                  ["Continuous", "continuous"],
                  ["Restart Each Page", "page"],
                  ["Restart Each Section", "section"],
                ].map(([label, value]) => (
                  <Choice
                    key={value}
                    label={label}
                    onClick={() => onAction(`lineNumbers:${value}`)}
                  />
                ))}
                <Choice
                  label="Suppress for Current Paragraph"
                  onClick={() =>
                    onParagraph({
                      suppressLineNumbers:
                        !selected?.properties.suppressLineNumbers,
                    })
                  }
                />
                <Choice
                  label="Line Numbering Options…"
                  onClick={() => onAction("dialog:lineNumbers")}
                />
              </Dropdown>
              <Dropdown
                label="Hyphenation"
                icon="hyphenation"
                disabled={pageDisabled}
              >
                <Choice
                  label="None"
                  onClick={() => onAction("hyphenation:none")}
                />
                <Choice
                  label="Automatic"
                  onClick={() => onAction("hyphenation:auto")}
                />
                <Choice
                  label="Manual…"
                  onClick={() => onAction("dialog:hyphenation")}
                />
                <Choice
                  label="Hyphenation Options…"
                  onClick={() => onAction("dialog:hyphenation")}
                />
              </Dropdown>
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
            <ArrangeControls
              count={objectCount}
              group={hasGroup}
              onAction={onAction}
            />{" "}
          </Group>
        </>
      )}
    </div>
  );
}
