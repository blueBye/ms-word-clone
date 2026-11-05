import type {
  DrawingInline,
  EditCommand,
  Table,
  TableCell,
} from "@local/docx-ir-engine/core";
import { Group } from "./Controls";
import {
  RibbonCommand as Command,
  RibbonDropdown as Dropdown,
  RibbonChoice as Choice,
} from "./RibbonControls";
import { ArrangeControls } from "./ArrangeControls";
export type ContextTab =
  | "Picture Format"
  | "Table Design"
  | "Table Layout"
  | "Equation"
  | "Shape Format";
type TableCommand = Extract<EditCommand, { type: "table" }>;
const colors = [
  "#FFFFFF",
  "#000000",
  "#4472C4",
  "#ED7D31",
  "#70AD47",
  "#E8EFF8",
  "#FFF2CC",
  "#E2F0D9",
  "#FCE4D6",
  "#D9E1F2",
  "#F2F2F2",
  "#FFFF00",
];
function ColorChoices({ onColor }: { onColor: (color: string) => void }) {
  return (
    <div className="context-color-grid">
      {colors.map((color) => (
        <button
          type="button"
          aria-label={color}
          title={color}
          key={color}
          style={{ background: color }}
          onClick={() => onColor(color)}
        />
      ))}
      <label>
        Custom
        <input
          type="color"
          aria-label="Custom color"
          onChange={(e) => onColor(e.target.value)}
        />
      </label>
    </div>
  );
}
function Dimension({
  label,
  value,
  onChange,
}: {
  label: string;
  value: number;
  onChange: (value: number) => void;
}) {
  return (
    <label className="context-dimension">
      {label}
      <input
        key={value}
        aria-label={label}
        type="number"
        min="0.02"
        max="220"
        step="0.01"
        defaultValue={(value / 72).toFixed(2)}
        onBlur={(e) => {
          const n = e.target.valueAsNumber * 72;
          if (Number.isFinite(n) && n > 0 && Math.abs(n - value) > 0.1)
            onChange(n);
        }}
      />
      <span>″</span>
    </label>
  );
}
export function ContextRibbon({
  tab,
  object,
  table,
  disabled,
  onDrawing,
  onTable,
  onAction,
}: {
  tab: ContextTab;
  object?: DrawingInline;
  table?: { table: Table; row: number; cell: TableCell };
  disabled: boolean;
  onDrawing: (drawing: DrawingInline) => void;
  onTable: (command: Omit<TableCommand, "type" | "id">) => void;
  onAction: (name: string) => void;
}) {
  const image = (patch: Partial<NonNullable<DrawingInline["image"]>>) =>
    object?.image &&
    onDrawing({ ...object, image: { ...object.image, ...patch } });
  const tableAction = (action: TableCommand["action"]) => () =>
    onTable({ action });
  const tableFormat = (format: TableCommand["format"]) => () =>
    onTable({ format });
  return (
    <div
      className="ribbon office-ribbon context-ribbon"
      id={`ribbon-${tab.replaceAll(" ", "-")}`}
      role="tabpanel"
      aria-labelledby={`tab-${tab.replaceAll(" ", "-")}`}
      aria-disabled={disabled}
      inert={disabled || undefined}
    >
      {object && ["Picture Format", "Shape Format"].includes(tab) && (
        <>
          {object.image ? (
            <Group label="Adjust">
              <Dropdown
                label="Corrections"
                icon="corrections"
                large
                disabled={disabled}
              >
                <p className="office-menu-heading">Brightness and Contrast</p>
                <div className="picture-corrections-grid">
                  {[-0.4, -0.2, 0, 0.2, 0.4].flatMap((contrast) =>
                    [-0.4, -0.2, 0, 0.2, 0.4].map((brightness) => (
                      <button
                        key={`${brightness}-${contrast}`}
                        type="button"
                        title={`Brightness ${brightness * 100}%, contrast ${contrast * 100}%`}
                        aria-label={`Brightness ${brightness * 100}%, contrast ${contrast * 100}%`}
                        onClick={() => image({ brightness, contrast })}
                      >
                        <img
                          src={object.image!.dataUrl}
                          alt=""
                          style={{
                            filter: `brightness(${1 + brightness}) contrast(${1 + contrast})`,
                          }}
                        />
                      </button>
                    )),
                  )}
                </div>
                <Choice
                  label="Picture Corrections Options…"
                  onClick={() => onAction("picture:options")}
                />
              </Dropdown>
              <div className="office-stack">
                <Dropdown label="Color" icon="picture" disabled={disabled}>
                  <Choice
                    label="Original Color"
                    onClick={() => image({ grayscale: false })}
                  />
                  <Choice
                    label="Grayscale"
                    onClick={() => image({ grayscale: true })}
                  />
                </Dropdown>
                <Dropdown
                  label="Artistic Effects"
                  icon="effects"
                  disabled={disabled}
                >
                  {[
                    ["None", 0],
                    ["Soft Blur", 1],
                    ["Blur", 3],
                    ["Strong Blur", 6],
                  ].map(([name, blur]) => (
                    <Choice
                      key={name}
                      label={String(name)}
                      onClick={() => image({ blurPt: Number(blur) })}
                    />
                  ))}
                </Dropdown>
                <Dropdown
                  label="Transparency"
                  icon="transparency"
                  disabled={disabled}
                >
                  {[0, 15, 30, 50, 65, 80, 95].map((n) => (
                    <Choice
                      key={n}
                      label={`${n}%`}
                      onClick={() => image({ opacity: 1 - n / 100 })}
                    />
                  ))}
                </Dropdown>
              </div>
              <div className="office-stack">
                <Command
                  label="Change Picture"
                  icon="picture"
                  disabled={disabled}
                  onClick={() => onAction("picture:replace")}
                />
                <Command
                  label="Reset Picture"
                  icon="undo"
                  disabled={disabled}
                  onClick={() =>
                    image({
                      brightness: 0,
                      contrast: 0,
                      grayscale: false,
                      blurPt: 0,
                      opacity: 1,
                      crop: { left: 0, right: 0, top: 0, bottom: 0 },
                    })
                  }
                />
              </div>
            </Group>
          ) : (
            <Group label="Shape Styles">
              <Dropdown label="Shape Fill" icon="shapes" disabled={disabled}>
                <ColorChoices
                  onColor={(fill) =>
                    onDrawing({
                      ...object,
                      shape: {
                        ...object.shape,
                        kind: object.shape?.kind ?? "rect",
                        fill,
                      },
                    })
                  }
                />
                <Choice
                  label="No Fill"
                  onClick={() =>
                    onDrawing({
                      ...object,
                      shape: {
                        ...object.shape,
                        kind: object.shape?.kind ?? "rect",
                        fill: undefined,
                      },
                    })
                  }
                />
              </Dropdown>
            </Group>
          )}
          <Group
            label={object.image ? "Picture Styles" : "Shape Outline"}
            more={() => onAction("picture:options")}
          >
            <Dropdown
              label={object.image ? "Picture Border" : "Shape Outline"}
              icon="borders"
              large
              disabled={disabled}
            >
              <ColorChoices
                onColor={(stroke) =>
                  onDrawing({
                    ...object,
                    shape: {
                      ...object.shape,
                      kind: object.shape?.kind ?? "rect",
                      stroke,
                      strokeWidthPt: object.shape?.strokeWidthPt ?? 1,
                    },
                  })
                }
              />
              <Choice
                label="No Outline"
                onClick={() =>
                  onDrawing({
                    ...object,
                    shape: {
                      ...object.shape,
                      kind: object.shape?.kind ?? "rect",
                      stroke: undefined,
                    },
                  })
                }
              />
              {[0.5, 1, 2, 3, 6].map((n) => (
                <Choice
                  key={n}
                  label={`${n} pt`}
                  onClick={() =>
                    onDrawing({
                      ...object,
                      shape: {
                        ...object.shape,
                        kind: object.shape?.kind ?? "rect",
                        stroke: object.shape?.stroke ?? "#000000",
                        strokeWidthPt: n,
                      },
                    })
                  }
                />
              ))}
            </Dropdown>
            <Command
              label="Alt Text"
              icon="edit"
              large
              disabled={disabled}
              onClick={() => onAction("picture:options")}
            />
          </Group>
          <Group label="Arrange">
            <ArrangeControls
              count={1}
              group={!!object.children}
              onAction={onAction}
            />
          </Group>
          <Group label="Size" more={() => onAction("picture:options")}>
            {object.image && (
              <Dropdown label="Crop" icon="crop" large disabled={disabled}>
                <Choice
                  label="Crop…"
                  onClick={() => onAction("picture:crop")}
                />
                <Choice
                  label="Reset Crop"
                  onClick={() =>
                    image({ crop: { left: 0, right: 0, top: 0, bottom: 0 } })
                  }
                />
              </Dropdown>
            )}
            <div className="office-stack">
              <Dimension
                label="Height"
                value={object.heightPt}
                onChange={(heightPt) =>
                  onDrawing({
                    ...object,
                    heightPt,
                    widthPt: (object.widthPt * heightPt) / object.heightPt,
                  })
                }
              />
              <Dimension
                label="Width"
                value={object.widthPt}
                onChange={(widthPt) =>
                  onDrawing({
                    ...object,
                    widthPt,
                    heightPt: (object.heightPt * widthPt) / object.widthPt,
                  })
                }
              />
            </div>
          </Group>
        </>
      )}
      {table && tab === "Table Design" && (
        <>
          <Group label="Table Style Options">
            <div className="office-stack">
              <label className="ribbon-check">
                <input
                  type="checkbox"
                  checked={!!table.table.rows[0].header}
                  onChange={(e) =>
                    onTable({
                      range: {
                        top: 0,
                        bottom: 0,
                        left: 0,
                        right:
                          Math.max(
                            table.table.columnsPt.length,
                            ...table.table.rows.flatMap((r) =>
                              r.cells.map((c) => c.column + c.columnSpan),
                            ),
                          ) - 1,
                      },
                      format: { repeatHeader: e.target.checked },
                    })
                  }
                />
                Header Row
              </label>
              <Command
                label="Banded Rows"
                icon="table"
                onClick={tableFormat({ style: "banded" })}
              />
            </div>
          </Group>
          <Group label="Table Styles">
            <div className="table-style-gallery">
              {(["grid", "banded", "plain"] as const).map((style) => (
                <button
                  type="button"
                  key={style}
                  aria-label={`Apply ${style} table style`}
                  onClick={tableFormat({ style })}
                >
                  <span className={`mini-table ${style}`}>
                    {Array.from({ length: 12 }, (_, i) => (
                      <i key={i} />
                    ))}
                  </span>
                  <span>
                    {style === "grid"
                      ? "Table Grid"
                      : style === "banded"
                        ? "Grid Table — Accent 1"
                        : "Plain Table"}
                  </span>
                </button>
              ))}
            </div>
            <Dropdown label="Shading" icon="shading" large disabled={disabled}>
              <ColorChoices
                onColor={(shading) => onTable({ format: { shading } })}
              />
              <Choice
                label="No Color"
                onClick={tableFormat({ shading: "none" })}
              />
            </Dropdown>
          </Group>
          <Group label="Borders">
            <Dropdown label="Borders" icon="borders" large disabled={disabled}>
              {(["all", "outside", "bottom", "none"] as const).map(
                (borders) => (
                  <Choice
                    key={borders}
                    label={`${borders === "none" ? "No" : borders[0].toUpperCase() + borders.slice(1)} Borders`}
                    onClick={tableFormat({ borders })}
                  />
                ),
              )}
            </Dropdown>
            <Command
              label="Borders and Shading"
              icon="borders"
              disabled={disabled}
              onClick={() => onAction("table:properties")}
            />
          </Group>
        </>
      )}
      {table && tab === "Table Layout" && (
        <>
          <Group label="Table">
            <div className="office-stack">
              <Dropdown label="Select" icon="select" disabled={disabled}>
                {["Cell", "Row", "Column", "Table"].map((value) => (
                  <Choice
                    key={value}
                    label={`Select ${value}`}
                    onClick={() =>
                      onAction(`table:select:${value.toLowerCase()}`)
                    }
                  />
                ))}
              </Dropdown>
              <Command
                label="Properties"
                icon="table"
                disabled={disabled}
                onClick={() => onAction("table:properties")}
              />
            </div>
          </Group>
          <Group label="Rows & Columns">
            <Dropdown label="Delete" icon="delete" large disabled={disabled}>
              {[
                ["Delete Cells…", "table:deleteCells"],
                ["Delete Columns", "table:deleteColumn"],
                ["Delete Rows", "table:deleteRow"],
                ["Delete Table", "table:deleteTable"],
              ].map(([label, action]) => (
                <Choice
                  key={label}
                  label={label}
                  onClick={() => onAction(action)}
                />
              ))}
            </Dropdown>
            <Command
              label="Insert Above"
              icon="insertAbove"
              large
              disabled={disabled}
              onClick={tableAction("insertRowAbove")}
            />
            <div className="office-stack">
              {[
                ["Insert Below", "insertRowBelow"],
                ["Insert Left", "insertColumnBefore"],
                ["Insert Right", "insertColumnAfter"],
              ].map(([label, action]) => (
                <Command
                  key={label}
                  label={label}
                  icon="table"
                  disabled={disabled}
                  onClick={tableAction(action as TableCommand["action"])}
                />
              ))}
            </div>
          </Group>
          <Group label="Merge">
            <div className="office-stack">
              <Command
                label="Merge Cells"
                icon="mergeCells"
                disabled={disabled}
                onClick={tableAction("merge")}
              />
              <Command
                label="Split Cells"
                icon="splitCells"
                disabled={disabled}
                onClick={() => onAction("table:split")}
              />
            </div>
          </Group>
          <Group label="Cell Size" more={() => onAction("table:properties")}>
            <Dropdown label="AutoFit" icon="table" large disabled={disabled}>
              {[
                ["AutoFit Contents", "autofitContents"],
                ["AutoFit Window", "autofitWindow"],
                ["Fixed Column Width", "fixedWidth"],
              ].map(([label, action]) => (
                <Choice
                  key={label}
                  label={label}
                  onClick={tableAction(action as TableCommand["action"])}
                />
              ))}
            </Dropdown>
            <div className="office-stack">
              <Dimension
                label="Row Height"
                value={table.table.rows[table.row].heightPt ?? 18}
                onChange={(rowHeightPt) => onTable({ format: { rowHeightPt } })}
              />
              <Dimension
                label="Column Width"
                value={table.table.columnsPt[table.cell.column] ?? 72}
                onChange={(columnWidthPt) =>
                  onTable({ format: { columnWidthPt } })
                }
              />
            </div>
            <div className="office-stack">
              <Command
                label="Distribute Rows"
                icon="table"
                disabled={disabled}
                onClick={tableAction("distributeRows")}
              />
              <Command
                label="Distribute Columns"
                icon="columns"
                disabled={disabled}
                onClick={tableAction("distributeColumns")}
              />
            </div>
          </Group>
          <Group label="Alignment">
            <div className="cell-alignment-grid">
              {(["top", "center", "bottom"] as const).flatMap((v) =>
                (["left", "center", "right"] as const).map((h) => (
                  <button
                    key={`${v}-${h}`}
                    type="button"
                    aria-label={`Align ${v} ${h}`}
                    title={`Align ${v} ${h}`}
                    onClick={() => onAction(`table:align:${v}:${h}`)}
                  >
                    <span
                      style={{
                        alignSelf:
                          v === "top"
                            ? "start"
                            : v === "bottom"
                              ? "end"
                              : "center",
                        textAlign: h,
                      }}
                    >
                      ☰
                    </span>
                  </button>
                )),
              )}
            </div>
            <Command
              label="Cell Margins"
              icon="margins"
              large
              disabled={disabled}
              onClick={() => onAction("table:properties")}
            />
          </Group>
          <Group label="Data">
            <div className="office-stack">
              <Command
                label="Repeat Header Rows"
                icon="header"
                disabled={disabled}
                onClick={tableFormat({
                  repeatHeader: !table.table.rows[table.row].header,
                })}
              />
              <Command
                label="Allow Row to Split"
                icon="pageBreak"
                disabled={disabled}
                onClick={tableFormat({
                  cantSplit: !table.table.rows[table.row].cantSplit,
                })}
              />
            </div>
          </Group>
        </>
      )}
      {object?.equation && tab === "Equation" && (
        <>
          <Group label="Tools">
            <Command
              label="Edit Equation"
              icon="equation"
              large
              onClick={() => onAction("equation:focus")}
            />
            <Command
              label="Delete Equation"
              icon="delete"
              onClick={() => onAction("object:delete")}
            />
          </Group>
          <Group label="Symbols">
            <div className="equation-symbols">
              {[
                "±",
                "∞",
                "=",
                "≠",
                "~",
                "×",
                "÷",
                "α",
                "β",
                "γ",
                "π",
                "θ",
                "≤",
                "≥",
                "∈",
                "∂",
                "∇",
                "∀",
                "∃",
                "→",
              ].map((symbol) => (
                <button
                  type="button"
                  key={symbol}
                  onClick={() => onAction(`equation:insert:${symbol}`)}
                >
                  {symbol}
                </button>
              ))}
            </div>
          </Group>
          <Group label="Structures">
            {[
              ["Fraction", "\\frac{a}{b}"],
              ["Script", "x^{2}"],
              ["Radical", "\\sqrt{x}"],
              ["Integral", "\\int_{a}^{b}f(x)"],
              ["Large Operator", "\\sum_{i=1}^{n}x_i"],
              ["Bracket", "(x+y)"],
            ].map(([label, formula]) => (
              <Dropdown key={label} label={label} icon="equation" large>
                <Choice
                  label={formula}
                  onClick={() => onAction(`equation:insert:${formula}`)}
                />
              </Dropdown>
            ))}
          </Group>
        </>
      )}
    </div>
  );
}
