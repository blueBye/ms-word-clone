import { type ReactNode } from "react";
import type {
  Paragraph,
  ParagraphFormat,
  RunFormat,
  EditCommand,
  DocxStyle,
} from "@local/docx-ir-engine";
import { Dialog } from "./Controls";
import { HomeIcon } from "./HomeRibbon";
import { HomeMenu, type MenuAnchor } from "./HomeMenu";
import { ListGallery, ListDefinitionForm } from "./ListGallery";
import { EffectsGallery } from "./EffectsGallery";
import { StyleForm } from "./StyleForm";
import { StylePreview } from "./StylePreview";
import { FontChooser, FontForm, Palette, dialogData } from "./FontDialogs";
import { SearchForm, type SearchProps } from "./SearchForm";

interface Props extends SearchProps {
  dialog: string | null;
  onClose: () => void;
  onDialog: (name: string) => void;
  selected?: Paragraph;
  format: RunFormat;
  families: string[];
  styles: DocxStyle[];
  anchor: MenuAnchor | null;
  contextStyle: string;
  onStyleContext: (id: string, anchor: MenuAnchor) => void;
  onStyleAction: (name: string, id: string) => void;
  clipboard: string[];
  onRun: (format: RunFormat) => void;
  onParagraph: (format: ParagraphFormat) => void;
  onAction: (name: string) => void;
  onText: (text: string) => void;
  onCommand: (command: EditCommand) => void;
  onSort: (count: number, descending: boolean, numeric: boolean) => void;
}
const titles: Record<string, string> = {
  paste: "Paste",
  clipboard: "Clipboard",
  font: "Font",
  fontFamily: "Font",
  fontSize: "Font size",
  case: "Change Case",
  underline: "Underline",
  highlight: "Text Highlight Color",
  color: "Font Color",
  shading: "Shading",
  borders: "Borders",
  bullets: "Bullets",
  numbering: "Numbering",
  multilevel: "Multilevel List",
  spacing: "Line and Paragraph Spacing",
  justify: "Justification",
  effects: "Text Effects and Typography",
  styles: "Styles",
  sort: "Sort Text",
  replace: "Find and Replace",
  findDialog: "Find and Replace",
  goto: "Go To",
  select: "Select",
  styleContext: "Style",
  styleModify: "Modify Style",
  styleRename: "Rename Style",
  defineBullet: "Define New Bullet",
  defineNumber: "Define New Number Format",
  defineMultilevel: "Define New Multilevel List",
  numberValue: "Set Numbering Value",
};
const highlightColors = [
  ["Yellow", "yellow", "#FFFF00"],
  ["Bright Green", "green", "#00FF00"],
  ["Turquoise", "cyan", "#00FFFF"],
  ["Pink", "magenta", "#FF00FF"],
  ["Blue", "blue", "#0000FF"],
  ["Red", "red", "#FF0000"],
  ["Dark Blue", "darkBlue", "#000080"],
  ["Teal", "darkCyan", "#008080"],
  ["Green", "darkGreen", "#008000"],
  ["Violet", "darkMagenta", "#800080"],
  ["Dark Red", "darkRed", "#800000"],
  ["Dark Yellow", "darkYellow", "#808000"],
  ["Gray 50%", "darkGray", "#808080"],
  ["Gray 25%", "lightGray", "#C0C0C0"],
  ["Black", "black", "#000000"],
];
export function HomeDialogs(p: Props) {
  if (!p.dialog) return null;
  const choose = (action: () => void) => {
      action();
      p.onClose();
    },
    id = p.selected?.id ?? "";
  return (
    <HomeSurface
      name={p.dialog}
      title={titles[p.dialog] ?? "Home"}
      anchor={p.anchor}
      onClose={p.onClose}
    >
      <div className={`home-dialog-body home-dialog-${p.dialog}`}>
        {(p.dialog === "paste" || p.dialog === "clipboard") && (
          <>
            <p className="pane-description">
              Paste plain text at the current selection, using destination
              formatting.
            </p>
            <button
              className="home-choice"
              onClick={() => choose(() => p.onAction("paste"))}
            >
              <HomeIcon name="paste" />
              Keep Text Only <kbd>Ctrl V</kbd>
            </button>
            {p.clipboard.length > 0 && (
              <>
                <h3>Copied in this workspace</h3>
                {p.clipboard.map((text, index) => (
                  <button
                    className="clipboard-entry"
                    key={index}
                    onClick={() => choose(() => p.onText(text))}
                  >
                    <span dir="auto">{text.slice(0, 180)}</span>
                  </button>
                ))}
              </>
            )}
            <p className="pane-description">
              You can also click the page and press Ctrl+V. Office clipboard
              history and embedded-object Paste Special formats are not
              available.
            </p>
          </>
        )}
        {p.dialog === "fontFamily" && (
          <FontChooser
            families={p.families}
            value={p.format.fontFamily ?? "Calibri"}
            onPick={(fontFamily) => choose(() => p.onRun({ fontFamily }))}
            onManage={() => choose(() => p.onAction("fonts"))}
          />
        )}
        {p.dialog === "fontSize" && (
          <div className="home-choice-grid">
            {[
              8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48,
              72,
            ].map((fontSizePt) => (
              <button
                key={fontSizePt}
                className="home-choice"
                onClick={() => choose(() => p.onRun({ fontSizePt }))}
              >
                {fontSizePt} pt
              </button>
            ))}
          </div>
        )}
        {p.dialog === "font" && (
          <FontForm
            format={p.format}
            onApply={(format) => choose(() => p.onRun(format))}
          />
        )}
        {p.dialog === "case" && (
          <div className="home-choice-list">
            {[
              ["Sentence case.", "sentence"],
              ["lowercase", "lower"],
              ["UPPERCASE", "upper"],
              ["Capitalize Each Word", "title"],
              ["tOGGLE cASE", "toggle"],
            ].map(([label, mode]) => (
              <button
                key={mode}
                className="home-choice"
                onClick={() => choose(() => p.onAction(`case:${mode}`))}
              >
                {label}
              </button>
            ))}
          </div>
        )}
        {p.dialog === "underline" && (
          <div className="home-choice-list">
            {[
              ["Single", "single"],
              ["Double", "double"],
              ["Thick", "thick"],
              ["Dotted", "dotted"],
              ["Dashed", "dash"],
              ["Dot-dash", "dotDash"],
              ["Dot-dot-dash", "dotDotDash"],
              ["Wave", "wave"],
              ["None", "none"],
            ].map(([label, underline]) => (
              <button
                className="home-choice"
                key={underline}
                onClick={() =>
                  choose(() =>
                    p.onRun({ underline: underline as RunFormat["underline"] }),
                  )
                }
              >
                <span
                  className="underline-preview"
                  style={{
                    textDecorationLine:
                      underline === "none" ? "none" : "underline",
                    textDecorationStyle:
                      underline === "double"
                        ? "double"
                        : underline === "wave"
                          ? "wavy"
                          : underline === "dotted"
                            ? "dotted"
                            : /dash/i.test(underline)
                              ? "dashed"
                              : "solid",
                    textDecorationThickness: underline === "thick" ? 3 : 1,
                  }}
                >
                  AaBbCc
                </span>
                {label}
              </button>
            ))}
          </div>
        )}
        {p.dialog === "highlight" && (
          <>
            <div className="highlight-palette">
              {highlightColors.map(([name, highlight, color]) => (
                <button
                  key={highlight}
                  aria-label={name}
                  title={name}
                  style={{ background: color }}
                  onClick={() => choose(() => p.onRun({ highlight }))}
                />
              ))}
            </div>
            <button
              className="home-choice"
              onClick={() => choose(() => p.onRun({ highlight: "none" }))}
            >
              No Color
            </button>
          </>
        )}
        {(p.dialog === "color" || p.dialog === "shading") && (
          <>
            <button
              className="home-choice"
              onClick={() =>
                choose(() =>
                  p.dialog === "color"
                    ? p.onRun({ color: "auto" })
                    : p.onParagraph({ shading: "none" }),
                )
              }
            >
              {p.dialog === "color" ? "Automatic" : "No Color"}
            </button>
            <Palette
              onPick={(color) =>
                choose(() =>
                  p.dialog === "color"
                    ? p.onRun({ color })
                    : p.onParagraph({ shading: color }),
                )
              }
            />
          </>
        )}
        {p.dialog === "borders" && (
          <div className="home-choice-list">
            {[
              ["Bottom Border", "bottom"],
              ["Top Border", "top"],
              ["Left Border", "left"],
              ["Right Border", "right"],
              ["No Border", "none"],
              ["All Borders", "all"],
              ["Outside Borders", "outside"],
            ].map(([label, borders]) => (
              <button
                key={borders}
                className="home-choice"
                onClick={() =>
                  choose(() =>
                    p.onParagraph({
                      borders: borders as ParagraphFormat["borders"],
                    }),
                  )
                }
              >
                <HomeIcon name="borders" />
                {label}
              </button>
            ))}
          </div>
        )}
        {(p.dialog === "bullets" ||
          p.dialog === "numbering" ||
          p.dialog === "multilevel") && (
          <ListGallery
            mode={p.dialog}
            selected={p.selected}
            apply={(command) => choose(() => p.onCommand(command))}
            onDialog={p.onDialog}
          />
        )}
        {[
          "defineBullet",
          "defineNumber",
          "defineMultilevel",
          "numberValue",
        ].includes(p.dialog) && (
          <ListDefinitionForm
            key={p.dialog}
            mode={p.dialog}
            selected={p.selected}
            apply={(command) => choose(() => p.onCommand(command))}
          />
        )}

        {p.dialog === "spacing" && (
          <>
            <div className="home-choice-grid">
              {[1, 1.15, 1.5, 2, 2.5, 3].map((lineSpacing) => (
                <button
                  key={lineSpacing}
                  className="home-choice"
                  onClick={() => choose(() => p.onParagraph({ lineSpacing }))}
                >
                  {lineSpacing}
                </button>
              ))}
            </div>
            <button
              className="home-choice"
              onClick={() =>
                choose(() =>
                  p.onParagraph({
                    spaceBeforePt: p.selected?.properties.spaceBeforePt
                      ? 0
                      : 12,
                  }),
                )
              }
            >
              {p.selected?.properties.spaceBeforePt ? "Remove" : "Add"} Space
              Before Paragraph
            </button>
            <button
              className="home-choice"
              onClick={() =>
                choose(() =>
                  p.onParagraph({
                    spaceAfterPt: p.selected?.properties.spaceAfterPt ? 0 : 8,
                  }),
                )
              }
            >
              {p.selected?.properties.spaceAfterPt ? "Remove" : "Add"} Space
              After Paragraph
            </button>
            <button
              className="home-choice"
              onClick={() => choose(() => p.onAction("paragraph"))}
            >
              Line Spacing Options…
            </button>
          </>
        )}
        {p.dialog === "justify" && (
          <>
            <button
              className="home-choice"
              onClick={() =>
                choose(() => p.onParagraph({ alignment: "justify" }))
              }
            >
              <HomeIcon name="justify" />
              Justify
            </button>
            <p className="pane-description">
              Word’s Low, Medium and High Arabic justification need additional
              shaping support. Standard justification is available.
            </p>
          </>
        )}
        {p.dialog === "effects" && (
          <EffectsGallery
            format={p.format}
            onPick={(format) => choose(() => p.onRun(format))}
          />
        )}
        {p.dialog === "styles" && (
          <>
            <div className="home-styles-list">
              {p.styles
                .filter((s) => s.gallery)
                .map((style) => (
                  <button
                    key={style.id}
                    onContextMenu={(e) => {
                      e.preventDefault();
                      p.onStyleContext(style.id, {
                        left: e.clientX,
                        right: e.clientX,
                        top: e.clientY,
                        bottom: e.clientY,
                      });
                    }}
                    onClick={() =>
                      choose(() => p.onParagraph({ styleId: style.id }))
                    }
                  >
                    <StylePreview style={style} />
                    <small>{style.name}</small>
                  </button>
                ))}
            </div>
            <hr />
            <button
              className="home-choice"
              onClick={() => choose(() => p.onAction("clear"))}
            >
              Clear Formatting
            </button>
            <button
              className="home-choice"
              onClick={() => p.onDialog("allStyles")}
            >
              Apply Styles…
            </button>
          </>
        )}
        {p.dialog === "allStyles" && (
          <div className="home-styles-list">
            {p.styles.map((style) => (
              <button
                key={style.id}
                onContextMenu={(e) => {
                  e.preventDefault();
                  p.onStyleContext(style.id, {
                    left: e.clientX,
                    right: e.clientX,
                    top: e.clientY,
                    bottom: e.clientY,
                  });
                }}
                onClick={() =>
                  choose(() => p.onParagraph({ styleId: style.id }))
                }
              >
                <StylePreview style={style} />
              </button>
            ))}
          </div>
        )}
        {p.dialog === "styleContext" &&
          (() => {
            const style = p.styles.find((s) => s.id === p.contextStyle);
            return (
              style && (
                <>
                  <button
                    className="home-choice"
                    onClick={() =>
                      choose(() => p.onStyleAction("update", style.id))
                    }
                  >
                    Update {style.name} to Match Selection
                  </button>
                  <button
                    className="home-choice"
                    onClick={() => p.onDialog("styleModify")}
                  >
                    Modify…
                  </button>
                  <button
                    className="home-choice"
                    onClick={() =>
                      choose(() => p.onStyleAction("select", style.id))
                    }
                  >
                    Select All Instances
                  </button>
                  <hr />
                  <button
                    className="home-choice"
                    onClick={() => p.onDialog("styleRename")}
                  >
                    Rename…
                  </button>
                  <button
                    className="home-choice"
                    onClick={() =>
                      choose(() => p.onStyleAction("remove", style.id))
                    }
                  >
                    Remove from Style Gallery
                  </button>
                </>
              )
            );
          })()}
        {(p.dialog === "styleModify" || p.dialog === "styleRename") &&
          (() => {
            const style = p.styles.find((s) => s.id === p.contextStyle);
            return (
              style && (
                <StyleForm
                  key={p.dialog + style.id}
                  style={style}
                  styles={p.styles}
                  rename={p.dialog === "styleRename"}
                  onApply={(command) => choose(() => p.onCommand(command))}
                />
              )
            );
          })()}
        {p.dialog === "sort" && (
          <form
            onSubmit={(event) => {
              const d = dialogData(event);
              choose(() =>
                p.onSort(
                  Number(d.get("count")),
                  d.get("order") === "descending",
                  d.get("type") === "number",
                ),
              );
            }}
          >
            <p className="pane-description">
              Sort consecutive paragraphs starting at the selected paragraph.
              Tables, fields and section boundaries are protected.
            </p>
            <div className="field-row">
              <label className="field">
                Sort by
                <select>
                  <option>Paragraphs</option>
                </select>
              </label>
              <label className="field">
                Type
                <select name="type">
                  <option value="text">Text</option>
                  <option value="number">Number</option>
                </select>
              </label>
            </div>
            <label className="field">
              Number of paragraphs
              <input
                name="count"
                type="number"
                min="2"
                max="100"
                defaultValue="2"
                required
              />
            </label>
            <label className="field">
              Order
              <select name="order">
                <option value="ascending">Ascending</option>
                <option value="descending">Descending</option>
              </select>
            </label>
            <div className="dialog-actions">
              <button className="primary-button">Sort</button>
            </div>
          </form>
        )}
        {(p.dialog === "replace" || p.dialog === "findDialog") && (
          <SearchForm
            key={p.dialog}
            query={p.query}
            replacement={p.replacement}
            options={p.options}
            replace={p.dialog === "replace"}
            onFind={p.onFind}
            onReplace={p.onReplace}
          />
        )}
        {p.dialog === "findMenu" && (
          <>
            <button
              className="home-choice"
              onClick={() => choose(() => p.onAction("find"))}
            >
              Find
            </button>
            <button
              className="home-choice"
              onClick={() => p.onDialog("findDialog")}
            >
              Advanced Find…
            </button>
            <button className="home-choice" onClick={() => p.onDialog("goto")}>
              Go To…
            </button>
          </>
        )}
        {p.dialog === "goto" && (
          <form
            onSubmit={(event) => {
              const d = dialogData(event);
              choose(() => p.onAction(`goto:${d.get("page")}`));
            }}
          >
            <label className="field">
              Go to page
              <input
                type="number"
                name="page"
                min="1"
                defaultValue="1"
                required
                autoFocus
              />
            </label>
            <button className="primary-button">Go To</button>
          </form>
        )}
        {p.dialog === "select" && (
          <div className="home-choice-list">
            <button
              className="home-choice"
              onClick={() => choose(() => p.onAction("selectAll"))}
            >
              Select All <kbd>Ctrl A</kbd>
            </button>
            <button
              className="home-choice"
              onClick={() => choose(() => p.onAction("selectSimilar"))}
            >
              Select Text with Similar Formatting
            </button>
            <button
              className="home-choice"
              onClick={() => choose(() => p.onAction("outline"))}
            >
              Selection / document outline…
            </button>
            <button className="home-choice" onClick={() => p.onDialog("goto")}>
              Go To…
            </button>
          </div>
        )}
      </div>
    </HomeSurface>
  );
}
const dropdowns = new Set([
  "paste",
  "fontFamily",
  "fontSize",
  "case",
  "underline",
  "highlight",
  "color",
  "shading",
  "borders",
  "bullets",
  "numbering",
  "multilevel",
  "spacing",
  "justify",
  "effects",
  "styles",
  "select",
  "styleContext",
  "findMenu",
]);
function HomeSurface({
  name,
  title,
  anchor,
  onClose,
  children,
}: {
  name: string;
  title: string;
  anchor: MenuAnchor | null;
  onClose: () => void;
  children: ReactNode;
}) {
  return dropdowns.has(name) ? (
    <HomeMenu label={title} anchor={anchor} onClose={onClose}>
      {children}
    </HomeMenu>
  ) : (
    <Dialog title={title} onClose={onClose}>
      {children}
    </Dialog>
  );
}
