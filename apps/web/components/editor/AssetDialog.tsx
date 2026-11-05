import { useState } from "react";
import type { DrawingInline, EditCommand } from "@local/docx-ir-engine/core";
import { ModelView } from "@local/docx-ir-engine/react";
import { Dialog } from "./Controls";
import { objectPreview } from "../../lib/object-preview";
export function AssetDialog({
  mode,
  id,
  object,
  execute,
  onClose,
}: {
  mode: "model3d" | "object" | "spreadsheet";
  id?: string;
  object?: { id: string; index: number; drawing: DrawingInline };
  execute: (commands: EditCommand[]) => Promise<boolean>;
  onClose: () => void;
}) {
  const [asset, setAsset] = useState<DrawingInline | null>(
      mode === "spreadsheet"
        ? {
            ...(object?.drawing.spreadsheet
              ? object.drawing
              : {
                  type: "drawing" as const,
                  widthPt: 360,
                  heightPt: 180,
                  alt: "Worksheet.xlsx",
                }),
            spreadsheet: object?.drawing.spreadsheet ?? {
              cells: [
                ["Item", "Quantity", "Price"],
                ["First item", "1", "10"],
                ["Second item", "2", "20"],
              ],
            },
          }
        : null,
    ),
    [error, setError] = useState("");
  return (
    <Dialog
      title={
        mode === "model3d"
          ? "Insert 3D Model"
          : mode === "spreadsheet"
            ? "Excel Spreadsheet"
            : "Insert Object"
      }
      onClose={onClose}
    >
      <form
        onSubmit={async (event) => {
          event.preventDefault();
          if (!asset) return;
          if (
            await execute(
              object && mode === "spreadsheet" && object.drawing.spreadsheet
                ? [
                    {
                      type: "drawing",
                      id: object.id,
                      index: object.index,
                      drawing: objectPreview(asset),
                    },
                  ]
                : [
                    {
                      type: "insertDrawing",
                      id,
                      drawing: asset.model3d ? asset : objectPreview(asset),
                    },
                  ],
            )
          )
            onClose();
        }}
      >
        {mode !== "spreadsheet" && (
          <label className="field">
            {mode === "model3d" ? "GLB model" : "File to embed"}
            <input
              type="file"
              accept={mode === "model3d" ? ".glb,model/gltf-binary" : undefined}
              onChange={async (event) => {
                const file = event.target.files?.[0];
                if (!file) return;
                setError("");
                try {
                  if (file.size > 20_000_000)
                    throw new Error("Choose a file smaller than 20 MB.");
                  const dataUrl = await new Promise<string>(
                    (resolve, reject) => {
                      const reader = new FileReader();
                      reader.onload = () => resolve(String(reader.result));
                      reader.onerror = () =>
                        reject(new Error("Could not read file."));
                      reader.readAsDataURL(
                        new File([file], file.name, {
                          type:
                            mode === "model3d"
                              ? "model/gltf-binary"
                              : file.type || "application/octet-stream",
                        }),
                      );
                    },
                  );
                  setAsset({
                    type: "drawing",
                    widthPt: mode === "model3d" ? 288 : 216,
                    heightPt: mode === "model3d" ? 216 : 72,
                    alt: file.name,
                    ...(mode === "model3d"
                      ? { model3d: { dataUrl } }
                      : {
                          attachment: {
                            dataUrl,
                            name: file.name,
                            mime: file.type || "application/octet-stream",
                            progId: file.name.endsWith(".xlsx")
                              ? "Excel.Sheet.12"
                              : "Package",
                          },
                        }),
                  });
                } catch (error) {
                  setError(
                    error instanceof Error
                      ? error.message
                      : "Could not load file.",
                  );
                }
              }}
            />
          </label>
        )}
        {asset?.model3d && (
          <div style={{ height: 280 }}>
            <ModelView
              model={asset.model3d}
              onReady={(poster, size, center) =>
                setAsset((current) =>
                  current?.model3d
                    ? {
                        ...current,
                        model3d: { ...current.model3d, poster, size, center },
                      }
                    : current,
                )
              }
            />
          </div>
        )}
        {asset?.spreadsheet && (
          <>
            <div className="worksheet-grid">
              <table>
                <tbody>
                  {asset.spreadsheet.cells.map((row, r) => (
                    <tr key={r}>
                      {row.map((value, c) => (
                        <td key={c}>
                          <input
                            aria-label={`Row ${r + 1}, column ${c + 1}`}
                            value={value}
                            onChange={(event) =>
                              setAsset((current) => {
                                if (!current?.spreadsheet) return current;
                                const cells = current.spreadsheet.cells.map(
                                  (row) => row.slice(),
                                );
                                cells[r][c] = event.target.value;
                                return { ...current, spreadsheet: { cells } };
                              })
                            }
                          />
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="field-row">
              <button
                type="button"
                className="secondary-button"
                onClick={() =>
                  setAsset((current) =>
                    current?.spreadsheet
                      ? {
                          ...current,
                          spreadsheet: {
                            cells: [
                              ...current.spreadsheet.cells,
                              current.spreadsheet.cells[0].map(() => ""),
                            ],
                          },
                        }
                      : current,
                  )
                }
              >
                Add row
              </button>
              <button
                type="button"
                className="secondary-button"
                onClick={() =>
                  setAsset((current) =>
                    current?.spreadsheet
                      ? {
                          ...current,
                          spreadsheet: {
                            cells: current.spreadsheet.cells.map((row) => [
                              ...row,
                              "",
                            ]),
                          },
                        }
                      : current,
                  )
                }
              >
                Add column
              </button>
            </div>
          </>
        )}
        {error && <p role="alert">{error}</p>}
        <div className="dialog-actions">
          <button type="button" className="secondary-button" onClick={onClose}>
            Cancel
          </button>
          <button className="primary-button" disabled={!asset}>
            Insert
          </button>
        </div>
      </form>
    </Dialog>
  );
}
