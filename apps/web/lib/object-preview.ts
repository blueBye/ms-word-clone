import type { DrawingInline } from "@local/docx-ir-engine/core";
/** A bitmap preview is required by Word's native OLE frame; the embedded file remains editable. */
export function objectPreview(d: DrawingInline): DrawingInline {
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(d.widthPt * 2);
  canvas.height = Math.round(d.heightPt * 2);
  const ctx = canvas.getContext("2d")!;
  ctx.scale(2, 2);
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, d.widthPt, d.heightPt);
  ctx.font = "11px Segoe UI, sans-serif";
  ctx.textBaseline = "middle";
  if (d.spreadsheet) {
    const rows = d.spreadsheet.cells,
      columns = Math.max(1, ...rows.map((r) => r.length)),
      w = d.widthPt / columns,
      h = Math.min(24, d.heightPt / rows.length);
    rows.forEach((row, r) => {
      for (let c = 0; c < columns; c++) {
        ctx.fillStyle = r === 0 ? "#E5EDF5" : "#FFFFFF";
        ctx.fillRect(c * w, r * h, w, h);
        ctx.strokeStyle = "#B8C2CF";
        ctx.strokeRect(c * w, r * h, w, h);
        ctx.save();
        ctx.beginPath();
        ctx.rect(c * w + 4, r * h, w - 8, h);
        ctx.clip();
        ctx.fillStyle = "#222222";
        ctx.fillText(row[c] ?? "", c * w + 4, r * h + h / 2);
        ctx.restore();
      }
    });
  } else {
    ctx.fillStyle = "#EEF2F6";
    ctx.fillRect(0, 0, d.widthPt, d.heightPt);
    ctx.fillStyle = "#294F76";
    ctx.textAlign = "center";
    ctx.fillText(
      d.attachment?.name ?? d.alt,
      d.widthPt / 2,
      d.heightPt / 2,
      d.widthPt - 16,
    );
  }
  return { ...d, image: { dataUrl: canvas.toDataURL("image/png") } };
}
