import type { DocxStyle } from "@local/docx-ir-engine";
import { textEffectsStyle } from "@local/docx-ir-engine/react";

export function StylePreview({
  style,
  compact = false,
}: {
  style: DocxStyle;
  compact?: boolean;
}) {
  const r = style.run;
  return (
    <span
      className="style-preview"
      style={{
        ...textEffectsStyle(r),
        fontFamily: r.fontFamily
          ? `${JSON.stringify(r.fontFamily)},sans-serif`
          : undefined,
        fontWeight: r.bold ? 700 : 400,
        fontStyle: r.italic ? "italic" : "normal",
        color: r.color ?? "inherit",
        fontSize: compact
          ? `${Math.min(24, 14 + (r.fontSizePt ?? 12) * 0.3)}px`
          : `${Math.min(34, r.fontSizePt ?? 12)}pt`,
        fontVariantCaps: r.smallCaps ? "small-caps" : undefined,
        textTransform: r.caps ? "uppercase" : undefined,
        textDecoration:
          r.underline && r.underline !== "none" ? "underline" : undefined,
      }}
    >
      {style.name}
    </span>
  );
}
