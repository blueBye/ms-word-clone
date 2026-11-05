export interface MenuBounds {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

/** Prefer below the button; flip only when there is more room above it. */
export function menuPosition(
  anchor: MenuBounds,
  menu: { width: number; height: number },
  viewport: { width: number; height: number },
  placement: "below" | "overlap" = "below",
) {
  const gap = 2,
    edge = 8;
  const preferredTop =
    placement === "overlap" ? anchor.top : anchor.bottom + gap;
  const below = Math.max(0, viewport.height - preferredTop - edge);
  const above = Math.max(0, anchor.top - gap - edge);
  const flip = menu.height > below && above > below;
  const maxHeight = flip ? above : below;
  const top = flip
    ? anchor.top - gap - Math.min(menu.height, maxHeight)
    : preferredTop;
  return {
    left: Math.max(
      edge,
      Math.min(anchor.left, viewport.width - menu.width - edge),
    ),
    top: Math.max(edge, top),
    maxHeight,
    clipped: menu.height > maxHeight,
  };
}
