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
) {
  const gap = 2,
    edge = 8;
  const below = Math.max(0, viewport.height - anchor.bottom - gap - edge);
  const above = Math.max(0, anchor.top - gap - edge);
  const flip = menu.height > below && above > below;
  const maxHeight = flip ? above : below;
  const top = flip
    ? anchor.top - gap - Math.min(menu.height, maxHeight)
    : anchor.bottom + gap;
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
