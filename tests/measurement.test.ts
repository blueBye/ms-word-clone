import { afterEach, describe, expect, it, vi } from 'vitest';
import { measureBlocks, measureStoryInset } from '../src/react/measurement.js';

// Controlled browser geometry: these checks do not claim to test font shaping.
function box(top: number, height: number, extra: Record<string, unknown> = {}) {
  return {
    tagName: 'P',
    children: [],
    style: {},
    parentElement: null,
    getBoundingClientRect: () => ({
      top,
      bottom: top + height,
      height,
      width: 200,
    }),
    querySelectorAll: () => [],
    closest: () => null,
    ...extra,
  } as unknown as Element;
}

afterEach(() => vi.unstubAllGlobals());

describe('browser page measurements', () => {
  it('reserves floating header content on every variant, including odd page three', () => {
    vi.stubGlobal(
      'getComputedStyle',
      (element: { style: unknown }) => element.style,
    );
    // The paper can be anywhere in the measurement stack; use paper-relative bounds.
    const paper = box(1000, 1122.533);
    const paragraph = box(1047.267, 18);
    const logo = box(1047.6, 86.467);
    const watermark = box(1000, 800, { closest: () => ({}) });
    const header = box(1047.267, 56, {
      parentElement: paper,
      querySelectorAll: () => [paragraph, logo, watermark],
    });
    expect(measureStoryInset(header, 'header')).toBeCloseTo(100.55, 2);
    expect(measureStoryInset(null, 'header')).toBe(0);
  });

  it('reserves footer textboxes extending upward beyond the footer flow box', () => {
    vi.stubGlobal(
      'getComputedStyle',
      (element: { style: unknown }) => element.style,
    );
    const footer = box(1050, 20, {
      parentElement: box(0, 1122.533),
      querySelectorAll: () => [box(1010, 60)],
    });
    expect(measureStoryInset(footer, 'footer')).toBeCloseTo(84.4, 1);
  });

  it('captures table column widths and cell bounds without counting collapsed margins twice', () => {
    vi.stubGlobal(
      'getComputedStyle',
      (element: { style: unknown }) => element.style,
    );
    const a = box(100, 20, {
      style: { marginTop: '4px', marginBottom: '8px' },
    });
    const b = box(132, 20, {
      style: { marginTop: '12px', marginBottom: '6px' },
    });
    const cell = box(90, 78, {
      children: [a, b],
      style: { paddingTop: '6px', paddingBottom: '6px' },
    });
    const row = box(90, 78, { cells: [cell] });
    const table = box(90, 78, {
      tagName: 'TABLE',
      querySelectorAll: (selector: string) =>
        selector.includes('tbody') ? [row] : [box(0, 0)],
    });
    const [measurement] = measureBlocks(box(0, 0, { children: [table] }));
    expect(measurement).toMatchObject({
      tableWidthPt: 150,
      columnWidthsPt: [150],
      rowHeightsPt: [58.5],
      cells: [
        [
          {
            insetPt: 9,
            blocks: [
              { startPt: 4.5, endPt: 28.5 },
              { startPt: 22.5, endPt: 51 },
            ],
          },
        ],
      ],
    });
  });
});
