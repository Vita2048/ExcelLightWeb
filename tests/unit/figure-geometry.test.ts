import { describe, expect, it } from 'vitest';
import { emptySheet } from '../../src/core/workbook';
import { figureCanvasForPlacement } from '../../src/python/figure-rendering';
import { figureLogicalRect, figurePlacementFromRect, figureScreenRect } from '../../src/grid/figure-geometry';
import { GridAxis } from '../../src/grid/grid-geometry';

describe('Python Figure placement geometry', () => {
  it('maps anchored offsets and dimensions through zoom, hidden tracks and scrolling', () => {
    const scale = 1.5;
    const rows = new GridAxis(12, 26, [{ index: 2, size: 40 }, { index: 3, hidden: true }], scale, 1);
    const columns = new GridAxis(10, 96, [{ index: 2, size: 150 }, { index: 3, hidden: true }], scale, 1);
    const placement = { id: 'plot', sheetId: 'sheet', row: 2, column: 2, rows: 3, columns: 3,
      offsetX: 7, offsetY: 4, width: 260, height: 120 };
    expect(figureLogicalRect(placement, rows, columns, scale)).toEqual({ x: 103, y: 30, width: 260, height: 120 });
    expect(figureScreenRect(placement, rows, columns, scale, 30, 20, 44, 26)).toEqual({
      x: 44 + 96 * scale - 30 + 7 * scale,
      y: 26 + 26 * scale - 20 + 4 * scale,
      width: 390, height: 180,
    });
  });

  it('round trips a dragged rectangle into stable one-based row and column geometry', () => {
    const rows = new GridAxis(1000, 26, [], 1.25);
    const columns = new GridAxis(26, 96, [], 1.25);
    const placement = figurePlacementFromRect('plot', 'sheet', { x: 135, y: 71, width: 360, height: 180 }, rows, columns, 1.25);
    expect(placement).toMatchObject({ row: 3, column: 2, offsetX: 39, offsetY: 19, width: 360, height: 180 });
    expect(figureLogicalRect(placement, rows, columns, 1.25)).toEqual({ x: 135, y: 71, width: 360, height: 180 });
  });

  it('bounds render sizes while preserving independent placement dimensions', () => {
    const sheet = emptySheet('sheet', 'Sheet');
    const source = { id: 'source', sheetId: 'sheet', row: 1, column: 1, rows: 5, columns: 5,
      width: 480, height: 240 };
    const second = { ...source, id: 'second', width: 300, height: 180 };
    expect(figureCanvasForPlacement(source, sheet, 1.5)).toEqual({ width: 720, height: 360, dpi: 144 });
    expect(figureCanvasForPlacement(second, sheet, 1.5)).toEqual({ width: 450, height: 270, dpi: 144 });
    const bounded = figureCanvasForPlacement({ ...source, width: 40_000, height: 8_000 }, sheet, 1.5);
    expect(bounded.width).toBeLessThanOrEqual(8192);
    expect(bounded.height).toBeLessThanOrEqual(8192);
    expect(bounded.width * bounded.height).toBeLessThanOrEqual(16_000_000);
    expect(bounded.dpi).toBeLessThanOrEqual(144);
  });
});
