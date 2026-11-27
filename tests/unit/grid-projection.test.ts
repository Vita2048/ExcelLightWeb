import { describe, expect, it } from 'vitest';
import { RowMask } from '../../src/core/row-mask';
import { emptySheet, ownDto, type SheetDocument } from '../../src/core/workbook';
import { GridAxis } from '../../src/grid/grid-geometry';
import { navigate, parseAddress, select, selectionAddress, selectionStatistics, singleSelection } from '../../src/grid/grid-selection';

const sheet = ownDto<SheetDocument>({
  ...emptySheet('sheet', 'Sheet1'),
  rows: [{ index: 3, hidden: true }, { index: 4, size: 38 }],
  cells: [
    { id: '1', row: 1, column: 1, content: { kind: 'number', value: 2 } },
    { id: '2', row: 2, column: 1, content: { kind: 'text', value: '' } },
    { id: '3', row: 3, column: 1, content: { kind: 'number', value: 100 } },
    { id: '4', row: 4, column: 1, content: { kind: 'formula', source: '=A1*2' }, cache: { kind: 'number', value: 4 } },
    { id: '5', row: 5, column: 1, content: { kind: 'number', value: 8 } },
    { id: '6', row: 7, column: 1, content: { kind: 'number', value: 10 } },
  ],
});

describe('worksheet projection', () => {
  it('maps partial tracks, hidden rows, custom extents and zoom using the same offsets', () => {
    const axis = new GridAxis(6, 26, sheet.rows, 1.5);
    expect(axis.total).toBe(213);
    expect(axis.at(77.9)?.index).toBe(2);
    expect(axis.at(78)?.index).toBe(4);
    expect(axis.at(213)).toBeUndefined();
    expect(axis.visible(40, 40).map(track => track.index)).toEqual([2, 4]);
    expect(axis.step(2, 1)).toBe(4);
    expect(axis.step(3, -1)).toBe(2);
    expect(axis.step(3, 1)).toBe(4);
    expect(axis.reveal(4, 0, 45)).toBe(78);
    expect(axis.span(2, 4)).toBe(96);
  });

  it('jumps to the edge of populated runs, skips hidden rows and stops at the last used coordinate', () => {
    const rows = new GridAxis(1000, 26, sheet.rows);
    const columns = new GridAxis(26, 96, []);
    const move = (row: number, key: string) => navigate(sheet, rows, columns, { row, column: 1 }, key, true, false, 10);
    expect(move(1, 'ArrowDown')).toEqual({ row: 5, column: 1 });
    expect(move(5, 'ArrowDown')).toEqual({ row: 7, column: 1 });
    expect(move(7, 'ArrowDown')).toEqual({ row: 1000, column: 1 });
    expect(move(2, 'End')).toEqual({ row: 7, column: 1 });
  });

  it('counts overlapping coordinates once and excludes hidden cells while using formula caches', () => {
    const areas = [{ top: 1, bottom: 2, left: 1, right: 1 }, { top: 2, bottom: 4, left: 1, right: 1 }];
    expect(selectionStatistics(sheet, areas)).toEqual({ count: 3, sum: 6, average: 3 });
    const larger = [{ top: 1, bottom: 4, left: 1, right: 2 }, areas[0]!];
    expect(selectionStatistics(sheet, larger)).toEqual({ count: 3, sum: 6, average: 3 });
    expect(selectionStatistics(sheet, [{ top: 4, bottom: 4, left: 1, right: 1 }])).toEqual({ count: 1, sum: 4, average: 4 });
    expect(selectionStatistics(sheet, [{ top: 4, bottom: 4, left: 1, right: 1 }], RowMask.hidden(1, 10, [4]))).toEqual({
      count: 0, sum: 0, average: undefined,
    });
  });

  it('preserves the active range if an earlier area is toggled and bounds the area count', () => {
    const first = select(singleSelection(), { row: 2, column: 2 }, 'cells', true);
    const second = select(first, { row: 3, column: 3 }, 'cells', false, true);
    expect(selectionAddress(second)).toBe('A1:B2, C3');
    expect(selectionAddress(select(second, { row: 2, column: 1 }, 'cells', false, true))).toBe('C3');
    expect(selectionAddress(select(second, { row: 3, column: 3 }, 'cells', false, true))).toBe('A1:B2');
    let many = singleSelection();
    for (let row = 2; row < 70; row++) many = select(many, { row, column: 1 }, 'cells', false, true);
    expect(many.additional).toHaveLength(63);
  });

  it('accepts addresses through XFD1048576 and rejects names, ranges and coordinates past the grid', () => {
    expect(parseAddress(' $z$1000 ')).toEqual({ row: 1000, column: 26 });
    expect(parseAddress('XFD1048576')).toEqual({ row: 1_048_576, column: 16_384 });
    expect(parseAddress('AA1')).toEqual({ row: 1, column: 27 });
    for (const value of ['A0', 'A1048577', 'XFE1', 'A1:B2', 'Sheet1!A1', 'C2\nD3']) {
      expect(parseAddress(value)).toBeUndefined();
    }
  });

  it('jumps Ctrl+Down on an empty million-row axis to the last row', () => {
    const rows = new GridAxis(1_048_576, 26, []);
    const columns = new GridAxis(16_384, 96, []);
    const blank = emptySheet('sheet', 'Sheet1');
    const started = performance.now();
    const point = navigate(blank, rows, columns, { row: 1, column: 1 }, 'ArrowDown', true, false, 10);
    expect(performance.now() - started).toBeLessThan(50);
    expect(point).toEqual({ row: 1_048_576, column: 1 });
  });
});
