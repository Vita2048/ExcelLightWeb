import { expect, it } from 'vitest';
import { drawGesture, hitBorder } from '../../src/grid/border-geometry';
import { GridAxis } from '../../src/grid/grid-geometry';
import { emptySheet } from '../../src/core/workbook';

it('picks the same physical edges from pinned and scrolling cells at zoom', () => {
  const sheet = emptySheet('sheet', 'Sheet');
  const view = {
    rows: new GridAxis(10, 26, [{ index: 3, hidden: true }], 1.5, 1),
    columns: new GridAxis(10, 96, [], 1.5, 1), scrollX: 144, scrollY: 39,
  };
  expect(view.rows.projected(39, 150).slice(0, 2).map(track => [track.index, track.start])).toEqual([[1, 0], [4, 39]]);
  expect(view.columns.projected(144, 400).slice(0, 2).map(track => [track.index, track.start])).toEqual([[1, 0], [3, 144]]);
  const pinned = hitBorder(sheet, view, 55, 38, false)!;
  const scrolled = hitBorder(sheet, view, 175, 40, false)!;
  expect(pinned.segment).toEqual({ edge: 'top', row: 2, column: 1 });
  expect(scrolled.segment).toEqual({ edge: 'top', row: 4, column: 3 });
  const end = hitBorder(sheet, view, 350, 40, false)!;
  expect(drawGesture('draw', scrolled, end)).toEqual({ segments: [
    { edge: 'top', row: 4, column: 3 }, { edge: 'top', row: 4, column: 4 },
  ] });
  expect(drawGesture('grid', scrolled, end)).toEqual({ area: { row: 4, column: 3, rows: 1, columns: 2 } });
  expect(drawGesture('draw', pinned, end)).toEqual({ area: { row: 1, column: 1, rows: 4, columns: 4 } });
});

it('picks merged perimeter segments and only detects diagonals for the eraser', () => {
  const sheet = { ...emptySheet('sheet', 'Sheet'), merges: [{ row: 1, column: 1, rows: 2, columns: 2 }],
    cells: [{ id: 'anchor', row: 1, column: 1, content: { kind: 'blank' } as const,
      style: { borderDetails: { diagonalDown: { style: 'thin' } } } }] };
  const view = { rows: new GridAxis(10, 26, []), columns: new GridAxis(10, 96, []), scrollX: 0, scrollY: 0 };
  expect(hitBorder(sheet, view, 120, 51, false)?.segment).toEqual({ edge: 'top', row: 3, column: 2 });
  expect(hitBorder(sheet, view, 96, 26, false)?.segment).toBeUndefined();
  expect(hitBorder(sheet, view, 96, 26, true)?.segment).toEqual({ edge: 'diagonalDown', row: 1, column: 1 });
});
