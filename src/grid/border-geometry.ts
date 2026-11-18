import { mergeAt } from '../application/border-commands';
import { cellBorders, type BorderMode, type BorderSegment } from '../core/formats/borders';
import { findCell, type CellExtent, type SheetDocument } from '../core/workbook';
import type { GridAxis } from './grid-geometry';

export interface BorderHit { readonly cell: CellExtent; readonly segment?: BorderSegment }
export interface BorderViewport {
  readonly rows: GridAxis;
  readonly columns: GridAxis;
  readonly scrollX: number;
  readonly scrollY: number;
}

export function hitBorder(sheet: SheetDocument, view: BorderViewport, x: number, y: number, diagonals: boolean): BorderHit | undefined {
  const row = view.rows.atScreen(y, view.scrollY);
  const column = view.columns.atScreen(x, view.scrollX);
  if (!row || !column) return;
  const cell = mergeAt(sheet, row.index, column.index) ?? { row: row.index, column: column.index, rows: 1, columns: 1 };
  const left = view.columns.screen(column, view.scrollX) - view.columns.span(cell.column, column.index) + column.size;
  const top = view.rows.screen(row, view.scrollY) - view.rows.span(cell.row, row.index) + row.size;
  const width = view.columns.span(cell.column, cell.column + cell.columns - 1);
  const height = view.rows.span(cell.row, cell.row + cell.rows - 1);
  const distances = [Math.abs(y - top), Math.abs(y - top - height), Math.abs(x - left), Math.abs(x - left - width)];
  const closest = Math.min(...distances);
  if (closest <= Math.min(8, Math.min(width, height) * .3)) {
    const index = distances.indexOf(closest);
    const horizontal = index < 2;
    return { cell, segment: {
      edge: horizontal ? 'top' : 'left', row: index === 1 ? cell.row + cell.rows : horizontal ? cell.row : row.index,
      column: index === 3 ? cell.column + cell.columns : horizontal ? column.index : cell.column,
    } };
  }
  if (diagonals) {
    const border = cellBorders(findCell(sheet, cell)?.style);
    const nx = (x - left) / width, ny = (y - top) / height;
    const tolerance = 6 / Math.min(width, height);
    if (border.diagonalDown && Math.abs(nx - ny) < tolerance) return { cell, segment: { edge: 'diagonalDown', row: cell.row, column: cell.column } };
    if (border.diagonalUp && Math.abs(1 - nx - ny) < tolerance) return { cell, segment: { edge: 'diagonalUp', row: cell.row, column: cell.column } };
  }
  return { cell };
}

export type DrawGesture = { readonly area: CellExtent } | { readonly segments: readonly BorderSegment[] };

export function drawGesture(mode: Exclude<BorderMode, 'erase'>, first: BorderHit, last: BorderHit): DrawGesture {
  const a = first.segment, b = last.segment;
  if (mode === 'draw' && a && b && a.edge === b.edge &&
      (a.edge === 'top' && a.row === b.row || a.edge === 'left' && a.column === b.column)) {
    const horizontal = a.edge === 'top';
    const start = Math.min(horizontal ? a.column : a.row, horizontal ? b.column : b.row);
    const end = Math.max(horizontal ? a.column : a.row, horizontal ? b.column : b.row);
    return { segments: Array.from({ length: end - start + 1 }, (_, offset) => ({
      edge: a.edge, row: horizontal ? a.row : start + offset, column: horizontal ? start + offset : a.column,
    })) };
  }
  const row = Math.min(first.cell.row, last.cell.row), column = Math.min(first.cell.column, last.cell.column);
  return { area: { row, column, rows: Math.max(first.cell.row + first.cell.rows, last.cell.row + last.cell.rows) - row,
    columns: Math.max(first.cell.column + first.cell.columns, last.cell.column + last.cell.columns) - column } };
}
