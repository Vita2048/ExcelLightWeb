import type { Axis } from '../core/axis-edit';
import type { AxisDimension, SheetDocument } from '../core/workbook';
import { gridMetrics, clamp } from './grid-geometry';
import { cellDisplay } from './grid-selection';

export function fitDimensions(sheet: SheetDocument, axis: Axis, indexes: readonly number[], showFormulas = false,
  separator: ',' | ';' = ','): AxisDimension[] {
  const context = document.createElement('canvas').getContext('2d')!;
  const font = getComputedStyle(document.documentElement).fontFamily;
  const sizes = new Map(indexes.map(index => [index, 0]));
  const columns = new Map(sheet.columns.map(column => [column.index, column]));
  const rows = new Map(sheet.rows.map(row => [row.index, row]));
  for (const cell of sheet.cells) {
    const index = cell[axis];
    const previous = sizes.get(index);
    if (previous === undefined) continue;
    const value = cellDisplay(cell, showFormulas, separator);
    if (!value) continue;
    const merge = sheet.merges.find(merge => merge.row === cell.row && merge.column === cell.column);
    if (axis === 'column' && merge && merge.columns > 1) continue;
    const size = (cell.style?.fontSize ?? 10.5) * 4 / 3;
    context.font = (cell.style?.italic ? 'italic ' : '') + (cell.style?.bold ? '600 ' : '') +
      size + 'px ' + (cell.style?.fontFamily ? JSON.stringify(cell.style.fontFamily) + ', ' : '') + font;
    const indent = (cell.style?.indent ?? 0) * size * 1.5;
    const lines = value.split(/\r\n|\r|\n/u);
    let needed: number;
    if (axis === 'column') {
      needed = Math.max(...lines.map(line => context.measureText(line).width)) + 10 + indent;
    } else {
      let width = 0;
      for (let column = cell.column; column < cell.column + (merge?.columns ?? 1); column++) {
        const track = columns.get(column);
        if (!track?.hidden) width += track?.size ?? gridMetrics.column;
      }
      let lineCount = 0;
      for (const line of lines) {
        lineCount += cell.style?.wrap ? Math.max(1, Math.ceil(context.measureText(line).width / Math.max(1, width - 10 - indent))) : 1;
      }
      needed = size * 1.25 * lineCount + 6;
      for (let row = cell.row + 1; row < cell.row + (merge?.rows ?? 1); row++) {
        const track = rows.get(row);
        if (!track?.hidden) needed -= track?.size ?? gridMetrics.row;
      }
    }
    sizes.set(index, Math.max(previous, needed));
  }
  return [...sizes].map(([index, size]) => ({ index, size: clamp(Math.ceil(size || gridMetrics[axis]), axis === 'row' ? 16 : 12, 4096) }));
}
