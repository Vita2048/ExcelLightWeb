import type { FigurePlacement } from '../core/workbook';
import { clamp, type GridAxis } from './grid-geometry';

export interface FigureRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export function figureLogicalRect(placement: FigurePlacement, rows: GridAxis, columns: GridAxis,
  scale: number): FigureRect | undefined {
  const row = rows.trackAtOrAfter(placement.row);
  const column = columns.trackAtOrAfter(placement.column);
  if (!row || !column) return undefined;
  const spanWidth = columns.span(placement.column, placement.column + placement.columns - 1) / scale;
  const spanHeight = rows.span(placement.row, placement.row + placement.rows - 1) / scale;
  return {
    x: column.start / scale + (placement.offsetX ?? 0),
    y: row.start / scale + (placement.offsetY ?? 0),
    width: placement.width && placement.width > 0 ? placement.width : spanWidth,
    height: placement.height && placement.height > 0 ? placement.height : spanHeight,
  };
}

export function figureScreenRect(placement: FigurePlacement, rows: GridAxis, columns: GridAxis, scale: number,
  scrollX: number, scrollY: number, headerLeft: number, headerTop: number): FigureRect | undefined {
  const logical = figureLogicalRect(placement, rows, columns, scale);
  const row = rows.trackAtOrAfter(placement.row);
  const column = columns.trackAtOrAfter(placement.column);
  if (!logical || !row || !column) return undefined;
  return {
    x: headerLeft + columns.screen(column, scrollX) + (placement.offsetX ?? 0) * scale,
    y: headerTop + rows.screen(row, scrollY) + (placement.offsetY ?? 0) * scale,
    width: logical.width * scale,
    height: logical.height * scale,
  };
}

export function logicalPoint(position: number, header: number, axis: GridAxis, scroll: number,
  scale: number): number {
  const relative = Math.max(0, position - header);
  const track = axis.atScreen(relative, scroll) ?? axis.last;
  if (!track) return 0;
  const start = header + axis.screen(track, scroll);
  return track.start / scale + clamp((position - start) / scale, 0, track.size / scale);
}

export function figurePlacementFromRect(id: string, sheetId: string, rect: FigureRect,
  rows: GridAxis, columns: GridAxis, scale: number): FigurePlacement {
  const maxWidth = columns.total / scale;
  const maxHeight = rows.total / scale;
  const width = Math.min(maxWidth, Math.max(80, rect.width));
  const height = Math.min(maxHeight, Math.max(60, rect.height));
  const x = clamp(rect.x, 0, Math.max(0, maxWidth - width));
  const y = clamp(rect.y, 0, Math.max(0, maxHeight - height));
  const firstColumn = columns.at(x * scale);
  const firstRow = rows.at(y * scale);
  const lastColumn = columns.at(Math.max(x, x + width - .001) * scale);
  const lastRow = rows.at(Math.max(y, y + height - .001) * scale);
  if (!firstColumn || !firstRow || !lastColumn || !lastRow) throw new Error('Figure placement is outside the visible worksheet.');
  return {
    id, sheetId, row: firstRow.index, column: firstColumn.index,
    rows: Math.max(1, lastRow.index - firstRow.index + 1),
    columns: Math.max(1, lastColumn.index - firstColumn.index + 1),
    offsetX: x - firstColumn.start / scale,
    offsetY: y - firstRow.start / scale,
    width: Math.max(80, Math.min(width, maxWidth - x)),
    height: Math.max(60, Math.min(height, maxHeight - y)),
  };
}
