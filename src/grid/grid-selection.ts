import { expandMerges } from '../application/editing-commands';
import { mergeAt } from '../application/border-commands';
import { filteredRows } from '../application/data-commands';
import { columnIndex } from '../core/formula-references';
import type { RowMask } from '../core/row-mask';
import { cellKey, findCell, worksheetViewport as viewport, type CellDocument, type Coordinate, type SheetDocument } from '../core/workbook';
import { formatCellValue, type FormattedCell } from '../core/formats/cell-format';
import { presentFormula } from '../features/formulas/formula-presentation';
import { address, type GridAxis } from './grid-geometry';

export type AreaKind = 'cells' | 'rows' | 'columns' | 'sheet';
export interface Area {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}
export interface GridSelection {
  readonly anchor: Coordinate;
  readonly end: Coordinate;
  readonly kind: AreaKind;
  readonly additional: readonly Area[];
  readonly mergedArea?: Area | undefined;
}

export function singleSelection(point: Coordinate = { row: 1, column: 1 }): GridSelection {
  return { anchor: point, end: point, kind: 'cells', additional: [] };
}

export function primaryArea({ anchor, end, kind, mergedArea }: GridSelection): Area {
  if (mergedArea) return mergedArea;
  const allRows = kind === 'columns' || kind === 'sheet';
  const allColumns = kind === 'rows' || kind === 'sheet';
  return {
    top: allRows ? 1 : Math.min(anchor.row, end.row),
    left: allColumns ? 1 : Math.min(anchor.column, end.column),
    bottom: allRows ? viewport.rows : Math.max(anchor.row, end.row),
    right: allColumns ? viewport.columns : Math.max(anchor.column, end.column),
  };
}

export function includes(area: Area, { row, column }: Coordinate): boolean {
  return row >= area.top && row <= area.bottom && column >= area.left && column <= area.right;
}

export function selectionAreas(selection: GridSelection): readonly Area[] {
  return [...selection.additional, primaryArea(selection)];
}

export function select(
  current: GridSelection, point: Coordinate, kind: AreaKind, extend = false, add = false,
): GridSelection {
  if (extend && current.kind === kind) return { ...current, end: point, mergedArea: undefined };
  if (!add || kind !== 'cells') return { ...singleSelection(point), kind };
  const prior = primaryArea(current);
  const index = current.additional.findIndex(area => includes(area, point));
  if (index >= 0) return { ...current, additional: current.additional.filter((_, i) => i !== index) };
  if (includes(prior, point)) {
    const previous = current.additional.at(-1);
    if (!previous) return current;
    return {
      anchor: { row: previous.top, column: previous.left },
      end: { row: previous.bottom, column: previous.right },
      kind: 'cells', additional: current.additional.slice(0, -1),
    };
  }
  return { ...singleSelection(point), additional: [...current.additional.slice(-62), prior] };
}

export function mergeSelection(sheet: SheetDocument, selection: GridSelection): GridSelection {
  if (!sheet.merges.length || selection.kind !== 'cells') return selection;
  function expand(area: Area): Area {
    const range = expandMerges({ sheetId: sheet.id, row: area.top, column: area.left,
      rows: area.bottom - area.top + 1, columns: area.right - area.left + 1 }, sheet);
    return { top: range.row, left: range.column, bottom: range.row + range.rows - 1, right: range.column + range.columns - 1 };
  }
  const merged = mergeAt(sheet, selection.anchor.row, selection.anchor.column);
  const anchor = merged ? { row: merged.row, column: merged.column } : selection.anchor;
  return { ...selection, anchor, mergedArea: expand(primaryArea(selection)), additional: selection.additional.map(expand) };
}

export function selectionAddress(selection: GridSelection): string {
  return selectionAreas(selection).map(({ top, left, bottom, right }) => {
    const first = address({ row: top, column: left });
    return top === bottom && left === right ? first : first + ':' + address({ row: bottom, column: right });
  }).join(', ');
}

export function parseAddress(text: string): Coordinate | undefined {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/iu.exec(text.trim());
  if (!match) return undefined;
  const row = Number(match[2]);
  const column = columnIndex(match[1]!);
  if (!Number.isSafeInteger(row) || row > viewport.rows || column > viewport.columns) return undefined;
  return { row, column };
}

export function cellDisplay(cell: CellDocument | undefined, source = false, separator: ',' | ';' = ',', numberFormat?: string): string {
  return presentCell(cell, source, separator, numberFormat).text;
}

export function presentCell(cell: CellDocument | undefined, source: boolean, separator: ',' | ';', numberFormat?: string): FormattedCell {
  if (cell?.content.kind === 'formula' && (source || !cell.cache)) {
    return { text: presentFormula(cell.content.source, ',', separator) };
  }
  if (cell?.content.kind === 'python' && source) return { text: '=PY()' };
  return formatCellValue(cell, numberFormat);
}

export function selectionStatistics(sheet: SheetDocument, areas: readonly Area[], filtered: RowMask = filteredRows(sheet)) {
  let hiddenRows: Set<number> | undefined;
  for (const row of sheet.rows) {
    if (row.hidden) (hiddenRows ??= new Set()).add(row.index);
  }
  let hiddenColumns: Set<number> | undefined;
  for (const column of sheet.columns) {
    if (column.hidden) (hiddenColumns ??= new Set()).add(column.index);
  }
  let count = 0;
  let numeric = 0;
  let sum = 0;
  function include(cell: CellDocument | undefined) {
    if (!cell || cell.content.kind === 'blank' || hiddenRows?.has(cell.row) || hiddenColumns?.has(cell.column) ||
        filtered.has(cell.row)) return;
    count++;
    const value = cell.content.kind === 'formula' || cell.content.kind === 'python' ? cell.cache : cell.content;
    if (value?.kind === 'number') {
      numeric++;
      sum += value.value;
    }
  }
  if (areas.length === 1 && areas[0]!.top === areas[0]!.bottom && areas[0]!.left === areas[0]!.right) {
    include(findCell(sheet, { row: areas[0]!.top, column: areas[0]!.left }));
    return { count, sum, average: numeric ? sum / numeric : undefined };
  }
  const extentSize = areas.reduce((size, area) => size + (area.bottom - area.top + 1) * (area.right - area.left + 1), 0);
  if (extentSize < sheet.cells.length) {
    const visited = new Set<number>();
    for (const area of areas) {
      for (let row = area.top; row <= area.bottom; row++) {
        for (let column = area.left; column <= area.right; column++) {
          const key = cellKey({ row, column });
          if (visited.has(key)) continue;
          visited.add(key);
          include(findCell(sheet, { row, column }));
        }
      }
    }
  } else {
    for (const cell of sheet.cells) {
      if (areas.some(area => includes(area, cell))) include(cell);
    }
  }
  return { count, sum, average: numeric ? sum / numeric : undefined };
}

function lowerBound(indexes: readonly number[], target: number): number {
  let low = 0;
  let high = indexes.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (indexes[middle]! < target) low = middle + 1;
    else high = middle;
  }
  return low;
}

/** Ctrl+Arrow: one scan of populated cells on the line, then the axis edge. */
function jumpOnLine(sheet: SheetDocument, axis: GridAxis, vertical: boolean, fixed: number,
  start: number, direction: number): number {
  const neighbor = axis.step(start, direction);
  if (neighbor === start) return start;
  const occupied: number[] = [];
  for (const cell of sheet.cells) {
    if (cell.content.kind === 'blank') continue;
    const index = vertical ? cell.row : cell.column;
    const other = vertical ? cell.column : cell.row;
    if (other !== fixed || !axis.track(index)) continue;
    occupied.push(index);
  }
  occupied.sort((left, right) => left - right);
  let write = 0;
  for (const index of occupied) {
    if (write === 0 || occupied[write - 1] !== index) occupied[write++] = index;
  }
  occupied.length = write;
  const contains = (index: number) => {
    const at = lowerBound(occupied, index);
    return at < occupied.length && occupied[at] === index;
  };
  const edge = direction > 0 ? axis.last?.index ?? start : axis.first?.index ?? start;
  if (contains(start) && contains(neighbor)) {
    let at = lowerBound(occupied, start);
    let current = start;
    while (true) {
      const nextAt = at + direction;
      if (nextAt < 0 || nextAt >= occupied.length) return current;
      const next = occupied[nextAt]!;
      const adjacent = direction > 0
        ? axis.trackAtOrAfter(current + 1)?.index === next
        : axis.trackAtOrAfter(next + 1)?.index === current;
      if (!adjacent) return current;
      current = next;
      at = nextAt;
    }
  }
  if (direction > 0) {
    const at = lowerBound(occupied, start + 1);
    return at < occupied.length ? occupied[at]! : edge;
  }
  const at = lowerBound(occupied, start) - 1;
  return at >= 0 ? occupied[at]! : edge;
}

export function navigate(
  sheet: SheetDocument, rows: GridAxis, columns: GridAxis, point: Coordinate,
  key: string, control: boolean, shift: boolean, pageRows: number,
): Coordinate | undefined {
  const { row, column } = point;
  const merge = mergeAt(sheet, row, column);
  const lastRow = merge ? merge.row + merge.rows - 1 : row;
  const lastColumn = merge ? merge.column + merge.columns - 1 : column;
  switch (key) {
    case 'Home': return { row: control ? rows.first?.index ?? row : row, column: columns.first?.index ?? column };
    case 'End': {
      let lastRow = control ? 1 : row;
      let lastColumn = 1;
      for (const cell of sheet.cells) {
        if (cell.content.kind === 'blank' || !rows.track(cell.row) || !columns.track(cell.column)) continue;
        if (control) lastRow = Math.max(lastRow, cell.row);
        if (control || cell.row === row) lastColumn = Math.max(lastColumn, cell.column);
      }
      return {
        row: rows.track(lastRow)?.index ?? rows.first?.index ?? row,
        column: columns.track(lastColumn)?.index ?? columns.first?.index ?? column,
      };
    }
    case 'PageUp': return { row: rows.step(row, -pageRows), column };
    case 'PageDown': return { row: rows.step(lastRow, pageRows), column };
    case 'Enter': return { row: rows.step(shift ? merge?.row ?? row : lastRow, shift ? -1 : 1), column };
    case 'Tab': return { row, column: columns.step(shift ? merge?.column ?? column : lastColumn, shift ? -1 : 1) };
    case 'ArrowLeft':
    case 'ArrowRight':
    case 'ArrowUp':
    case 'ArrowDown': {
      const vertical = key === 'ArrowUp' || key === 'ArrowDown';
      const direction = key === 'ArrowUp' || key === 'ArrowLeft' ? -1 : 1;
      const axis = vertical ? rows : columns;
      const start = vertical ? direction > 0 ? lastRow : merge?.row ?? row :
        direction > 0 ? lastColumn : merge?.column ?? column;
      const next = control ? jumpOnLine(sheet, axis, vertical, vertical ? column : row, start, direction) :
        axis.step(start, direction);
      return vertical ? { row: next, column } : { row, column: next };
    }
    default: return undefined;
  }
}
