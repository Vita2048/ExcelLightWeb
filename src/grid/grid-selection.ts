import { cellKey, findCell, worksheetViewport as viewport, type CellDocument, type Coordinate, type SheetDocument } from '../core/workbook';
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
}

export function singleSelection(point: Coordinate = { row: 1, column: 1 }): GridSelection {
  return { anchor: point, end: point, kind: 'cells', additional: [] };
}

export function primaryArea({ anchor, end, kind }: GridSelection): Area {
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
  if (extend && current.kind === kind) return { ...current, end: point };
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

export function selectionAddress(selection: GridSelection): string {
  return selectionAreas(selection).map(({ top, left, bottom, right }) => {
    const first = address({ row: top, column: left });
    return top === bottom && left === right ? first : first + ':' + address({ row: bottom, column: right });
  }).join(', ');
}

export function parseAddress(text: string): Coordinate | undefined {
  const match = /^\$?([A-Z])\$?([1-9][0-9]{0,3})$/iu.exec(text.trim());
  if (!match) return undefined;
  const row = Number(match[2]);
  return row <= viewport.rows ? { row, column: match[1]!.toUpperCase().charCodeAt(0) - 64 } : undefined;
}

export function cellDisplay(cell: CellDocument | undefined, source = false): string {
  if (!cell) return '';
  const { content, cache } = cell;
  if (content.kind === 'python' && source) return '=PY()';
  if (content.kind === 'formula' && source) return content.source;
  const value = cache && !source ? cache : content;
  switch (value.kind) {
    case 'blank': return '';
    case 'formula': return value.source;
    case 'python': return '=PY()';
    case 'boolean': return value.value ? 'TRUE' : 'FALSE';
    default: return String(value.value);
  }
}

export function selectionStatistics(sheet: SheetDocument, areas: readonly Area[]) {
  const hiddenRows = new Set(sheet.rows.filter(row => row.hidden).map(row => row.index));
  const hiddenColumns = new Set(sheet.columns.filter(column => column.hidden).map(column => column.index));
  let count = 0;
  let numeric = 0;
  let sum = 0;
  function include(cell: CellDocument | undefined) {
    if (!cell || cell.content.kind === 'blank' || hiddenRows.has(cell.row) || hiddenColumns.has(cell.column)) return;
    count++;
    const value = cell.content.kind === 'formula' || cell.content.kind === 'python' ? cell.cache : cell.content;
    if (value?.kind === 'number') {
      numeric++;
      sum += value.value;
    }
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

export function navigate(
  sheet: SheetDocument, rows: GridAxis, columns: GridAxis, point: Coordinate,
  key: string, control: boolean, shift: boolean, pageRows: number,
): Coordinate | undefined {
  const { row, column } = point;
  switch (key) {
    case 'Home': return { row: control ? rows.tracks[0]?.index ?? row : row, column: columns.tracks[0]?.index ?? column };
    case 'End': {
      let lastRow = control ? 1 : row;
      let lastColumn = 1;
      for (const cell of sheet.cells) {
        if (cell.content.kind === 'blank' || !rows.track(cell.row) || !columns.track(cell.column)) continue;
        if (control) lastRow = Math.max(lastRow, cell.row);
        if (control || cell.row === row) lastColumn = Math.max(lastColumn, cell.column);
      }
      return {
        row: rows.track(lastRow)?.index ?? rows.tracks[0]?.index ?? row,
        column: columns.track(lastColumn)?.index ?? columns.tracks[0]?.index ?? column,
      };
    }
    case 'PageUp': return { row: rows.step(row, -pageRows), column };
    case 'PageDown': return { row: rows.step(row, pageRows), column };
    case 'Enter': return { row: rows.step(row, shift ? -1 : 1), column };
    case 'Tab': return { row, column: columns.step(column, shift ? -1 : 1) };
    case 'ArrowLeft':
    case 'ArrowRight':
    case 'ArrowUp':
    case 'ArrowDown': {
      const vertical = key === 'ArrowUp' || key === 'ArrowDown';
      const direction = key === 'ArrowUp' || key === 'ArrowLeft' ? -1 : 1;
      const axis = vertical ? rows : columns;
      const start = vertical ? row : column;
      const occupies = (index: number) => {
        const cell = findCell(sheet, vertical ? { row: index, column } : { row, column: index });
        return cell !== undefined && cell.content.kind !== 'blank';
      };
      let next = axis.step(start, direction);
      if (control && next !== start) {
        const contiguous = occupies(start) && occupies(next);
        let current = start;
        while (next !== current) {
          if (contiguous && !occupies(next)) { next = current; break; }
          if (!contiguous && occupies(next)) break;
          current = next;
          next = axis.step(current, direction);
        }
      }
      return vertical ? { row: next, column } : { row, column: next };
    }
    default: return undefined;
  }
}
