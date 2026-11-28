import {
  cellKey, findCell, isCoordinate, ownDto,
  type CellDocument, type CellReference, type SheetDocument, type WorkbookDocument,
} from './workbook';
import { requireDocument } from './workbook-validation';

export function equalDto(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object' ||
      Array.isArray(left) !== Array.isArray(right)) return false;
  const keys = Object.keys(left);
  return keys.length === Object.keys(right).length &&
    keys.every(key => Object.hasOwn(right, key) && equalDto(
      (left as Record<string, unknown>)[key], (right as Record<string, unknown>)[key],
    ));
}

function withoutCaches(cell: CellDocument): Omit<CellDocument, 'cache' | 'spillExtent'> {
  const { cache: _cache, spillExtent: _spillExtent, ...rest } = cell;
  if (rest.spillKind === 'formula') {
    const { spillOwner: _owner, spillKind: _kind, content: _content, ...styled } = rest;
    return { ...styled, content: { kind: 'blank' } };
  }
  return rest;
}

function isAuthoredCell(cell: CellDocument): boolean {
  return cell.spillKind !== 'formula' || cell.style !== undefined;
}

export function sameAuthoredCell(left: CellDocument, right: CellDocument): boolean {
  return left === right || equalDto(withoutCaches(left), withoutCaches(right));
}

export function sameAuthoredWorkbook(left: WorkbookDocument, right: WorkbookDocument): boolean {
  if (left === right) return true;
  if (left.id !== right.id || left.formatVersion !== right.formatVersion || left.dateSystem !== right.dateSystem ||
      !equalDto(left.names, right.names) || !equalDto(left.pivots ?? [], right.pivots ?? []) ||
      !equalDto(left.charts ?? [], right.charts ?? []) || left.sheets.length !== right.sheets.length) return false;
  return left.sheets.every((sheet, index) => {
    const other = right.sheets[index]!;
    if (sheet === other) return true;
    const leftCells = sheet.cells.filter(isAuthoredCell);
    const rightCells = other.cells.filter(isAuthoredCell);
    if (!equalDto(sheetState(sheet), sheetState(other)) || leftCells.length !== rightCells.length) return false;
    return leftCells.every(cell => {
      const match = findCell(other, cell);
      return match !== undefined && isAuthoredCell(match) && sameAuthoredCell(cell, match);
    });
  });
}

function sheetState({ cells: _cells, rows, columns, merges, filter, ...state }: SheetDocument): unknown {
  return {
    ...state,
    rows: [...rows].sort((a, b) => a.index - b.index),
    columns: [...columns].sort((a, b) => a.index - b.index),
    merges: [...merges].sort((a, b) => cellKey(a) - cellKey(b)),
    filter: filter && { ...filter, criteria: [...filter.criteria].sort((a, b) => a.column - b.column) },
  };
}

function exactSheetState({ cells: _cells, ...state }: SheetDocument): unknown {
  return state;
}

export interface CellChange {
  readonly target: CellReference;
  readonly before?: CellDocument;
  readonly after?: CellDocument;
  readonly beforeIndex?: number;
  readonly afterIndex?: number;
}
export type WorkbookDelta =
  | { readonly kind: 'cells'; readonly changes: readonly CellChange[] }
  | { readonly kind: 'structure'; readonly before: WorkbookDocument; readonly after: WorkbookDocument };

export function applyCellChanges(document: WorkbookDocument, changes: readonly CellChange[]): WorkbookDocument {
  if (changes.length === 0) return document;
  const groups = new Map<string, Map<number, CellChange>>();
  for (const change of changes) {
    requireDocument(!change.target.invalid && isCoordinate(change.target) &&
      (!change.after || change.after.row === change.target.row && change.after.column === change.target.column),
      'Invalid coordinate in cell delta.');
    const { sheetId } = change.target;
    let group = groups.get(sheetId);
    if (!group) groups.set(sheetId, group = new Map());
    const key = cellKey(change.target);
    requireDocument(!group.has(key), 'Duplicate cell delta target.');
    group.set(key, change);
  }
  const sheets = document.sheets.map(sheet => {
    const group = groups.get(sheet.id);
    if (!group) return sheet;
    groups.delete(sheet.id);
    const remainder: CellDocument[] = [];
    const positioned: CellChange[] = [];
    for (const cell of sheet.cells) {
      const change = group.get(cellKey(cell));
      if (!change) remainder.push(cell);
      else {
        group.delete(cellKey(cell));
        if (change.afterIndex !== undefined) positioned.push(change);
        else if (change.after) remainder.push(change.after);
      }
    }
    for (const change of group.values()) {
      if (change.afterIndex !== undefined) positioned.push(change);
      else if (change.after) remainder.push(change.after);
    }
    const cells = new Array<CellDocument>(remainder.length + positioned.length);
    for (const change of positioned) {
      const index = change.afterIndex!;
      requireDocument(Number.isInteger(index) && index >= 0 && index < cells.length && !cells[index] && change.after,
        'Invalid cell position in delta.');
      cells[index] = change.after;
    }
    let next = 0;
    for (let i = 0; i < cells.length; i++) cells[i] ??= remainder[next++]!;
    return { ...sheet, cells };
  });
  requireDocument(groups.size === 0, 'Cell delta refers to a missing sheet.');
  return ownDto({ ...document, sheets });
}

export function applyDelta(document: WorkbookDocument, delta: WorkbookDelta): WorkbookDocument {
  return delta.kind === 'structure' ? delta.after : applyCellChanges(document, delta.changes);
}

export function invertDelta(delta: WorkbookDelta): WorkbookDelta {
  if (delta.kind === 'structure') return ownDto({ kind: 'structure', before: delta.after, after: delta.before });
  return ownDto({
    kind: 'cells',
    changes: delta.changes.map(({ target, before, after, beforeIndex, afterIndex }) => ({
      target, ...(after && { before: after }), ...(afterIndex !== undefined && { beforeIndex: afterIndex }),
      ...(before && { after: before }), ...(beforeIndex !== undefined && { afterIndex: beforeIndex }),
    })),
  });
}

export function diffWorkbook(before: WorkbookDocument, after: WorkbookDocument): WorkbookDelta {
  if (before.id !== after.id || before.dateSystem !== after.dateSystem || before.formatVersion !== after.formatVersion ||
      !equalDto(before.names, after.names) || !equalDto(before.pivots ?? [], after.pivots ?? []) ||
      !equalDto(before.charts ?? [], after.charts ?? []) || before.sheets.length !== after.sheets.length ||
      before.sheets.some((sheet, i) => sheet !== after.sheets[i] &&
        !equalDto(exactSheetState(sheet), exactSheetState(after.sheets[i]!)))) {
    return ownDto({ kind: 'structure', before, after });
  }
  const changes: CellChange[] = [];
  before.sheets.forEach((sheet, i) => {
    const other = after.sheets[i]!;
    if (sheet.cells === other.cells) return;
    const nextIndexes = new Map(other.cells.map((cell, index) => [cellKey(cell), index]));
    const oldKeys = new Set<number>();
    let previous = -1;
    let reordered = false;
    const changed = new Set<number>();
    sheet.cells.forEach((cell, index) => {
      const key = cellKey(cell);
      oldKeys.add(key);
      const afterIndex = nextIndexes.get(key);
      const next = afterIndex === undefined ? undefined : other.cells[afterIndex]!;
      if (!equalDto(cell, next)) {
        changed.add(key);
        changes.push({
          target: { sheetId: sheet.id, row: cell.row, column: cell.column },
          before: cell, beforeIndex: index, ...(next && { after: next, afterIndex: afterIndex! }),
        });
      } else {
        reordered ||= afterIndex! < previous;
        previous = afterIndex!;
      }
    });
    if (reordered) sheet.cells.forEach((cell, index) => {
      const key = cellKey(cell);
      if (changed.has(key)) return;
      const afterIndex = nextIndexes.get(key)!;
      changes.push({
        target: { sheetId: sheet.id, row: cell.row, column: cell.column },
        before: cell, beforeIndex: index, after: other.cells[afterIndex]!, afterIndex,
      });
    });
    other.cells.forEach((cell, index) => {
      if (!oldKeys.has(cellKey(cell))) changes.push({
        target: { sheetId: sheet.id, row: cell.row, column: cell.column }, after: cell, afterIndex: index,
      });
    });
  });
  return ownDto({ kind: 'cells', changes });
}

// Conservative accounting of retained DTOs in both directions, including keys,
// strings and per-record overhead. It does not measure process memory.
export function estimateDtoBytes(value: unknown): number {
  if (typeof value === 'string') return 24 + 2 * value.length;
  if (value === null || typeof value !== 'object') return 16;
  if (Array.isArray(value)) return 64 + value.reduce<number>((size, item) => size + 16 + estimateDtoBytes(item), 0);
  return Object.entries(value).reduce((size, [key, item]) => size + 32 + estimateDtoBytes(key) + estimateDtoBytes(item), 64);
}
