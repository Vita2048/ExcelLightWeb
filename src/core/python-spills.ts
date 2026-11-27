import { cellKey, documentLimits, type CellCache, type CellContent, type CellDocument, type CellExtent, type CellReference,
  type PythonBlock, type PythonTable, type SheetDocument, type WorkbookDocument } from './workbook';
import { lastColumn, lastRow } from './workbook-names';
import { overlaps, requireDocument } from './workbook-validation';

export interface SpillUpdate {
  readonly target: CellReference;
  readonly cache?: CellCache;
  readonly spill?: {
    readonly rows: number;
    readonly columns: number;
    readonly cells: readonly (CellCache | null)[];
  };
}

function contentFromCache(value: CellCache | null): CellContent {
  if (!value) return { kind: 'blank' };
  switch (value.kind) {
    case 'number': return { kind: 'number', value: value.value };
    case 'text': return { kind: 'text', value: value.value };
    case 'boolean': return { kind: 'boolean', value: value.value };
    default: return { kind: 'blank' };
  }
}

function spillExtent(anchor: CellReference, table: PythonTable): CellExtent {
  return { row: anchor.row, column: anchor.column, rows: table.rows, columns: table.columns };
}

function revertGeneratedCell(cell: CellDocument): CellDocument[] {
  const { spillOwner: _owner, spillKind: _kind, spillExtent: _extent, cache: _cache,
    requiresExternalCalculation: _external, content: _content, ...rest } = cell;
  return cell.style ? [{ ...rest, content: { kind: 'blank' as const } }] : [];
}

/** Apply one validated Python result and its generated cells as a single workbook snapshot. */
export function applyPythonResult(document: WorkbookDocument, anchor: CellReference, block: PythonBlock,
  makeId: () => string): WorkbookDocument {
  const sheetIndex = document.sheets.findIndex(sheet => sheet.id === anchor.sheetId);
  requireDocument(sheetIndex >= 0, 'Python anchor worksheet no longer exists.');
  const sheet = document.sheets[sheetIndex]!;
  const original = sheet.cells.find(cell => cell.row === anchor.row && cell.column === anchor.column);
  requireDocument(original?.content.kind === 'python' && original.content.blockId === block.id,
    'Python anchor changed before its result was applied.');

  const oldChildren = new Map<string, CellDocument>();
  const sheets = document.sheets.map(current => ({
    ...current,
    cells: current.cells.flatMap<CellDocument>(cell => {
      if (cell.spillOwner !== block.id || cell.spillKind === 'formula') return [cell];
      oldChildren.set(`${current.id}:${cell.row}:${cell.column}`, cell);
      return revertGeneratedCell(cell);
    }),
  }));
  let updatedSheet = sheets[sheetIndex]!;
  const anchorCell = updatedSheet.cells.find(cell => cell.id === original.id)!;
  updatedSheet = { ...updatedSheet, cells: updatedSheet.cells.filter(cell => cell.id !== anchorCell.id) };

  if (block.spillActive) {
    const table = block.lastTable;
    requireDocument(table, 'Python spill is missing its table result.');
    requireDocument(table.rows >= 1 && table.columns >= 1 && table.rows <= documentLimits.rows &&
      table.columns <= documentLimits.columns && table.rows * table.columns <= 100_000 &&
      table.cells.length === table.rows * table.columns, 'Python result table exceeds workbook limits.');
    const extent = spillExtent(anchor, table);
    requireDocument(lastRow(extent) <= documentLimits.rows && lastColumn(extent) <= documentLimits.columns,
      '#SPILL! Output exceeds worksheet limits.');
    requireDocument(!sheet.merges.some(merge => overlaps(merge, extent)), '#SPILL! Output overlaps merged cells.');
    requireDocument(!document.names.some(name => name.isTable && !name.area.invalid && name.area.sheetId === sheet.id &&
      overlaps(name.area, extent)), '#SPILL! Output overlaps a workbook table.');

    const byPosition = new Map(updatedSheet.cells.map(cell => [`${cell.row}:${cell.column}`, cell]));
    for (let rowOffset = 0; rowOffset < table.rows; rowOffset++) for (let columnOffset = 0; columnOffset < table.columns; columnOffset++) {
      if (rowOffset === 0 && columnOffset === 0) continue;
      const row = anchor.row + rowOffset;
      const column = anchor.column + columnOffset;
      const existing = byPosition.get(`${row}:${column}`);
      const oldSpillCell = oldChildren.get(`${sheet.id}:${row}:${column}`);
      requireDocument(!existing || (existing.content.kind === 'blank' && !existing.spillOwner && !(existing.spillKind === 'formula' && existing.cache)) ||
        (existing.spillOwner === block.id && existing.spillKind !== 'formula'),
        '#SPILL! Destination contains data; clear it and Run again.');
      const value = table.cells[rowOffset * table.columns + columnOffset] ?? null;
      const resultCell: CellDocument = {
        id: oldSpillCell?.id ?? existing?.id ?? makeId(), row, column, content: contentFromCache(value),
        ...(existing?.style && { style: existing.style }), spillOwner: block.id, spillKind: 'python',
      };
      const existingIndex = updatedSheet.cells.findIndex(cell => cell.row === row && cell.column === column);
      if (existingIndex >= 0) updatedSheet.cells[existingIndex] = resultCell;
      else updatedSheet.cells.push(resultCell);
    }
  }

  updatedSheet.cells.push({ ...anchorCell, content: { kind: 'python', blockId: block.id, source: block.source, block } });
  updatedSheet = { ...updatedSheet, cells: updatedSheet.cells.sort((left, right) => cellKey(left) - cellKey(right)) };
  return { ...document, sheets: sheets.map((current, index) => index === sheetIndex ? updatedSheet : current) };
}

/** Remove generated Python cells whose owning anchor was deleted or whose output was cleared. */
export function removeOrphanedPythonSpillCells(document: WorkbookDocument): WorkbookDocument {
  const active = new Set(document.sheets.flatMap(sheet => sheet.cells.flatMap(cell =>
    cell.content.kind === 'python' && cell.content.block?.spillActive ? [cell.content.blockId] : [])));
  let changed = false;
  const sheets = document.sheets.map(sheet => {
    let sheetChanged = false;
    const cells = sheet.cells.flatMap<CellDocument>(cell => {
      if (!cell.spillOwner || cell.spillKind === 'formula' || active.has(cell.spillOwner)) return [cell];
      changed = sheetChanged = true;
      return revertGeneratedCell(cell);
    });
    return sheetChanged ? { ...sheet, cells } : sheet;
  });
  return changed ? { ...document, sheets } : document;
}

function canSpillInto(document: WorkbookDocument, sheet: SheetDocument, occupied: ReadonlyMap<number, CellDocument>,
  anchor: CellDocument, rows: number, columns: number): boolean {
  if (rows < 1 || columns < 1 || rows * columns > 100_000) return false;
  const extent: CellExtent = { row: anchor.row, column: anchor.column, rows, columns };
  if (lastRow(extent) > documentLimits.rows || lastColumn(extent) > documentLimits.columns) return false;
  if (sheet.merges.some(merge => overlaps(merge, extent))) return false;
  if (document.names.some(name => name.isTable && !name.area.invalid && name.area.sheetId === sheet.id &&
      overlaps(name.area, extent))) return false;
  for (let rowOffset = 0; rowOffset < rows; rowOffset++) for (let colOffset = 0; colOffset < columns; colOffset++) {
    if (rowOffset === 0 && colOffset === 0) continue;
    const existing = occupied.get(cellKey({ row: anchor.row + rowOffset, column: anchor.column + colOffset }));
    if (existing && (existing.content.kind !== 'blank' || existing.spillOwner !== undefined || existing.cache !== undefined)) {
      return false;
    }
  }
  return true;
}

/** Apply native formula cache and dynamic-array spill updates atomically. */
export function applyFormulaSpillUpdates(document: WorkbookDocument, updates: readonly SpillUpdate[],
  makeId: () => string): WorkbookDocument {
  if (updates.length === 0) return document;
  const bySheet = new Map<string, SpillUpdate[]>();
  for (const update of updates) {
    let list = bySheet.get(update.target.sheetId);
    if (!list) bySheet.set(update.target.sheetId, list = []);
    list.push(update);
  }

  let changed = false;
  const sheets = document.sheets.map(sheet => {
    const sheetUpdates = bySheet.get(sheet.id);
    if (!sheetUpdates) return sheet;
    bySheet.delete(sheet.id);

    const byKey = new Map<number, CellDocument>(sheet.cells.map(cell => [cellKey(cell), cell]));
    const updatedAnchors = new Set<string>();
    const resolvedAnchors: { anchor: CellDocument; update: SpillUpdate }[] = [];
    for (const update of sheetUpdates) {
      const cell = byKey.get(cellKey(update.target));
      requireDocument(cell?.content.kind === 'formula', 'Calculation can update only existing native formula caches.');
      updatedAnchors.add(cell.id);
      resolvedAnchors.push({ anchor: cell, update });
    }

    const oldChildren = new Map<number, CellDocument>();
    const occupied = new Map<number, CellDocument>();
    for (const cell of sheet.cells) {
      if (cell.spillKind === 'formula' && cell.spillOwner && updatedAnchors.has(cell.spillOwner)) {
        oldChildren.set(cellKey(cell), cell);
        for (const reverted of revertGeneratedCell(cell)) occupied.set(cellKey(reverted), reverted);
      } else {
        occupied.set(cellKey(cell), cell);
      }
    }

    for (const { anchor, update } of resolvedAnchors) {
      const { cache: _oldCache, spillExtent: _oldSpill, ...rest } = anchor;
      const baseAnchor: CellDocument = update.cache === undefined ? rest : { ...rest, cache: update.cache };
      occupied.set(cellKey(anchor), baseAnchor);
    }

    for (const { anchor, update } of resolvedAnchors) {
      const spill = update.spill;
      if (!spill || spill.rows * spill.columns <= 1) continue;
      requireDocument(spill.cells.length === spill.rows * spill.columns, 'Formula spill payload length mismatch.');
      const currentAnchor = occupied.get(cellKey(anchor))!;
      if (!canSpillInto(document, sheet, occupied, currentAnchor, spill.rows, spill.columns)) {
        const { spillExtent: _drop, ...rest } = currentAnchor;
        occupied.set(cellKey(anchor), { ...rest, cache: { kind: 'error', value: '#SPILL!' } });
        continue;
      }
      occupied.set(cellKey(anchor), {
        ...currentAnchor,
        spillExtent: { rows: spill.rows, columns: spill.columns },
      });
      for (let rowOffset = 0; rowOffset < spill.rows; rowOffset++) for (let colOffset = 0; colOffset < spill.columns; colOffset++) {
        if (rowOffset === 0 && colOffset === 0) continue;
        const row = anchor.row + rowOffset;
        const column = anchor.column + colOffset;
        const key = cellKey({ row, column });
        const existing = occupied.get(key);
        const oldChild = oldChildren.get(key);
        const val = spill.cells[rowOffset * spill.columns + colOffset] ?? null;
        const isErr = val?.kind === 'error';
        const child: CellDocument = {
          id: oldChild?.id ?? existing?.id ?? makeId(),
          row,
          column,
          content: isErr ? { kind: 'blank' } : contentFromCache(val),
          ...(existing?.style && { style: existing.style }),
          ...(isErr && { cache: val }),
          spillOwner: anchor.id,
          spillKind: 'formula',
        };
        occupied.set(key, child);
      }
    }

    const nextCells = [...occupied.values()].sort((left, right) => cellKey(left) - cellKey(right));
    if (nextCells.length === sheet.cells.length && nextCells.every((cell, index) => cell === sheet.cells[index])) {
      return sheet;
    }
    changed = true;
    return { ...sheet, cells: nextCells };
  });

  requireDocument(bySheet.size === 0, 'Calculation target refers to a missing worksheet.');
  return changed ? { ...document, sheets } : document;
}

/** Remove generated formula spill cells whose anchor was deleted, edited, or whose extent no longer matches. */
export function removeOrphanedFormulaSpillCells(document: WorkbookDocument): WorkbookDocument {
  let changed = false;
  const sheets = document.sheets.map(sheet => {
    const anchors = new Map<string, CellDocument>();
    const childCounts = new Map<string, number>();
    for (const cell of sheet.cells) {
      if (cell.content.kind === 'formula' && cell.spillExtent) anchors.set(cell.id, cell);
      if (cell.spillKind === 'formula' && cell.spillOwner) {
        childCounts.set(cell.spillOwner, (childCounts.get(cell.spillOwner) ?? 0) + 1);
      }
    }
    const validOwners = new Set<string>();
    for (const [id, anchor] of anchors) {
      const expected = anchor.spillExtent!.rows * anchor.spillExtent!.columns - 1;
      if (expected > 0 && childCounts.get(id) === expected) validOwners.add(id);
    }

    let sheetChanged = false;
    const cells = sheet.cells.flatMap<CellDocument>(cell => {
      if (cell.spillKind === 'formula') {
        if (cell.spillOwner && validOwners.has(cell.spillOwner)) return [cell];
        changed = sheetChanged = true;
        return revertGeneratedCell(cell);
      }
      if (cell.spillExtent && !validOwners.has(cell.id)) {
        changed = sheetChanged = true;
        const { spillExtent: _drop, ...rest } = cell;
        return [rest];
      }
      return [cell];
    });
    return sheetChanged ? { ...sheet, cells } : sheet;
  });
  return changed ? { ...document, sheets } : document;
}
