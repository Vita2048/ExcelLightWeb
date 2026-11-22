import { cellKey, documentLimits, type CellContent, type CellDocument, type CellExtent, type CellReference,
  type PythonBlock, type PythonTable, type WorkbookDocument } from './workbook';
import { lastColumn, lastRow } from './workbook-names';
import { overlaps, requireDocument } from './workbook-validation';

function contentFromCache(value: PythonTable['cells'][number]): CellContent {
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
      if (cell.spillOwner !== block.id) return [cell];
      oldChildren.set(`${current.id}:${cell.row}:${cell.column}`, cell);
      const { spillOwner: _owner, cache: _cache, requiresExternalCalculation: _external,
        content: _content, ...rest } = cell;
      return cell.style ? [{ ...rest, content: { kind: 'blank' as const } }] : [];
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

    const byPosition = new Map(sheet.cells.map(cell => [`${cell.row}:${cell.column}`, cell]));
    for (let rowOffset = 0; rowOffset < table.rows; rowOffset++) for (let columnOffset = 0; columnOffset < table.columns; columnOffset++) {
      if (rowOffset === 0 && columnOffset === 0) continue;
      const row = anchor.row + rowOffset;
      const column = anchor.column + columnOffset;
      const existing = byPosition.get(`${row}:${column}`);
      const oldSpillCell = oldChildren.get(`${sheet.id}:${row}:${column}`);
      requireDocument(!existing || existing.content.kind === 'blank' || existing.spillOwner === block.id,
        '#SPILL! Destination contains data; clear it and Run again.');
      const value = table.cells[rowOffset * table.columns + columnOffset] ?? null;
      updatedSheet.cells.push({
        id: oldSpillCell?.id ?? existing?.id ?? makeId(), row, column, content: contentFromCache(value),
        ...(existing?.style && { style: existing.style }), spillOwner: block.id,
      });
    }
  }

  updatedSheet.cells.push({ ...anchorCell, content: { kind: 'python', blockId: block.id, source: block.source, block } });
  updatedSheet = { ...updatedSheet, cells: updatedSheet.cells.sort((left, right) => cellKey(left) - cellKey(right)) };
  return { ...document, sheets: sheets.map((current, index) => index === sheetIndex ? updatedSheet : current) };
}

/** Remove generated cells whose owning anchor was deleted or whose output was cleared. */
export function removeOrphanedPythonSpillCells(document: WorkbookDocument): WorkbookDocument {
  const active = new Set(document.sheets.flatMap(sheet => sheet.cells.flatMap(cell =>
    cell.content.kind === 'python' && cell.content.block?.spillActive ? [cell.content.blockId] : [])));
  let changed = false;
  const sheets = document.sheets.map(sheet => {
    let sheetChanged = false;
    const cells = sheet.cells.flatMap<CellDocument>(cell => {
      if (!cell.spillOwner || active.has(cell.spillOwner)) return [cell];
      changed = sheetChanged = true;
      const { spillOwner: _owner, cache: _cache, requiresExternalCalculation: _external,
        content: _content, ...rest } = cell;
      return cell.style ? [{ ...rest, content: { kind: 'blank' as const } }] : [];
    });
    return sheetChanged ? { ...sheet, cells } : sheet;
  });
  return changed ? { ...document, sheets } : document;
}
