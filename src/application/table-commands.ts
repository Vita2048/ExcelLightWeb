import { columnLetters, parseFormula, quoteSheet, referenceBox, renameTableReferences, shiftEnclosedRows } from '../core/formula-references';
import { formatCellValue } from '../core/formats/cell-format';
import { cellKey, findCell, type CellContent, type CellDocument, type IdFactory, type SheetDocument,
  type TableDesign, type WorkbookDocument, type WorkbookName } from '../core/workbook';
import { lastColumn, lastRow, ownsTableFilter, tableDesign, tableInputExtent } from '../core/workbook-names';
import { ordinalNameKey, overlaps, requireDocument, validateExtent } from '../core/workbook-validation';
import { editAxis } from './structure-commands';
import { renamePythonContent } from '../python/references';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { StructureCommand } from './workbook-session';

type Extent = WorkbookName['area'];

export function refersTo(document: WorkbookDocument, area: Extent): string {
  if (area.invalid) return '=#REF!';
  const sheet = document.sheets.find(sheet => sheet.id === area.sheetId)!;
  const address = (row: number, column: number) => '$' + columnLetters(column) + '$' + row;
  const end = area.rows === 1 && area.columns === 1 ? '' : ':' + address(lastRow(area), lastColumn(area));
  return '=' + quoteSheet(sheet.name) + address(area.row, area.column) + end;
}

export function parseNameExtent(document: WorkbookDocument, sheetId: string, value: string): Extent {
  const source = '=' + value.trim().replace(/^=/u, '');
  let parsed: ReturnType<typeof parseFormula> | undefined;
  try { parsed = parseFormula(source); } catch { /* Show the same extent guidance for all syntax errors. */ }
  const ref = parsed?.expression.kind === 'reference' ? parsed.expression.reference : undefined;
  requireDocument(ref && ref.first.row !== undefined && ref.first.column !== undefined &&
    source.slice(1, ref.start).trim() === '' && source.slice(ref.end).trim() === '',
  'Enter one A1 cell or rectangle, for example =Sheet1!$A$3:$F$28.');
  const sheet = document.sheets.find(sheet => ref.sheet === undefined ? sheet.id === sheetId :
    ordinalNameKey(sheet.name) === ordinalNameKey(ref.sheet));
  requireDocument(sheet, 'The reference must identify an existing worksheet.');
  const bounds = referenceBox(ref);
  const area = { sheetId: sheet.id, row: bounds.top, column: bounds.left,
    rows: bounds.bottom - bounds.top + 1, columns: bounds.right - bounds.left + 1 };
  validateExtent(area);
  return area;
}

export function nextTableName(document: WorkbookDocument): string {
  const used = new Set(document.names.map(name => ordinalNameKey(name.name)));
  let n = 1;
  while (used.has('TABLE_' + n)) n++;
  return 'Table_' + n;
}

function tableSheet(document: WorkbookDocument, area: Extent) {
  validateExtent(area);
  const sheet = document.sheets.find(sheet => sheet.id === area.sheetId);
  requireDocument(sheet, 'The worksheet no longer exists.');
  return sheet;
}

function replaceSheet(document: WorkbookDocument, next: SheetDocument): WorkbookDocument {
  return { ...document, sheets: document.sheets.map(sheet => sheet.id === next.id ? next : sheet) };
}

function renamePython(document: WorkbookDocument, oldName: string, newName: string): WorkbookDocument {
  return { ...document, sheets: document.sheets.map(sheet => {
    const cells = sheet.cells.map(cell => {
      if (cell.content.kind === 'python') {
        return { ...cell, content: renamePythonContent(cell.content, oldName, newName) };
      }
      if (cell.content.kind === 'formula') {
        const source = renameTableReferences(cell.content.source, oldName, newName);
        if (source === cell.content.source) return cell;
        const { cache: _cache, ...rest } = cell;
        return { ...rest, content: { ...cell.content, source } };
      }
      return cell;
    });
    return cells.some((cell, index) => cell !== sheet.cells[index]) ? { ...sheet, cells } : sheet;
  }) };
}

function put(document: WorkbookDocument, entry: WorkbookName): WorkbookDocument {
  const prior = document.names.find(name => name.id === entry.id);
  requireDocument(!prior || !!prior.isTable === !!entry.isTable, 'A definition cannot change kind.');
  const result = { ...document, names: prior ? document.names.map(name => name === prior ? entry : name) : [...document.names, entry] };
  return prior && prior.name !== entry.name ? renamePython(result, prior.name, entry.name) : result;
}

function tableById(document: WorkbookDocument, id: string): WorkbookName {
  const table = document.names.find(name => name.id === id && name.isTable && !name.area.invalid);
  requireDocument(table, 'Select a valid table.');
  return table;
}

function checkTableSpace(document: WorkbookDocument, table: WorkbookName) {
  const { area } = table;
  const sheet = tableSheet(document, area);
  requireDocument(!document.names.some(name => name.id !== table.id && name.isTable && !name.area.invalid &&
    name.area.sheetId === area.sheetId && overlaps(name.area, area)), 'Tables cannot overlap.');
  requireDocument(!sheet.merges.some(merge => overlaps(merge, area)), 'Unmerge cells before including them in a table.');
  return sheet;
}

function rowCells(document: WorkbookDocument, area: Extent, row: number) {
  const sheet = tableSheet(document, area);
  return Array.from({ length: area.columns }, (_, index) => findCell(sheet, { row, column: area.column + index }));
}

function writeRow(document: WorkbookDocument, area: Extent, row: number,
  contents: readonly (CellContent | undefined)[], makeId: IdFactory, offset = 0): WorkbookDocument {
  const sheet = tableSheet(document, area);
  const cells = new Map(sheet.cells.map(cell => [cellKey(cell), cell]));
  contents.forEach((content, index) => {
    if (content === undefined) return;
    const point = { row, column: area.column + offset + index };
    const key = cellKey(point);
    const old = cells.get(key);
    if (content.kind === 'text' && old?.content.kind === 'text' && old.content.value === content.value) return;
    requireDocument(old?.content.kind !== 'python', 'A table header or total cannot replace a Python cell.');
    if (content.kind === 'blank' && !old?.style) cells.delete(key);
    else cells.set(key, { ...point, id: old?.id ?? makeId(), content, ...(old?.style && { style: old.style }) });
  });
  return replaceSheet(document, { ...sheet, cells: [...cells.values()] });
}

const blank = { kind: 'blank' } as const;
const labelContent = (value: string) => ({ kind: 'text', value } as const);

function labelsFromHeader(document: WorkbookDocument, area: Extent) {
  return rowCells(document, area, area.row).map(cell => cell?.content.kind === 'text' ? cell.content.value : '');
}

function uniqueLabels(labels: readonly string[]): string[] {
  const used = new Set<string>();
  return labels.map((text, index) => {
    const base = text.trim() ? text : 'Column' + (index + 1);
    let value = base, suffix = 2;
    while (used.has(ordinalNameKey(value))) value = base + suffix++;
    used.add(ordinalNameKey(value));
    return value;
  });
}

function validLabels(labels: readonly string[]) {
  return labels.every(label => label.trim()) && new Set(labels.map(ordinalNameKey)).size === labels.length;
}

export function selectionHasHeaders(document: WorkbookDocument, area: Extent): boolean {
  return validLabels(labelsFromHeader(document, area));
}

function normalizeHeaders(document: WorkbookDocument, table: WorkbookName, makeId: IdFactory) {
  if (!tableDesign(table).headerRow) return document;
  const { area } = table;
  const labels = uniqueLabels(rowCells(document, area, area.row).map(cell =>
    cell?.content.kind === 'text' ? cell.content.value : formatCellValue(cell).text));
  return writeRow(document, area, area.row, labels.map(labelContent), makeId);
}

function withoutFilter(sheet: SheetDocument): SheetDocument {
  const { filter: _filter, ...rest } = sheet;
  return rest;
}

function syncFilter(document: WorkbookDocument, table: WorkbookName, prior?: WorkbookName): WorkbookDocument {
  const source = prior && !prior.area.invalid ? tableSheet(document, prior.area) : undefined;
  const owned = source && prior && ownsTableFilter(source.filter, prior) ? source.filter : undefined;
  if (source && owned) document = replaceSheet(document, withoutFilter(source));
  const { headerRow, totalRow: _total, filterButton } = tableDesign(table);
  if (!headerRow || !filterButton || table.area.invalid) return document;
  const sheet = tableSheet(document, table.area);
  requireDocument(!owned || !sheet.filter, 'Remove the destination sheet filter before moving this filtered table.');
  if (sheet.filter) return document;
  const { row, column, rows, columns } = tableInputExtent(table);
  const inColumn = (index: number) => index >= column && index < column + columns;
  const sort = owned?.sort;
  return replaceSheet(document, { ...sheet, filter: { row, column, rows, columns,
    criteria: owned?.criteria.filter(item => inColumn(item.column)) ?? [],
    ...(sort && inColumn(sort.column) && { sort }) } });
}

export function defineName(basis: WorkbookSnapshot, entry: WorkbookName): StructureCommand {
  return { kind: 'structure', label: 'Define workbook name', document: put(basis.document, entry) };
}

export function deleteName(basis: WorkbookSnapshot, id: string): StructureCommand {
  const entry = basis.document.names.find(name => name.id === id);
  requireDocument(entry && !entry.isTable, 'Select the entire table on the worksheet and press Delete to remove it.');
  return { kind: 'structure', label: 'Delete workbook name',
    document: renamePython({ ...basis.document, names: basis.document.names.filter(name => name !== entry) },
      entry.name, '#REF!') };
}

export function createTable(basis: WorkbookSnapshot, entry: WorkbookName, hasHeaders: boolean, makeId: IdFactory): StructureCommand {
  let { document } = basis;
  let table: WorkbookName = { ...entry, isTable: true, design: { ...tableDesign(entry), style: entry.design?.style ?? 'Blue',
    headerRow: true, totalRow: false } };
  const { area } = table;
  const sheet = checkTableSpace(document, table);
  const populated = sheet.cells.some(cell => overlaps(area, { ...cell, rows: 1, columns: 1 }) && cell.content.kind !== 'blank');
  if (!hasHeaders && populated) {
    document = editAxis(basis, sheet.id, { axis: 'row', start: area.row, count: 1,
      lane: [area.column, lastColumn(area)], delete: false }).document;
    table = { ...table, area: { ...area, rows: area.rows + 1 } };
  } else if (!populated && area.rows === 1) {
    table = { ...table, area: { ...area, rows: 2 } };
  }
  checkTableSpace(document, table);
  document = put(normalizeHeaders(document, table, makeId), table);
  return { kind: 'structure', label: 'Create table', document: syncFilter(document, table) };
}

export const totalFunctions = [
  ['None', 0], ['Average', 101], ['Count', 103], ['Count Numbers', 102], ['Max', 104], ['Min', 105], ['Sum', 109],
] as const;

export function totalFunction(cell: CellDocument | undefined): number {
  if (cell?.content.kind !== 'formula') return 0;
  const source = cell.content.source;
  return Number(/^=SUBTOTAL\(\s*(\d+)\s*,/iu.exec(source)?.[1] ?? 0);
}

function totalContent(table: WorkbookName, column: number, code: number): CellContent {
  if (!code) return blank;
  const letter = columnLetters(column);
  const input = tableInputExtent(table);
  const first = input.row + Number(input.headerRow), last = lastRow(input);
  const source = first <= last ? `=SUBTOTAL(${code},${letter}${first}:${letter}${last})`
    : code === 101 ? '=0/0' : '=0';
  return { kind: 'formula', source };
}

function ensureTotalSpace(document: WorkbookDocument, table: WorkbookName) {
  checkTableSpace(document, table);
  requireDocument(rowCells(document, table.area, lastRow(table.area)).every(cell => !cell || cell.content.kind === 'blank'),
    'The total row would overwrite occupied cells. Move the content below the table first.');
}

function shiftRowBlock(document: WorkbookDocument, lane: Extent, firstRow: number, lastRow: number, dr: number): WorkbookDocument {
  const sheet = tableSheet(document, lane);
  const cells = new Map(sheet.cells.map(cell => [cellKey(cell), cell]));
  const endColumn = lastColumn(lane);
  const downward = dr > 0;
  for (let row = downward ? lastRow : firstRow; downward ? row >= firstRow : row <= lastRow; row += downward ? -1 : 1) {
    for (let column = lane.column; column <= endColumn; column++) {
      const source = cells.get(cellKey({ row, column }));
      const destKey = cellKey({ row: row + dr, column });
      if (!source) {
        cells.delete(destKey);
        continue;
      }
      const content = source.content.kind === 'formula'
        ? { kind: 'formula' as const, source: shiftEnclosedRows(source.content.source, firstRow, lastRow, dr) }
        : source.content;
      const { cache: _cache, requiresExternalCalculation: _external, ...rest } = source;
      cells.set(destKey, { ...rest, row: row + dr, column, content });
    }
  }
  const vacated = downward ? firstRow : lastRow;
  for (let column = lane.column; column <= endColumn; column++) cells.delete(cellKey({ row: vacated, column }));
  return replaceSheet(document, { ...sheet, cells: [...cells.values()] });
}

function columnLane(prior: Extent, next: Extent): Extent {
  const column = Math.min(prior.column, next.column);
  const columns = Math.max(lastColumn(prior), lastColumn(next)) - column + 1;
  return { ...prior, column, columns };
}

export function setTotalFunction(basis: WorkbookSnapshot, id: string, column: number, code: number, makeId: IdFactory): StructureCommand {
  const table = tableById(basis.document, id);
  requireDocument(Number.isInteger(column) && column >= table.area.column && column <= lastColumn(table.area) &&
    tableDesign(table).totalRow && totalFunctions.some(([, option]) => option === code),
  'Select a total-row cell and a supported function.');
  const document = writeRow(basis.document, table.area, lastRow(table.area),
    [totalContent(table, column, code)], makeId, column - table.area.column);
  return { kind: 'structure', label: 'Table total function', document };
}

export function changeTableDesign(basis: WorkbookSnapshot, id: string, patch: Partial<TableDesign>, makeId: IdFactory): StructureCommand {
  let { document } = basis;
  const prior = tableById(document, id);
  const old = tableDesign(prior), design = { ...old, ...patch };
  let table = { ...prior, design };
  let { area } = prior;
  if (old.headerRow && !design.headerRow) {
    requireDocument(area.rows > 1 + Number(old.totalRow), 'Add a data row before removing the table header.');
    const labels = labelsFromHeader(document, area);
    requireDocument(validLabels(labels), 'Table headers must be nonblank unique text. Repair the first row of the table.');
    document = writeRow(document, area, area.row, labels.map(() => blank), makeId);
    area = { ...area, row: area.row + 1, rows: area.rows - 1 };
    table = { ...table, headerLabels: labels };
  } else if (!old.headerRow && design.headerRow) {
    requireDocument(area.row > 1, 'There is no worksheet row above the table for its header.');
    const labels = prior.headerLabels!;
    const row = area.row - 1;
    requireDocument(rowCells(document, area, row).every((cell, i) => !cell || cell.content.kind === 'blank' ||
      cell.content.kind === 'text' && cell.content.value === labels[i]),
    'The former header row contains changed worksheet values. Move them before restoring the header.');
    document = writeRow(document, area, row, labels.map(labelContent), makeId);
    area = { ...area, row, rows: area.rows + 1 };
    const { headerLabels: _labels, ...rest } = table;
    table = rest;
  }
  if (!old.totalRow && design.totalRow) {
    area = { ...area, rows: area.rows + 1 };
    ensureTotalSpace(document, { ...table, area });
    const totals = Array.from({ length: area.columns }, (_, i) => i === area.columns - 1
      ? totalContent({ ...table, area }, area.column + i, 109) : i === 0 ? labelContent('Total') : undefined);
    document = writeRow(document, area, lastRow(area), totals, makeId);
  } else if (old.totalRow && !design.totalRow) {
    requireDocument(area.rows > 1, 'Add a data row before removing totals.');
    document = writeRow(document, area, lastRow(area), Array.from({ length: area.columns }, () => blank), makeId);
    area = { ...area, rows: area.rows - 1 };
  }
  table = { ...table, area };
  checkTableSpace(document, table);
  return { kind: 'structure', label: 'Table Design', document: syncFilter(put(document, table), table, prior) };
}

export function resizeTable(basis: WorkbookSnapshot, entry: WorkbookName, makeId: IdFactory): StructureCommand {
  let { document } = basis;
  const prior = document.names.find(name => name.id === entry.id);
  requireDocument(prior?.isTable, 'Select a table to resize.');
  let table = { ...prior, name: entry.name, comment: entry.comment ?? '', area: entry.area };
  const { area } = table;
  if (area.invalid) return { kind: 'structure', label: 'Edit table', document: put(document, table) };
  const { headerRow, totalRow } = tableDesign(table);
  checkTableSpace(document, table);
  requireDocument(!totalRow || area.rows > Number(headerRow), 'Include a row for table totals.');
  if (!headerRow) {
    const labels = prior.headerLabels ?? [];
    table = { ...table, headerLabels: uniqueLabels(Array.from({ length: area.columns }, (_, i) =>
      labels[area.column + i - prior.area.column] ?? '')) };
  }
  const moved = prior.area.invalid || prior.area.sheetId !== area.sheetId || prior.area.row !== area.row ||
    prior.area.column !== area.column || prior.area.rows !== area.rows || prior.area.columns !== area.columns;
  if (!moved) return { kind: 'structure', label: 'Edit table', document: put(document, table) };
  if (totalRow) {
    const previous = prior.area.invalid ? [] : rowCells(document, prior.area, lastRow(prior.area));
    const functions = previous.map(totalFunction);
    const sameOrigin = !prior.area.invalid && prior.area.sheetId === area.sheetId && prior.area.row === area.row;
    const oldLast = prior.area.invalid ? 0 : lastRow(prior.area);
    const newLast = lastRow(area);
    if (sameOrigin && newLast < oldLast) document = shiftRowBlock(document, prior.area, newLast, oldLast - 1, 1);
    else if (sameOrigin && newLast > oldLast) document = shiftRowBlock(document, columnLane(prior.area, area), oldLast + 1, newLast, -1);
    else if (previous.length) document = writeRow(document, prior.area, lastRow(prior.area), previous.map(() => blank), makeId);
    ensureTotalSpace(document, table);
    const totals = Array.from({ length: area.columns }, (_, i) => {
      const code = functions[i] ?? 0;
      if (code) return totalContent(table, area.column + i, code);
      if (i === 0 && area.columns > 1 && functions[0] === 0) return labelContent('Total');
      return blank;
    });
    document = writeRow(document, area, lastRow(area), totals, makeId);
  }
  document = put(normalizeHeaders(document, table, makeId), table);
  return { kind: 'structure', label: 'Resize table', document: syncFilter(document, table, prior) };
}

export function renameTable(basis: WorkbookSnapshot, id: string, name: string): StructureCommand {
  const table = tableById(basis.document, id);
  const trimmed = name.trim();
  requireDocument(trimmed.length > 0, 'Enter a valid table name.');
  return { kind: 'structure', label: 'Rename table', document: put(basis.document, { ...table, name: trimmed }) };
}

export function deleteTableSelection(basis: WorkbookSnapshot, selections: readonly { sheetId: string; row: number; column: number; rows: number; columns: number }[]): StructureCommand | undefined {
  const { document } = basis;
  const tablesToRemove: WorkbookName[] = [];
  for (const table of document.names) {
    if (!table.isTable || table.area.invalid) continue;
    const isEnclosed = selections.some(sel =>
      sel.sheetId === table.area.sheetId &&
      sel.row <= table.area.row &&
      lastRow(sel) >= lastRow(table.area) &&
      sel.column <= table.area.column &&
      lastColumn(sel) >= lastColumn(table.area)
    );
    if (isEnclosed) tablesToRemove.push(table);
  }
  if (!tablesToRemove.length) return undefined;

  const removedIds = new Set(tablesToRemove.map(t => t.id));
  const names = document.names.filter(n => !removedIds.has(n.id));
  const sheets = document.sheets.map(sheet => {
    const sheetSelections = selections.filter(s => s.sheetId === sheet.id);
    if (!sheetSelections.length) return sheet;
    const cells = sheet.cells.map(cell => {
      const inSel = sheetSelections.some(sel =>
        cell.row >= sel.row && cell.row <= lastRow(sel) &&
        cell.column >= sel.column && cell.column <= lastColumn(sel)
      );
      if (!inSel) return cell;
      const { cache: _c, ...restCell } = cell;
      return { ...restCell, content: blank, requiresExternalCalculation: false };
    }).filter(cell => cell.content.kind !== 'blank' || cell.style !== undefined);

    let filter = sheet.filter;
    if (filter && tablesToRemove.some(t => ownsTableFilter(filter, t))) {
      filter = undefined;
    }
    const { filter: _f, ...rest } = sheet;
    return { ...rest, cells, ...(filter && { filter }) };
  });

  return {
    kind: 'structure',
    label: 'Delete selection',
    document: { ...document, sheets, names },
  };
}
