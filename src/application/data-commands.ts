import { formatCellValue } from '../core/formats/cell-format';
import { columnLetters, copyReferences } from '../core/formula-references';
import { maskRows, RowMask } from '../core/row-mask';
import { cellKey, documentLimits, findCell, worksheetViewport, type CellDocument, type CellExtent, type SheetDocument,
  type SheetFilter, type WorkbookDocument } from '../core/workbook';
import { extentContains, lastColumn, lastRow, tableDesign, tableInputExtent } from '../core/workbook-names';
import { overlaps, requireDocument, validateExtent } from '../core/workbook-validation';
import { rewritePythonContent } from '../python/references';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { StructureCommand } from './workbook-session';

// Single-key row sort and one selected-value AutoFilter. Visibility is derived from criteria;
// manually hidden rows stay on the axis and are never cleared by the filter.
export type SortKind = 'number' | 'text' | 'boolean' | 'error' | 'blank';
export interface SortKey {
  readonly kind: SortKind;
  readonly number: number;
  readonly text: string;
  readonly flag: boolean;
  readonly index: number;
}
export type TextOrder = (left: string, right: string, caseSensitive: boolean) => number;

const ordinalOrder: TextOrder = (left, right, caseSensitive) => {
  const first = caseSensitive ? left : left.toLowerCase();
  const second = caseSensitive ? right : right.toLowerCase();
  if (first !== second) return first < second ? -1 : 1;
  return left < right ? -1 : left > right ? 1 : 0;
};

/** Builds the locale collation used by the sort and filter menus. */
export function localeTextOrder(locale: string): TextOrder {
  const insensitive = new Intl.Collator(locale, { sensitivity: 'accent', usage: 'sort' });
  const sensitive = new Intl.Collator(locale, { sensitivity: 'variant', usage: 'sort' });
  return (left, right, caseSensitive) => (caseSensitive ? sensitive : insensitive).compare(left, right);
}
const blankKey = (index: number): SortKey => ({ kind: 'blank', number: 0, text: '', flag: false, index });

/** Display text used as a filter value, shared with the native SUBTOTAL filter key. */
export function filterKey(cell: CellDocument | undefined): string {
  return formatCellValue(cell).text;
}

export function sortKeyFor(cell: CellDocument | undefined, index: number): SortKey {
  if (!cell || cell.content.kind === 'blank') return blankKey(index);
  const content = cell.content;
  if (content.kind === 'number') return { kind: 'number', number: content.value, text: '', flag: false, index };
  if (content.kind === 'text') return content.value.length === 0 ? blankKey(index) :
    { kind: 'text', number: 0, text: content.value, flag: false, index };
  if (content.kind === 'boolean') return { kind: 'boolean', number: 0, text: '', flag: content.value, index };
  const cache = cell.cache;
  if (!cache) return blankKey(index);
  if (cache.kind === 'number') return { kind: 'number', number: cache.value, text: '', flag: false, index };
  if (cache.kind === 'boolean') return { kind: 'boolean', number: 0, text: '', flag: cache.value, index };
  if (cache.kind === 'error') return { kind: 'error', number: 0, text: cache.value, flag: false, index };
  return cache.value.length === 0 ? blankKey(index) :
    { kind: 'text', number: 0, text: cache.value, flag: false, index };
}

function sortRank(kind: SortKind, ascending: boolean): number {
  switch (kind) {
    case 'blank': return 4;
    case 'number': return ascending ? 0 : 3;
    case 'text': return ascending ? 1 : 2;
    case 'boolean': return ascending ? 2 : 1;
    default: return ascending ? 3 : 0;
  }
}

export function compareSortKeys(left: SortKey, right: SortKey, ascending: boolean, caseSensitive: boolean,
  order: TextOrder = ordinalOrder): number {
  const rank = sortRank(left.kind, ascending) - sortRank(right.kind, ascending);
  if (rank !== 0) return rank;
  if (left.kind === 'blank') return left.index - right.index;
  const direction = ascending ? 1 : -1;
  const value = left.kind === 'number' ? left.number === right.number ? 0 : left.number < right.number ? -1 : 1 :
    left.kind === 'text' ? order(left.text, right.text, caseSensitive) :
    left.kind === 'boolean' ? (left.flag === right.flag ? 0 : left.flag ? 1 : -1) :
    ordinalOrder(left.text, right.text, true);
  return value !== 0 ? value * direction : left.index - right.index;
}

function sameExtent(left: CellExtent, right: CellExtent) {
  return left.row === right.row && left.column === right.column &&
    left.rows === right.rows && left.columns === right.columns;
}

function extentOf(area: CellExtent): CellExtent {
  return { row: area.row, column: area.column, rows: area.rows, columns: area.columns };
}

function replaceSheet(document: WorkbookDocument, next: SheetDocument): WorkbookDocument {
  return { ...document, sheets: document.sheets.map(sheet => sheet.id === next.id ? next : sheet) };
}

function sheetById(document: WorkbookDocument, sheetId: string): SheetDocument | undefined {
  return document.sheets.find(sheet => sheet.id === sheetId);
}

export function setFilter(basis: WorkbookSnapshot, sheetId: string, filter: SheetFilter | null): StructureCommand {
  const sheet = sheetById(basis.document, sheetId);
  requireDocument(sheet, 'Sheet no longer exists.');
  const { filter: _prior, ...rest } = sheet;
  return {
    kind: 'structure',
    label: filter === null ? 'Remove filter' : 'Filter',
    document: replaceSheet(basis.document, { ...rest, ...(filter !== null && { filter }) }),
  };
}

/**
 * Sort complete records by one column. Blanks stay last in both directions, equal keys keep their
 * original order, and moving formulas shift their relative references by the row delta.
 */
export function sortData(basis: WorkbookSnapshot, range: CellExtent & { readonly sheetId: string }, keyColumn: number,
  hasHeader: boolean, ascending: boolean, caseSensitive: boolean, order: TextOrder = ordinalOrder): StructureCommand {
  validateExtent(range);
  const sheet = sheetById(basis.document, range.sheetId);
  requireDocument(sheet, 'The worksheet no longer exists.');
  requireDocument(!sheet.cells.some(cell => extentContains(range, cell) &&
    (cell.spillOwner || cell.content.kind === 'python' && cell.content.block?.spillActive)),
  'Cannot sort a Python spill. Sort the DataFrame in its source or insert a values copy first.');
  requireDocument(keyColumn >= range.column && keyColumn <= lastColumn(range),
    'The sort key must be a column inside the selection.');
  let work = extentOf(range);
  for (const table of basis.document.names) {
    if (!table.isTable || table.area.invalid) continue;
    const extent = extentOf(table.area);
    if (!overlaps(extent, work)) continue;
    const design = tableDesign(table);
    if (design.totalRow && lastRow(work) >= lastRow(extent)) {
      if (sameExtent(work, extent)) work = extentOf(tableInputExtent(table));
      else throw new Error("Exclude the table's total row from sorting.");
    }
    const firstData = hasHeader ? work.row + 1 : work.row;
    if (firstData < extent.row + (design.headerRow ? 1 : 0) || lastRow(work) > lastRow(extent) ||
        work.column !== extent.column || lastColumn(work) !== lastColumn(extent)) {
      throw new Error('Sort complete table records across all table columns, excluding its header and cells outside its extent.');
    }
  }
  const first = hasHeader ? work.row + 1 : work.row;
  requireDocument(first <= lastRow(work), 'There are no records to sort under the header.');
  requireDocument(lastRow(work) - first + 1 <= documentLimits.cells, 'Sort selection is too large.');
  requireDocument(!sheet.merges.some(merge => overlaps(merge, work)), 'Unmerge cells before sorting this range.');
  const rows = Array.from({ length: lastRow(work) - first + 1 }, (_, index) => first + index);
  const located = new Map(sheet.cells.map(cell => [cellKey(cell), cell]));
  const ordered = rows.map((row, index) => ({
    row, key: sortKeyFor(located.get(cellKey({ row, column: keyColumn })), index),
  })).sort((left, right) => compareSortKeys(left.key, right.key, ascending, caseSensitive, order))
    .map(item => item.row);
  const block = new Map<number, CellDocument>();
  const rest: CellDocument[] = [];
  for (const cell of sheet.cells) {
    if (cell.row >= first && cell.row <= lastRow(work) && cell.column >= work.column && cell.column <= lastColumn(work)) {
      block.set(cellKey(cell), cell);
    } else rest.push(cell);
  }
  const placed: CellDocument[] = [];
  for (let index = 0; index < rows.length; index++) {
    const source = ordered[index]!;
    const destination = rows[index]!;
    for (let column = work.column; column <= lastColumn(work); column++) {
      const cell = block.get(cellKey({ row: source, column }));
      if (!cell) continue;
      const content = destination === source ? cell.content : cell.content.kind === 'formula' ?
        { kind: 'formula' as const, source: copyReferences(cell.content.source, destination - source, 0) } :
        cell.content.kind === 'python' ?
          rewritePythonContent(cell.content, reference => copyReferences(reference, destination - source, 0)) :
        cell.content;
      placed.push(destination === source && content === cell.content ? cell : { ...cell, row: destination, content });
    }
  }
  let updated: SheetDocument = { ...sheet, cells: [...rest, ...placed] };
  const filter = sheet.filter;
  if (filter && work.row <= lastRow(filter) && lastRow(work) >= filter.row &&
      work.column <= lastColumn(filter) && lastColumn(work) >= filter.column) {
    const matches = work.column <= filter.column && lastColumn(work) >= lastColumn(filter) &&
      keyColumn >= filter.column && keyColumn <= lastColumn(filter) &&
      lastRow(work) === lastRow(filter) && (hasHeader ? work.row === filter.row : work.row === filter.row + 1);
    const { sort: _sort, ...withoutSort } = filter;
    updated = { ...updated, filter: matches ? { ...filter, sort: { column: keyColumn, ascending } } : withoutSort };
  }
  return { kind: 'structure', label: 'Sort', document: replaceSheet(basis.document, updated) };
}

/** Rows hidden by filter criteria. Manually hidden dimensions are not part of this mask. */
export function filteredRows(sheet: SheetDocument, filter: SheetFilter | undefined = sheet.filter,
  exceptColumn?: number): RowMask {
  if (!filter) return RowMask.none();
  const criteria = filter.criteria.filter(criterion => criterion.column !== exceptColumn);
  if (!criteria.length) return RowMask.none();
  const from = filter.row + 1;
  const to = lastRow(filter);
  const columns = new Set(criteria.map(criterion => criterion.column));
  const rules = criteria.map(criterion => ({ column: criterion.column, allowed: new Set(criterion.allowed) }));
  const cells = [];
  for (const cell of sheet.cells) {
    if (cell.row < from || cell.row > to || !columns.has(cell.column)) continue;
    cells.push({ row: cell.row, column: cell.column, key: filterKey(cell) });
  }
  return maskRows(from, to, rules, filterKey(undefined), cells);
}

export function recordSummary(sheet: SheetDocument): string | undefined {
  const filter = sheet.filter;
  if (!filter || filter.rows < 2) return undefined;
  const from = filter.row + 1;
  const to = lastRow(filter);
  const filtered = filteredRows(sheet);
  let manualOnly = 0;
  for (const row of sheet.rows) {
    if (!row.hidden || row.index < from || row.index > to || filtered.has(row.index)) continue;
    manualOnly++;
  }
  const visible = (to - from + 1) - filtered.countIn(from, to) - manualOnly;
  return visible + ' of ' + (filter.rows - 1) + ' records';
}

export function looksLikeHeader(basis: WorkbookSnapshot, range: CellExtent & { readonly sheetId: string }): boolean {
  if (range.rows < 2) return false;
  const sheet = sheetById(basis.document, range.sheetId);
  if (!sheet) return false;
  for (let column = range.column; column <= lastColumn(range); column++) {
    const header = findCell(sheet, { row: range.row, column });
    const below = findCell(sheet, { row: range.row + 1, column });
    if (header?.content.kind === 'text' && header.content.value.length > 0 &&
        (below?.content.kind === 'number' || below?.content.kind === 'boolean')) return true;
  }
  return false;
}

function isPopulated(cell: CellDocument): boolean {
  return cell.content.kind !== 'blank' && !(cell.content.kind === 'text' && cell.content.value.length === 0);
}

export function currentRegion(document: WorkbookDocument, seed: CellExtent & { readonly sheetId: string },
  maxRow = worksheetViewport.rows, maxColumn = worksheetViewport.columns): CellExtent & { readonly sheetId: string } {
  validateExtent(seed);
  const sheet = sheetById(document, seed.sheetId);
  requireDocument(sheet, 'The worksheet no longer exists.');
  const populated = sheet.cells.filter(isPopulated);
  const filled = new Set(populated.map(cell => cellKey(cell)));
  const seedBottom = lastRow(seed);
  const seedRight = lastColumn(seed);
  const queue: { row: number; column: number }[] = [];
  for (const cell of populated) {
    if (cell.row >= seed.row && cell.row <= seedBottom && cell.column >= seed.column && cell.column <= seedRight) {
      queue.push({ row: cell.row, column: cell.column });
    }
  }
  if (!queue.length) return seed;
  const seen = new Set<number>();
  const sameSheet = (row: number, column: number) => filled.has(cellKey({ row, column })) && !seen.has(cellKey({ row, column }));
  let top = Number.MAX_SAFE_INTEGER, left = Number.MAX_SAFE_INTEGER, bottom = 0, right = 0;
  while (queue.length) {
    const { row, column } = queue.shift()!;
    if (!seen.add(cellKey({ row, column }))) continue;
    top = Math.min(top, row); left = Math.min(left, column);
    bottom = Math.max(bottom, row); right = Math.max(right, column);
    for (const [nr, nc] of [[row - 1, column], [row + 1, column], [row, column - 1], [row, column + 1]] as const) {
      if (nr >= 1 && nc >= 1 && nr <= maxRow && nc <= maxColumn && sameSheet(nr, nc)) queue.push({ row: nr, column: nc });
    }
  }
  return { sheetId: seed.sheetId, row: top, column: left, rows: bottom - top + 1, columns: right - left + 1 };
}

export function filterRange(document: WorkbookDocument, selection: CellExtent & { readonly sheetId: string },
  maxRow = worksheetViewport.rows, maxColumn = worksheetViewport.columns): CellExtent & { readonly sheetId: string } {
  validateExtent(selection);
  const table = document.names.find(name => name.isTable && !name.area.invalid && name.area.sheetId === selection.sheetId &&
    extentContains(name.area, { row: selection.row, column: selection.column }) &&
    extentContains(name.area, { row: lastRow(selection), column: lastColumn(selection) }));
  if (table) {
    const input = tableInputExtent(table);
    return { sheetId: input.sheetId, row: input.row, column: input.column, rows: input.rows, columns: input.columns };
  }
  return selection.rows * selection.columns === 1 ? currentRegion(document, selection, maxRow, maxColumn) : selection;
}

export interface FilterValue { readonly key: string; readonly label: string }

/** Distinct keys of one column, excluding rows another column already hides. */
export function columnValues(sheet: SheetDocument, filter: SheetFilter, column: number,
  order: TextOrder = ordinalOrder): readonly FilterValue[] {
  const from = filter.row + 1;
  const to = lastRow(filter);
  const excluded = filteredRows(sheet, filter, column);
  const keys = new Set<string>();
  const list: { key: string; label: string; sort: SortKey }[] = [];
  let populatedVisible = 0;
  for (const cell of sheet.cells) {
    if (cell.column !== column || cell.row < from || cell.row > to || excluded.has(cell.row)) continue;
    populatedVisible++;
    const key = filterKey(cell);
    if (keys.has(key)) continue;
    keys.add(key);
    list.push({ key, label: key.length === 0 ? '(Blanks)' : key, sort: sortKeyFor(cell, cell.row) });
  }
  const visible = (to - from + 1) - excluded.countIn(from, to);
  const blank = filterKey(undefined);
  if (visible > populatedVisible && !keys.has(blank)) {
    list.push({ key: blank, label: '(Blanks)', sort: sortKeyFor(undefined, from) });
  }
  list.sort((left, right) => compareSortKeys(left.sort, right.sort, true, false, order));
  return list.map(({ key, label }) => ({ key, label }));
}

export function headerLabel(sheet: SheetDocument, row: number, column: number): string {
  const text = filterKey(sheet.cells.find(cell => cell.row === row && cell.column === column));
  return text.length === 0 ? 'Column ' + columnLetters(column) : text;
}

export function rangeLabel(range: CellExtent): string {
  const first = columnLetters(range.column) + range.row;
  const last = columnLetters(lastColumn(range)) + lastRow(range);
  return range.rows === 1 && range.columns === 1 ? first : first + ':' + last;
}
