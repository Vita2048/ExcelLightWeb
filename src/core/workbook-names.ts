import type { StructuredReference } from './formula-references';
import { findCell, type CellExtent, type Coordinate, type TableDesign, type WorkbookDocument, type WorkbookName } from './workbook';
import { ordinalNameKey } from './workbook-validation';

export const tableStyles = ['Blue', 'Green', 'Orange', 'Purple'] as const;
export const newTableDesign: Readonly<Required<TableDesign>> = Object.freeze({
  style: 'Blue', headerRow: true, totalRow: false, firstColumn: false, lastColumn: false,
  bandedRows: true, bandedColumns: false, filterButton: true,
});

export function tableDesign(name: WorkbookName): Required<TableDesign> {
  const design = name.design;
  return { ...newTableDesign, style: 'Neutral', ...design,
    bandedRows: design?.bandedRows ?? true, bandedColumns: design?.bandedColumns ?? false,
    firstColumn: design?.firstColumn ?? false, lastColumn: design?.lastColumn ?? false,
    filterButton: design?.filterButton ?? true };
}

export const lastRow = (area: CellExtent) => area.row + area.rows - 1;
export const lastColumn = (area: CellExtent) => area.column + area.columns - 1;

export function extentContains(area: CellExtent, point: Coordinate): boolean {
  return point.row >= area.row && point.row <= lastRow(area) &&
    point.column >= area.column && point.column <= lastColumn(area);
}

export function tableAt(tables: readonly WorkbookName[], point: Coordinate): WorkbookName | undefined {
  return tables.find(table => extentContains(table.area, point));
}

export function tablesOnSheet(document: WorkbookDocument, sheetId: string): readonly WorkbookName[] {
  return document.names.filter(name => name.isTable && !name.area.invalid && name.area.sheetId === sheetId);
}

export function tableInputExtent(table: WorkbookName) {
  const { headerRow, totalRow } = tableDesign(table);
  return { ...table.area, rows: table.area.rows - Number(totalRow), headerRow };
}

export function ownsTableFilter(filter: CellExtent | undefined, table: WorkbookName): boolean {
  const input = tableInputExtent(table);
  return !!filter && input.headerRow && filter.row === input.row && filter.column === input.column &&
    filter.rows === input.rows && filter.columns === input.columns;
}

export function resolveWorkbookName(document: WorkbookDocument, reference: string): WorkbookName['area'] | undefined {
  const match = /^([a-z_][a-z_0-9]*)(\[#(?:all|data)\])?$/iu.exec(reference.trim());
  if (!match) return;
  const selector = match[2]?.toUpperCase();
  const entry = document.names.find(name => name.name.toUpperCase() === match[1]!.toUpperCase());
  if (!entry || selector && !entry.isTable) return;
  if (entry.area.invalid || selector !== '[#DATA]') return entry.area;
  const input = tableInputExtent(entry);
  return { ...entry.area, row: input.row + Number(input.headerRow), rows: input.rows - Number(input.headerRow) };
}

function tableColumnLabels(document: WorkbookDocument, table: WorkbookName): readonly string[] {
  const design = tableDesign(table);
  if (!design.headerRow) return table.headerLabels ?? [];
  const sheet = document.sheets.find(item => item.id === table.area.sheetId);
  if (!sheet) return [];
  const labels: string[] = [];
  for (let col = table.area.column; col <= lastColumn(table.area); col++) {
    const cell = findCell(sheet, { row: table.area.row, column: col });
    if (cell?.content.kind === 'text') labels.push(cell.content.value);
    else if (cell?.content.kind === 'number') labels.push(String(cell.content.value));
    else if (cell?.content.kind === 'boolean') labels.push(cell.content.value ? 'TRUE' : 'FALSE');
    else labels.push('');
  }
  return labels;
}

export function resolveStructuredReference(
  document: WorkbookDocument,
  ref: StructuredReference,
  formulaLocation?: { readonly sheetId: string; readonly row: number; readonly column: number },
): WorkbookName['area'] | undefined {
  const table = ref.table
    ? document.names.find(name => name.isTable && ordinalNameKey(name.name) === ordinalNameKey(ref.table!))
    : formulaLocation
      ? tableAt(tablesOnSheet(document, formulaLocation.sheetId), formulaLocation)
      : undefined;
  if (!table) return undefined;
  if (table.area.invalid) return table.area;

  const design = tableDesign(table);
  const top = table.area.row;
  const bottom = lastRow(table.area);
  const dataTop = top + Number(design.headerRow);
  const dataBottom = bottom - Number(design.totalRow);

  const has = (spec: StructuredReference['specifiers'][number]) => ref.specifiers.includes(spec);
  const hasAll = has('#All');
  const hasHeaders = has('#Headers');
  const hasTotals = has('#Totals');
  const hasThisRow = has('@') || has('#This Row');
  const hasData = has('#Data') || ref.specifiers.length === 0;

  if (hasHeaders && !design.headerRow) return undefined;
  if (hasTotals && !design.totalRow) return undefined;
  if (hasHeaders && hasTotals && !has('#Data')) return undefined;

  let row: number;
  let rows: number;
  if (hasThisRow) {
    if (!formulaLocation || formulaLocation.sheetId !== table.area.sheetId ||
        formulaLocation.row < dataTop || formulaLocation.row > dataBottom) {
      return undefined;
    }
    row = formulaLocation.row;
    rows = 1;
  } else if (hasAll) {
    row = top;
    rows = table.area.rows;
  } else if (hasHeaders && has('#Data') && hasTotals) {
    row = top;
    rows = table.area.rows;
  } else if (hasHeaders && has('#Data')) {
    row = top;
    rows = dataBottom - top + 1;
  } else if (has('#Data') && hasTotals) {
    row = dataTop;
    rows = bottom - dataTop + 1;
  } else if (hasHeaders) {
    row = top;
    rows = 1;
  } else if (hasTotals) {
    row = bottom;
    rows = 1;
  } else if (hasData) {
    row = dataTop;
    rows = dataBottom - dataTop + 1;
  } else {
    return undefined;
  }

  if (rows < 1) return undefined;

  let column = table.area.column;
  let columns = table.area.columns;
  if (ref.startColumn !== undefined) {
    const labels = tableColumnLabels(document, table);
    const startIdx = labels.findIndex(label => ordinalNameKey(label) === ordinalNameKey(ref.startColumn!));
    if (startIdx < 0) return undefined;
    const endIdx = ref.endColumn !== undefined
      ? labels.findIndex(label => ordinalNameKey(label) === ordinalNameKey(ref.endColumn!))
      : startIdx;
    if (endIdx < 0) return undefined;
    const left = Math.min(startIdx, endIdx);
    const right = Math.max(startIdx, endIdx);
    column = table.area.column + left;
    columns = right - left + 1;
  }

  return { sheetId: table.area.sheetId, row, column, rows, columns };
}

export function migrateDetachedHeaders(document: WorkbookDocument): WorkbookDocument {
  if (!document.names || document.names.length === 0) return document;
  const migrated = document.names.map(name => {
    if (!name || !name.area || name.area.invalid || !name.isTable) return name;
    const sheet = document.sheets.find(s => s.id === name.area.sheetId);
    if (!sheet) return name;
    const design = tableDesign(name);
    if (design.headerRow || (name.headerLabels && name.headerLabels.length > 0)) return name;
    try {
      const labels: string[] = [];
      const seen = new Set<string>();
      for (let col = name.area.column; col <= lastColumn(name.area); col++) {
        const cell = findCell(sheet, { row: name.area.row, column: col });
        if (!cell || cell.content.kind !== 'text') throw new Error('Table headers must be nonblank unique text.');
        const val = cell.content.value;
        if (!val || !val.trim()) throw new Error('Table headers must be nonblank unique text.');
        const key = val.toUpperCase();
        if (seen.has(key)) throw new Error('Table headers must be nonblank unique text.');
        seen.add(key);
        labels.push(val);
      }
      return name.area.rows > 1
        ? { ...name, area: { ...name.area, row: name.area.row + 1, rows: name.area.rows - 1 }, headerLabels: labels }
        : { ...name, area: { ...name.area, invalid: true }, headerLabels: labels };
    } catch {
      return { ...name, area: { ...name.area, invalid: true } };
    }
  });
  return { ...document, names: migrated };
}
