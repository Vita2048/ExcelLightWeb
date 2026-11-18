import type { CellExtent, Coordinate, TableDesign, WorkbookDocument, WorkbookName } from './workbook';

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
