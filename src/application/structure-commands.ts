import { axisLimit, shiftBox, shiftIndex, validateAxisEdit, type Axis, type AxisEdit } from '../core/axis-edit';
import { intersection, rewriteSheetReferences, sameSheet, shiftReferences, type ReferenceBox } from '../core/formula-references';
import { documentLimits, emptySheet, type AxisDimension, type CellDocument, type CellExtent,
  type IdFactory, type SheetDocument, type WorkbookName } from '../core/workbook';
import { ordinalNameKey, requireDocument } from '../core/workbook-validation';
import { lastColumn, lastRow, ownsTableFilter, tableDesign, tablesOnSheet } from '../core/workbook-names';
import { rewritePythonContent } from '../python/references';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { StructureCommand } from './workbook-session';

export type SheetChange =
  | { kind: 'add'; after: boolean }
  | { kind: 'rename'; name: string }
  | { kind: 'delete' }
  | { kind: 'place'; before: string; copy: boolean };

export function changeSheet(basis: WorkbookSnapshot, id: string, change: SheetChange, makeId: IdFactory) {
  const { document } = basis;
  let sheets = [...document.sheets];
  let names = document.names;
  const index = sheets.findIndex(sheet => sheet.id === id);
  requireDocument(index >= 0, 'The worksheet no longer exists.');
  const target = sheets[index]!;
  let active = id;
  let label: string;
  if (change.kind === 'delete') {
    requireDocument(sheets.length > 1, 'A workbook must retain at least one worksheet.');
    sheets.splice(index, 1);
    active = sheets[Math.min(index, sheets.length - 1)]!.id;
    sheets = sheets.map(sheet => rewriteCells(sheet, source => rewriteSheetReferences(source, target.name)));
    names = names.map(name => name.area.sheetId === id ? { ...name, area: { ...name.area, invalid: true } } : name);
    label = 'Delete Sheet';
  } else if (change.kind === 'rename') {
    const name = change.name.trim();
    requireDocument(name.length > 0 && name.length <= documentLimits.sheetNameLength &&
      !sheets.some(sheet => sheet.id !== id && sameSheet(sheet.name, name)), 'Enter a unique worksheet name of 1–31 characters.');
    if (name !== target.name) sheets = sheets.map(sheet => {
      const next = rewriteCells(sheet, source => rewriteSheetReferences(source, target.name, name));
      return sheet.id === id ? { ...next, name } : next;
    });
    label = 'Rename Sheet';
  } else if (change.kind === 'place' && !change.copy) {
    if (change.before !== id) {
      sheets.splice(index, 1);
      const before = change.before ? sheets.findIndex(sheet => sheet.id === change.before) : sheets.length;
      requireDocument(before >= 0, 'The destination worksheet no longer exists.');
      sheets.splice(before, 0, target);
    }
    label = 'Move Sheet';
  } else {
    requireDocument(sheets.length < documentLimits.sheets, 'The workbook already has the maximum number of worksheets.');
    let n = 1;
    while (sheets.some(sheet => sameSheet(sheet.name, 'Sheet' + n))) n++;
    const name = 'Sheet' + n;
    active = makeId();
    let next = emptySheet(active, name);
    let before = index;
    if (change.kind === 'add') {
      before += change.after ? 1 : 0;
      label = 'Insert Sheet';
    } else {
      before = change.before ? sheets.findIndex(sheet => sheet.id === change.before) : sheets.length;
      requireDocument(before >= 0, 'The destination worksheet no longer exists.');
      const copy = rewriteCells(target, source => rewriteSheetReferences(source, target.name, name));
      next = { ...copy, id: active, name, cells: copy.cells.map(cell => {
        let content = cell.content;
        if (content.kind === 'python') {
          const blockId = makeId();
          content = { kind: 'python', blockId, source: content.source,
            ...(content.block && { block: { ...content.block, id: blockId } }) };
        }
        return { ...cell, id: makeId(), content };
      }) };
      const used = new Set(names.map(entry => ordinalNameKey(entry.name)));
      const copies = names.filter(entry => entry.area.sheetId === id).map(entry => {
        let suffix = 2;
        let candidate: string;
        do { candidate = entry.name.slice(0, 254 - String(suffix).length) + '_' + suffix++; }
        while (used.has(ordinalNameKey(candidate)));
        used.add(ordinalNameKey(candidate));
        return { ...entry, id: makeId(), name: candidate, area: { ...entry.area, sheetId: active } };
      });
      names = [...names, ...copies];
      label = 'Copy Sheet';
    }
    sheets.splice(before, 0, next);
  }
  return { active, command: { kind: 'structure', label, document: { ...document, sheets, names } } satisfies StructureCommand };
}

function rewriteCell(cell: CellDocument, rewrite: (source: string) => string): CellDocument {
  if (cell.content.kind === 'python') {
    const content = rewritePythonContent(cell.content, rewrite);
    return content === cell.content ? cell : { ...cell, content };
  }
  if (cell.content.kind !== 'formula') return cell;
  const source = rewrite(cell.content.source);
  if (source === cell.content.source) return cell;
  const { cache: _cache, ...rest } = cell;
  return { ...rest, content: { ...cell.content, source } };
}

function rewriteCells(sheet: SheetDocument, rewrite: (source: string) => string): SheetDocument {
  const cells = sheet.cells.map(cell => rewriteCell(cell, rewrite));
  return cells.some((cell, index) => cell !== sheet.cells[index]) ? { ...sheet, cells } : sheet;
}
function box(extent: CellExtent): ReferenceBox {
  return { top: extent.row, left: extent.column, bottom: extent.row + extent.rows - 1, right: extent.column + extent.columns - 1 };
}
function extent(box: ReferenceBox): CellExtent {
  return { row: box.top, column: box.left, rows: box.bottom - box.top + 1, columns: box.right - box.left + 1 };
}
function rangeShift(edit: AxisEdit, range: CellExtent) {
  const parts = shiftBox(edit, box(range));
  return parts?.length === 1 ? extent(parts[0]!) : undefined;
}

function shiftName(name: WorkbookName, id: string, edit: AxisEdit): WorkbookName {
  if (name.area.sheetId !== id || name.area.invalid) return name;
  const area = rangeShift(edit, name.area);
  const original = box(name.area);
  const intersectsLane = !edit.lane || original.left <= edit.lane[1] && original.right >= edit.lane[0];
  const removedHeader = name.isTable && name.design?.headerRow !== false && edit.axis === 'row' &&
    edit.delete && intersectsLane && name.area.row >= edit.start && name.area.row < edit.start + edit.count;
  if (!area || removedHeader) return { ...name, area: { ...name.area, invalid: true } };
  let next = { ...name, area: { ...name.area, ...area } };
  if (name.isTable && name.design?.headerRow === false && edit.axis === 'column' && name.headerLabels &&
      (!edit.lane || original.top >= edit.lane[0] && original.bottom <= edit.lane[1])) {
    const mapped = new Map(name.headerLabels.map((label, i) => [shiftIndex(edit, original.left + i), label]));
    const used = new Set([...mapped.values()].map(ordinalNameKey));
    const headerLabels = Array.from({ length: area.columns }, (_, i) => {
      const existing = mapped.get(area.column + i);
      if (existing !== undefined) return existing;
      const base = 'Column' + (i + 1);
      let label = base, suffix = 2;
      while (used.has(ordinalNameKey(label))) label = base + suffix++;
      used.add(ordinalNameKey(label));
      return label;
    });
    next = { ...next, headerLabels };
  }
  return next;
}

export function editAxis(basis: WorkbookSnapshot, id: string, edit: AxisEdit): StructureCommand {
  validateAxisEdit(edit);
  const target = basis.document.sheets.find(sheet => sheet.id === id);
  requireDocument(target, 'The worksheet no longer exists.');
  const key = edit.axis === 'row' ? 'rows' : 'columns';
  const other = edit.axis === 'row' ? 'column' : 'row';
  const affected = edit.axis === 'row'
    ? { top: edit.start, bottom: axisLimit(edit.axis), left: edit.lane?.[0] ?? 1, right: edit.lane?.[1] ?? documentLimits.columns }
    : { left: edit.start, right: axisLimit(edit.axis), top: edit.lane?.[0] ?? 1, bottom: edit.lane?.[1] ?? documentLimits.rows };
  requireDocument(!edit.lane || !target.merges.some(merge => intersection(box(merge), affected)),
    'Unmerge cells before shifting this lane.');
  const tables = tablesOnSheet(basis.document, id);
  for (const table of tables) {
    if (!intersection(box(table.area), affected)) continue;
    const vertical = edit.axis === 'row';
    const first = vertical ? table.area.column : table.area.row;
    const last = vertical ? lastColumn(table.area) : lastRow(table.area);
    requireDocument(!edit.lane || (edit.lane[0] <= first && edit.lane[1] >= last),
      'Shift the full width or height of the table to keep its records together.');
    requireDocument(edit.delete || rangeShift(edit, table.area), 'Insertion would move a table beyond the worksheet boundary.');
    const design = tableDesign(table);
    if (edit.delete && vertical && design.totalRow) {
      const end = edit.start + edit.count - 1;
      requireDocument(end < lastRow(table.area) || edit.start > lastRow(table.area) || edit.start <= table.area.row,
        'Turn off the Total Row before deleting it.');
    }
  }
  const names = basis.document.names.map(name => shiftName(name, id, edit));
  function moved(index: number): number | undefined {
    const next = shiftIndex(edit, index);
    requireDocument(next !== undefined || edit.delete, 'Insertion would discard authored cells or dimensions at the worksheet boundary.');
    return next;
  }
  const sheets = basis.document.sheets.map(sheet => {
    if (sheet.id !== id) return rewriteCells(sheet, source => shiftReferences(source, sheet.name, target.name, edit));
    const cells = sheet.cells.flatMap(cell => {
      const hit = !edit.lane || cell[other] >= edit.lane[0] && cell[other] <= edit.lane[1];
      const index = hit ? moved(cell[edit.axis]) : cell[edit.axis];
      if (index === undefined) return [];
      const next = rewriteCell(cell, source => shiftReferences(source, sheet.name, target.name, edit));
      return index === cell[edit.axis] ? [next] : [{ ...next, [edit.axis]: index }];
    });
    const dimensions = sheet[key].flatMap(dimension => {
      if (edit.lane) return [dimension];
      const index = moved(dimension.index);
      return index === undefined ? [] : [{ ...dimension, index }];
    });
    const merges = sheet.merges.flatMap(merge => {
      if (edit.lane) return [merge];
      const next = rangeShift(edit, merge);
      requireDocument(next || edit.delete, 'Insertion would discard a merged range at the worksheet boundary.');
      return next && next.rows * next.columns > 1 ? [next] : [];
    });
    const { filter, ...rest } = sheet;
    const owner = tables.find(table => ownsTableFilter(filter, table));
    if (edit.lane && !owner) return { ...sheet, cells };
    if (owner && names.find(name => name.id === owner.id)?.area.invalid) return { ...rest, cells, merges, [key]: dimensions };
    const nextFilter = filter && rangeShift(edit, filter);
    if (!filter || !nextFilter) return { ...rest, cells, merges, [key]: dimensions };
    const { sort, criteria } = filter;
    const shiftColumns = edit.axis === 'column' && (!edit.lane || edit.lane[0] <= filter.row && edit.lane[1] >= lastRow(filter));
    const column = sort && (shiftColumns ? shiftIndex(edit, sort.column) : sort.column);
    const nextCriteria = !shiftColumns ? criteria : criteria.flatMap(criterion => {
      const column = shiftIndex(edit, criterion.column);
      return column === undefined ? [] : [{ ...criterion, column }];
    });
    return { ...rest, cells, merges, [key]: dimensions, filter: {
      ...nextFilter, criteria: nextCriteria, ...(sort && column !== undefined && { sort: { ...sort, column } }),
    } };
  });
  return { kind: 'structure', document: { ...basis.document, sheets, names },
    label: (edit.delete ? 'Delete ' : 'Insert ') + (edit.lane ? 'cells' : edit.axis + 's') };
}

export function changeDimensions(basis: WorkbookSnapshot, id: string, axis: Axis,
  changes: readonly AxisDimension[]): StructureCommand {
  const target = basis.document.sheets.find(sheet => sheet.id === id);
  requireDocument(target, 'The worksheet no longer exists.');
  const key = axis === 'row' ? 'rows' : 'columns';
  const dimensions = new Map(target[key].map(dimension => [dimension.index, dimension]));
  for (const patch of changes) {
    requireDocument(Number.isInteger(patch.index) && patch.index >= 1 && patch.index <= axisLimit(axis), 'Invalid dimension.');
    const next = { ...dimensions.get(patch.index), ...patch };
    if (next.size === undefined && !next.hidden && !next.style) dimensions.delete(patch.index);
    else dimensions.set(patch.index, next);
  }
  const sheets = basis.document.sheets.map(sheet => sheet.id === id
    ? { ...sheet, [key]: [...dimensions.values()].sort((a, b) => a.index - b.index) } : sheet);
  return { kind: 'structure', document: { ...basis.document, sheets }, label: 'Size / visibility' };
}
