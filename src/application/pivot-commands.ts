import { collectFieldCollapsibleKeys, evaluatePivotTable, extractPivotCellDrillDown } from '../calculation/pivot-engine';
import {
  nextPivotName,
  type PivotAxisField,
  type PivotFilterField,
  type PivotReportLayout,
  type PivotSource,
  type PivotStyleId,
  type PivotSubtotalPosition,
  type PivotTableDocument,
  type PivotValueField,
} from '../core/pivot';
import {
  cellKey,
  documentLimits,
  emptySheet,
  type CellDocument,
  type CellExtent,
  type CellReference,
  type DocumentId,
  type IdFactory,
  type WorkbookDocument,
} from '../core/workbook';
import { ordinalNameKey, overlaps, requireDocument } from '../core/workbook-validation';
import { createTable, deleteTableSelection, nextTableName, parseNameExtent } from './table-commands';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { StructureCommand } from './workbook-session';

export interface CreatePivotInput {
  readonly source: PivotSource;
  readonly target: CellReference;
  readonly name?: string;
  readonly filters?: readonly PivotFilterField[];
  readonly columns?: readonly PivotAxisField[];
  readonly rows?: readonly PivotAxisField[];
  readonly values?: readonly PivotValueField[];
  readonly valuesInRows?: boolean;
  readonly reportLayout?: PivotReportLayout;
  readonly subtotals?: PivotSubtotalPosition;
  readonly grandTotals?: { readonly rows: boolean; readonly columns: boolean };
  readonly repeatLabels?: boolean;
  readonly blankRows?: boolean;
  readonly showButtons?: boolean;
  readonly showFieldHeaders?: boolean;
  readonly emptyCellText?: string;
  readonly autoFitColumns?: boolean;
  readonly style?: PivotStyleId;
  readonly styleOptions?: {
    readonly rowHeaders: boolean;
    readonly columnHeaders: boolean;
    readonly bandedRows: boolean;
    readonly bandedColumns: boolean;
  };
}

function revertPivotCellsOnSheet(document: WorkbookDocument, pivotId: DocumentId): WorkbookDocument {
  const pivots = (document.pivots ?? []).filter(item => item.id !== pivotId);
  const charts = (document.charts ?? []).filter(chart => !(chart.source.kind === 'pivot' && chart.source.pivotId === pivotId));
  const { pivots: _prev, charts: _prevCharts, ...restDoc } = document;
  const sheets = document.sheets.map(sheet => {
    const cells = sheet.cells.filter(cell => !(cell.spillKind === 'pivot' && cell.spillOwner === pivotId));
    return cells.length === sheet.cells.length ? sheet : { ...sheet, cells };
  });
  return {
    ...restDoc,
    sheets,
    ...(pivots.length > 0 && { pivots }),
    ...(charts.length > 0 && { charts }),
  };
}

export function materializePivotDocument(
  document: WorkbookDocument,
  pivot: PivotTableDocument,
  makeId: IdFactory,
): WorkbookDocument {
  const sheetIndex = document.sheets.findIndex(sheet => sheet.id === pivot.target.sheetId);
  requireDocument(sheetIndex >= 0, 'PivotTable target worksheet no longer exists.');
  const evaluated = evaluatePivotTable(document, pivot);
  const extent: CellExtent = {
    row: pivot.target.row,
    column: pivot.target.column,
    rows: evaluated.rows,
    columns: evaluated.columns,
  };
  requireDocument(
    extent.row + extent.rows - 1 <= documentLimits.rows &&
    extent.column + extent.columns - 1 <= documentLimits.columns &&
    extent.rows * extent.columns <= 100_000,
    'PivotTable exceeds worksheet limits.',
  );

  const sheet = document.sheets[sheetIndex]!;
  requireDocument(!sheet.merges.some(merge => overlaps(merge, extent)), 'PivotTable cannot overlap merged cells.');
  requireDocument(
    !document.names.some(name => name.isTable && !name.area.invalid && name.area.sheetId === sheet.id && overlaps(name.area, extent)),
    'PivotTable cannot overlap a workbook table.',
  );
  requireDocument(
    !(document.pivots ?? []).some(other => other.id !== pivot.id && other.target.sheetId === sheet.id && other.extent && overlaps(other.extent, extent)),
    'PivotTable cannot overlap another PivotTable.',
  );

  const oldOwned = new Map<number, CellDocument>();
  const retained = new Map<number, CellDocument>();
  for (const cell of sheet.cells) {
    if (cell.spillKind === 'pivot' && cell.spillOwner === pivot.id) {
      oldOwned.set(cellKey(cell), cell);
    } else {
      retained.set(cellKey(cell), cell);
    }
  }

  // Collision check against non-owned populated cells, formula spills, and Python spills inside the target rectangle.
  for (let r = extent.row; r < extent.row + extent.rows; r++) {
    for (let c = extent.column; c < extent.column + extent.columns; c++) {
      const existing = retained.get(cellKey({ row: r, column: c }));
      requireDocument(
        !existing || (existing.content.kind === 'blank' && !existing.spillOwner && existing.cache === undefined),
        'PivotTable destination overlaps existing worksheet data or spills.',
      );
    }
  }

  const desiredWidths = new Map<number, number>();
  for (const item of evaluated.cells) {
    const row = extent.row + item.rowOffset;
    const column = extent.column + item.colOffset;
    const key = cellKey({ row, column });
    const existing = retained.get(key);
    const oldCell = oldOwned.get(key);
    const mergedStyle = existing?.style || item.style ? { ...existing?.style, ...item.style } : undefined;
    retained.set(key, {
      id: oldCell?.id ?? existing?.id ?? makeId(),
      row,
      column,
      content: item.content,
      ...(mergedStyle && { style: mergedStyle }),
      ...(item.cache && { cache: item.cache }),
      spillOwner: pivot.id,
      spillKind: 'pivot',
    });

    let charEstimate = 0;
    if (item.content.kind === 'text') charEstimate = item.content.value.length;
    else if (item.content.kind === 'number') {
      const digits = Math.max(1, String(Math.round(Math.abs(item.content.value))).length);
      charEstimate = item.style?.numberFormat ? digits + Math.floor(digits / 3) + 5 : digits + 2;
    } else if (item.cache?.kind === 'error') {
      charEstimate = item.cache.value.length;
    }
    const indentAllowance = (item.style?.indent ?? 0) * 20;
    const controlAllowance = (item.role === 'header' || (item.role === 'filter' && item.colOffset === 1)) ? 24 : 0;
    const needed = Math.min(300, Math.max(96, Math.ceil(charEstimate * 7.8 + indentAllowance + controlAllowance + 18)));
    const prev = desiredWidths.get(column) ?? 96;
    if (needed > prev) desiredWidths.set(column, needed);
  }

  const nextColumns = [...sheet.columns];
  if (pivot.autoFitColumns !== false) {
    for (const [colIndex, width] of desiredWidths) {
      if (width <= 96) continue;
      const existingIdx = nextColumns.findIndex(col => col.index === colIndex);
      if (existingIdx >= 0) {
        const cur = nextColumns[existingIdx]!;
        const curSize = cur.size ?? 96;
        if (width > curSize) nextColumns[existingIdx] = { ...cur, size: width };
      } else {
        nextColumns.push({ index: colIndex, hidden: false, size: width });
      }
    }
    nextColumns.sort((a, b) => a.index - b.index);
  }

  const nextCells = [...retained.values()].sort((a, b) => cellKey(a) - cellKey(b));
  const nextSheets = document.sheets.map((item, idx) => (idx === sheetIndex ? { ...item, cells: nextCells, columns: nextColumns } : item));
  const nextPivot: PivotTableDocument = { ...pivot, extent };
  const existingPivots = document.pivots ?? [];
  const exists = existingPivots.some(item => item.id === pivot.id);
  const nextPivots = exists
    ? existingPivots.map(item => (item.id === pivot.id ? nextPivot : item))
    : [...existingPivots, nextPivot];

  return { ...document, sheets: nextSheets, pivots: nextPivots };
}

export function createPivotTable(
  basis: WorkbookSnapshot,
  input: CreatePivotInput,
  makeId: IdFactory,
): StructureCommand {
  const name = (input.name ?? nextPivotName(basis.document)).trim();
  requireDocument(name.length > 0 && name.length <= 255, 'PivotTable requires a valid name.');
  const nameKey = ordinalNameKey(name);
  requireDocument(
    !(basis.document.pivots ?? []).some(item => ordinalNameKey(item.name) === nameKey),
    'A PivotTable with this name already exists.',
  );

  const pivot: PivotTableDocument = {
    id: makeId(),
    name,
    source: input.source,
    target: input.target,
    filters: input.filters ?? [],
    columns: input.columns ?? [],
    rows: input.rows ?? [],
    values: input.values ?? [],
    ...(input.valuesInRows !== undefined && { valuesInRows: input.valuesInRows }),
    reportLayout: input.reportLayout ?? 'compact',
    subtotals: input.subtotals ?? 'top',
    grandTotals: input.grandTotals ?? { rows: true, columns: true },
    ...(input.repeatLabels !== undefined && { repeatLabels: input.repeatLabels }),
    ...(input.blankRows !== undefined && { blankRows: input.blankRows }),
    showButtons: input.showButtons ?? true,
    showFieldHeaders: input.showFieldHeaders ?? true,
    ...(input.emptyCellText !== undefined && input.emptyCellText !== '' && { emptyCellText: input.emptyCellText }),
    ...(input.autoFitColumns !== undefined && { autoFitColumns: input.autoFitColumns }),
    style: input.style ?? 'Blue',
    styleOptions: input.styleOptions ?? {
      rowHeaders: true,
      columnHeaders: true,
      bandedRows: false,
      bandedColumns: false,
    },
  };

  return {
    kind: 'structure',
    label: 'Insert PivotTable',
    document: materializePivotDocument(basis.document, pivot, makeId),
  };
}

export function updatePivotLayout(
  basis: WorkbookSnapshot,
  pivotId: DocumentId,
  patch: Partial<Omit<PivotTableDocument, 'id' | 'extent'>>,
  makeId: IdFactory,
): StructureCommand {
  const current = (basis.document.pivots ?? []).find(item => item.id === pivotId);
  requireDocument(current, 'PivotTable no longer exists.');
  const updated: PivotTableDocument = { ...current, ...patch, id: current.id };
  return {
    kind: 'structure',
    label: 'Update PivotTable',
    document: materializePivotDocument(basis.document, updated, makeId),
  };
}

export function refreshPivotTable(
  basis: WorkbookSnapshot,
  pivotId: DocumentId,
  makeId: IdFactory,
): StructureCommand {
  const current = (basis.document.pivots ?? []).find(item => item.id === pivotId);
  requireDocument(current, 'PivotTable no longer exists.');
  return {
    kind: 'structure',
    label: 'Refresh PivotTable',
    document: materializePivotDocument(basis.document, current, makeId),
  };
}

export function refreshAllPivotTables(
  basis: WorkbookSnapshot,
  makeId: IdFactory,
): StructureCommand {
  let doc = basis.document;
  for (const pivot of basis.document.pivots ?? []) {
    const latest = (doc.pivots ?? []).find(item => item.id === pivot.id) ?? pivot;
    doc = materializePivotDocument(doc, latest, makeId);
  }
  return {
    kind: 'structure',
    label: 'Refresh All PivotTables',
    document: doc,
  };
}

export function togglePivotCollapse(
  basis: WorkbookSnapshot,
  pivotId: DocumentId,
  key: string,
  makeId: IdFactory,
): StructureCommand {
  const current = (basis.document.pivots ?? []).find(item => item.id === pivotId);
  requireDocument(current, 'PivotTable no longer exists.');
  const collapsed = new Set(current.collapsedKeys ?? []);
  if (collapsed.has(key)) collapsed.delete(key);
  else collapsed.add(key);
  const { collapsedKeys: _prevCollapsed, ...basePivot } = current;
  const updated: PivotTableDocument = {
    ...basePivot,
    ...(collapsed.size > 0 && { collapsedKeys: [...collapsed] }),
  };
  return {
    kind: 'structure',
    label: 'Toggle PivotTable Detail',
    document: materializePivotDocument(basis.document, updated, makeId),
  };
}

export function movePivotTable(
  basis: WorkbookSnapshot,
  pivotId: DocumentId,
  target: CellReference,
  makeId: IdFactory,
): StructureCommand {
  const current = (basis.document.pivots ?? []).find(item => item.id === pivotId);
  requireDocument(current, 'PivotTable no longer exists.');
  // First revert cells at the previous target sheet so cross-sheet or overlapping moves validate cleanly.
  const reverted = revertPivotCellsOnSheet(basis.document, pivotId);
  return {
    kind: 'structure',
    label: 'Move PivotTable',
    document: materializePivotDocument(reverted, { ...current, target }, makeId),
  };
}

export function deletePivotTable(
  basis: WorkbookSnapshot,
  pivotId: DocumentId,
): StructureCommand {
  return {
    kind: 'structure',
    label: 'Delete PivotTable',
    document: revertPivotCellsOnSheet(basis.document, pivotId),
  };
}

export function deletePivotSelection(
  basis: WorkbookSnapshot,
  selections: readonly { readonly sheetId: DocumentId; readonly row: number; readonly column: number; readonly rows: number; readonly columns: number }[],
): StructureCommand | undefined {
  const pivots = basis.document.pivots ?? [];
  if (!pivots.length || !selections.length) return undefined;

  const encloses = (
    sel: { readonly sheetId: DocumentId; readonly row: number; readonly column: number; readonly rows: number; readonly columns: number },
    sheetId: DocumentId,
    ext: CellExtent,
  ) =>
    sel.sheetId === sheetId &&
    sel.row <= ext.row &&
    sel.row + sel.rows - 1 >= ext.row + ext.rows - 1 &&
    sel.column <= ext.column &&
    sel.column + sel.columns - 1 >= ext.column + ext.columns - 1;

  const pivotsToRemove: PivotTableDocument[] = [];
  const pivotExtent = (pivot: PivotTableDocument): CellExtent =>
    pivot.extent ?? { row: pivot.target.row, column: pivot.target.column, rows: 1, columns: 1 };

  for (const pivot of pivots) {
    const ext = pivotExtent(pivot);
    if (selections.some(sel => encloses(sel, pivot.target.sheetId, ext))) {
      pivotsToRemove.push(pivot);
    }
  }
  if (!pivotsToRemove.length) return undefined;

  const removedIds = new Set(pivotsToRemove.map(p => p.id));
  for (const other of pivots) {
    if (removedIds.has(other.id)) continue;
    const ext = pivotExtent(other);
    requireDocument(
      !selections.some(sel => sel.sheetId === other.target.sheetId && overlaps(ext, sel)),
      'Cannot change part of a PivotTable. Select the entire PivotTable to delete it.',
    );
  }

  let doc = basis.document;
  for (const pivot of pivotsToRemove) {
    doc = revertPivotCellsOnSheet(doc, pivot.id);
  }

  for (const sheet of doc.sheets) {
    const sheetSelections = selections.filter(s => s.sheetId === sheet.id);
    if (!sheetSelections.length) continue;
    for (const cell of sheet.cells) {
      const inSel = sheetSelections.some(sel =>
        cell.row >= sel.row && cell.row <= sel.row + sel.rows - 1 &&
        cell.column >= sel.column && cell.column <= sel.column + sel.columns - 1,
      );
      if (!inSel) continue;
      requireDocument(
        !cell.spillOwner && !(cell.content.kind === 'python' && cell.content.block?.spillActive),
        'Cannot clear part of a spill range.',
      );
    }
  }

  const tableCmd = deleteTableSelection({ ...basis, document: doc }, selections);
  if (tableCmd) {
    doc = tableCmd.document;
  } else {
    const sheets = doc.sheets.map(sheet => {
      const sheetSelections = selections.filter(s => s.sheetId === sheet.id);
      if (!sheetSelections.length) return sheet;
      const cells = sheet.cells.map(cell => {
        const inSel = sheetSelections.some(sel =>
          cell.row >= sel.row && cell.row <= sel.row + sel.rows - 1 &&
          cell.column >= sel.column && cell.column <= sel.column + sel.columns - 1,
        );
        if (!inSel) return cell;
        const { cache: _c, ...restCell } = cell;
        return { ...restCell, content: { kind: 'blank' as const }, requiresExternalCalculation: false };
      }).filter(cell => cell.content.kind !== 'blank' || cell.style !== undefined);
      return { ...sheet, cells };
    });
    doc = { ...doc, sheets };
  }

  return {
    kind: 'structure',
    label: pivotsToRemove.length === 1 ? 'Delete PivotTable' : 'Delete PivotTables',
    document: doc,
  };
}

export function renamePivotTable(
  basis: WorkbookSnapshot,
  pivotId: DocumentId,
  name: string,
  makeId: IdFactory,
): StructureCommand {
  const current = (basis.document.pivots ?? []).find(item => item.id === pivotId);
  requireDocument(current, 'PivotTable no longer exists.');
  const trimmed = name.trim();
  requireDocument(trimmed.length > 0 && trimmed.length <= 255, 'PivotTable requires a valid name.');
  const nameKey = ordinalNameKey(trimmed);
  requireDocument(
    !(basis.document.pivots ?? []).some(item => item.id !== pivotId && ordinalNameKey(item.name) === nameKey),
    'A PivotTable with this name already exists.',
  );
  return {
    kind: 'structure',
    label: 'Rename PivotTable',
    document: materializePivotDocument(basis.document, { ...current, name: trimmed }, makeId),
  };
}

function clearFieldHiddenItems<T extends PivotFilterField | PivotAxisField>(field: T): T {
  const { hiddenItems: _hidden, ...rest } = field;
  return rest as T;
}

export function clearPivotTable(
  basis: WorkbookSnapshot,
  pivotId: DocumentId,
  mode: 'all' | 'filters',
  makeId: IdFactory,
): StructureCommand {
  const current = (basis.document.pivots ?? []).find(item => item.id === pivotId);
  requireDocument(current, 'PivotTable no longer exists.');
  const { collapsedKeys: _prevCollapsed, ...basePivot } = current;
  const updated: PivotTableDocument = mode === 'all'
    ? {
      ...basePivot,
      filters: [],
      columns: [],
      rows: [],
      values: [],
    }
    : {
      ...current,
      filters: current.filters.map(clearFieldHiddenItems),
      columns: current.columns.map(clearFieldHiddenItems),
      rows: current.rows.map(clearFieldHiddenItems),
    };
  return {
    kind: 'structure',
    label: mode === 'all' ? 'Clear PivotTable' : 'Clear PivotTable Filters',
    document: materializePivotDocument(basis.document, updated, makeId),
  };
}

export function expandCollapsePivotField(
  basis: WorkbookSnapshot,
  pivotId: DocumentId,
  fieldName: string,
  expand: boolean,
  makeId: IdFactory,
): StructureCommand {
  const current = (basis.document.pivots ?? []).find(item => item.id === pivotId);
  requireDocument(current, 'PivotTable no longer exists.');
  const keys = collectFieldCollapsibleKeys(basis.document, current, fieldName);
  const collapsed = new Set(current.collapsedKeys ?? []);
  for (const key of keys) {
    if (expand) collapsed.delete(key);
    else collapsed.add(key);
  }
  const { collapsedKeys: _prevCollapsed, ...basePivot } = current;
  const updated: PivotTableDocument = {
    ...basePivot,
    ...(collapsed.size > 0 && { collapsedKeys: [...collapsed] }),
  };
  return {
    kind: 'structure',
    label: expand ? 'Expand PivotTable Field' : 'Collapse PivotTable Field',
    document: materializePivotDocument(basis.document, updated, makeId),
  };
}

export function changePivotSource(
  basis: WorkbookSnapshot,
  pivotId: DocumentId,
  source: PivotSource,
  makeId: IdFactory,
): StructureCommand {
  const current = (basis.document.pivots ?? []).find(item => item.id === pivotId);
  requireDocument(current, 'PivotTable no longer exists.');
  return {
    kind: 'structure',
    label: 'Change PivotTable Data Source',
    document: materializePivotDocument(basis.document, { ...current, source }, makeId),
  };
}

export function movePivotTableWithSheet(
  basis: WorkbookSnapshot,
  pivotId: DocumentId,
  input: { readonly target?: CellReference; readonly newSheet?: boolean; readonly insertBeforeSheetId?: DocumentId },
  makeId: IdFactory,
): { readonly command: StructureCommand; readonly sheetId: DocumentId } {
  const current = (basis.document.pivots ?? []).find(item => item.id === pivotId);
  requireDocument(current, 'PivotTable no longer exists.');
  const reverted = revertPivotCellsOnSheet(basis.document, pivotId);
  if (input.newSheet) {
    requireDocument(reverted.sheets.length < documentLimits.sheets, `A workbook may contain at most ${documentLimits.sheets} worksheets.`);
    const newSheetId = makeId();
    const sheetName = nextAutoSheetName(reverted);
    const newSheet = emptySheet(newSheetId, sheetName);
    const sheets = [...reverted.sheets];
    const insertAt = input.insertBeforeSheetId ? sheets.findIndex(s => s.id === input.insertBeforeSheetId) : -1;
    if (insertAt >= 0) sheets.splice(insertAt, 0, newSheet);
    else sheets.push(newSheet);
    const target: CellReference = { sheetId: newSheetId, row: 3, column: 1 };
    return {
      command: {
        kind: 'structure',
        label: 'Move PivotTable',
        document: materializePivotDocument({ ...reverted, sheets }, { ...current, target }, makeId),
      },
      sheetId: newSheetId,
    };
  }
  requireDocument(input.target, 'PivotTable requires a destination cell.');
  return {
    command: {
      kind: 'structure',
      label: 'Move PivotTable',
      document: materializePivotDocument(reverted, { ...current, target: input.target }, makeId),
    },
    sheetId: input.target.sheetId,
  };
}

export function findPivotAt(document: WorkbookDocument, target: CellReference): PivotTableDocument | undefined {
  for (const pivot of document.pivots ?? []) {
    if (pivot.target.sheetId !== target.sheetId) continue;
    const extent = pivot.extent;
    if (extent) {
      if (target.row >= extent.row && target.row < extent.row + extent.rows &&
          target.column >= extent.column && target.column < extent.column + extent.columns) {
        return pivot;
      }
    } else if (target.row === pivot.target.row && target.column === pivot.target.column) {
      return pivot;
    }
  }
  return undefined;
}

export function parsePivotRangeReference(
  document: WorkbookDocument,
  activeSheetId: DocumentId,
  text: string,
): PivotSource | undefined {
  const trimmed = text.trim();
  if (!trimmed) return undefined;
  const key = ordinalNameKey(trimmed);
  const table = document.names.find(entry => entry.isTable && !entry.area.invalid && ordinalNameKey(entry.name) === key);
  if (table) return { kind: 'table', tableId: table.id };
  const area = parseNameExtent(document, activeSheetId, trimmed);
  return { kind: 'range', area };
}

export function parsePivotTargetCell(
  document: WorkbookDocument,
  activeSheetId: DocumentId,
  text: string,
): CellReference | undefined {
  const trimmed = text.trim();
  if (!trimmed) return undefined;
  const area = parseNameExtent(document, activeSheetId, trimmed);
  return { sheetId: area.sheetId, row: area.row, column: area.column };
}

export function createPivotTableWithSheet(
  basis: WorkbookSnapshot,
  input: Omit<CreatePivotInput, 'target'> & { readonly target?: CellReference; readonly newSheet?: boolean; readonly insertBeforeSheetId?: DocumentId },
  makeId: IdFactory,
): { readonly command: StructureCommand; readonly sheetId: DocumentId; readonly pivotId: DocumentId } {
  if (input.newSheet) {
    requireDocument(basis.document.sheets.length < documentLimits.sheets, `A workbook may contain at most ${documentLimits.sheets} worksheets.`);
    const used = new Set(basis.document.sheets.map(s => ordinalNameKey(s.name)));
    let n = basis.document.sheets.length + 1;
    while (used.has(ordinalNameKey(`Sheet${n}`))) n++;
    const newSheetId = makeId();
    const sheetName = `Sheet${n}`;
    const newSheet = emptySheet(newSheetId, sheetName);
    const sheets = [...basis.document.sheets];
    const insertAt = input.insertBeforeSheetId ? sheets.findIndex(s => s.id === input.insertBeforeSheetId) : -1;
    if (insertAt >= 0) sheets.splice(insertAt, 0, newSheet);
    else sheets.push(newSheet);
    const nextDoc = { ...basis.document, sheets };
    const command = createPivotTable(
      { ...basis, document: nextDoc },
      { ...input, target: { sheetId: newSheetId, row: 3, column: 1 } },
      makeId,
    );
    const createdPivot = command.document.pivots![command.document.pivots!.length - 1]!;
    return { command, sheetId: newSheetId, pivotId: createdPivot.id };
  }
  requireDocument(input.target, 'PivotTable requires a destination cell.');
  const command = createPivotTable(basis, { ...input, target: input.target }, makeId);
  const createdPivot = command.document.pivots![command.document.pivots!.length - 1]!;
  return { command, sheetId: input.target.sheetId, pivotId: createdPivot.id };
}

function nextAutoSheetName(document: WorkbookDocument): string {
  const used = new Set(document.sheets.map(s => ordinalNameKey(s.name)));
  let n = document.sheets.length + 1;
  while (used.has(ordinalNameKey(`Sheet${n}`))) n++;
  return `Sheet${n}`;
}

export function drillDownPivotCell(
  basis: WorkbookSnapshot,
  pivotId: DocumentId,
  target: CellReference,
  makeId: IdFactory,
): { readonly command: StructureCommand; readonly sheetId: DocumentId } | undefined {
  const pivot = (basis.document.pivots ?? []).find(item => item.id === pivotId);
  if (!pivot) return undefined;
  const drill = extractPivotCellDrillDown(basis.document, pivot, target);
  if (!drill || drill.fields.length === 0) return undefined;

  requireDocument(basis.document.sheets.length < documentLimits.sheets, `A workbook may contain at most ${documentLimits.sheets} worksheets.`);
  const newSheetId = makeId();
  const sheetName = nextAutoSheetName(basis.document);
  const cells: CellDocument[] = [];

  for (let c = 0; c < drill.fields.length; c++) {
    cells.push({
      id: makeId(),
      row: 1,
      column: c + 1,
      content: { kind: 'text', value: drill.fields[c]! },
    });
  }

  for (let r = 0; r < drill.records.length; r++) {
    const rec = drill.records[r]!;
    for (let c = 0; c < drill.fields.length; c++) {
      const scalar = rec[drill.fields[c]!] ?? null;
      if (scalar === null || scalar === '') continue;
      if (typeof scalar === 'number') {
        cells.push({ id: makeId(), row: r + 2, column: c + 1, content: { kind: 'number', value: scalar } });
      } else if (typeof scalar === 'boolean') {
        cells.push({ id: makeId(), row: r + 2, column: c + 1, content: { kind: 'boolean', value: scalar } });
      } else if (typeof scalar === 'object') {
        cells.push({ id: makeId(), row: r + 2, column: c + 1, content: { kind: 'blank' }, cache: scalar });
      } else {
        cells.push({ id: makeId(), row: r + 2, column: c + 1, content: { kind: 'text', value: scalar } });
      }
    }
  }

  const newSheet = { ...emptySheet(newSheetId, sheetName), cells };
  const sheets = [...basis.document.sheets];
  const insertAt = sheets.findIndex(s => s.id === pivot.target.sheetId);
  if (insertAt >= 0) sheets.splice(insertAt, 0, newSheet);
  else sheets.push(newSheet);

  const populatedBasis: WorkbookSnapshot = { ...basis, document: { ...basis.document, sheets } };
  const tableExtent = {
    sheetId: newSheetId,
    row: 1,
    column: 1,
    rows: Math.max(2, drill.records.length + 1),
    columns: Math.max(1, drill.fields.length),
  };
  const tableCmd = createTable(populatedBasis, {
    id: makeId(),
    name: nextTableName(populatedBasis.document),
    comment: '',
    area: tableExtent,
    isTable: true,
  }, true, makeId);

  return {
    command: {
      kind: 'structure',
      label: 'Show PivotTable Details',
      document: tableCmd.document,
    },
    sheetId: newSheetId,
  };
}
