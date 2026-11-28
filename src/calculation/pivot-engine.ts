import { formatNumber } from '../core/formats/cell-format';
import {
  defaultPivotValueLabel,
  type PivotAggregation,
  type PivotAxisField,
  type PivotShowValuesAs,
  type PivotSource,
  type PivotSourceArea,
  type PivotTableDocument,
  type PivotValueField,
} from '../core/pivot';
import type {
  CellCache,
  CellContent,
  CellDocument,
  CellReference,
  CellStyle,
  DocumentId,
  WorkbookDocument,
} from '../core/workbook';

export type PivotScalar = string | number | boolean | null | Extract<CellCache, { kind: 'error' }>;
export type PivotRecord = Readonly<Record<string, PivotScalar>>;

export interface PivotToggleOverlay {
  readonly row: number;
  readonly column: number;
  readonly key: string;
  readonly collapsed: boolean;
  readonly label: string;
  readonly level: number;
}

export interface PivotFilterOverlay {
  readonly row: number;
  readonly column: number;
  readonly kind: 'reportFilter' | 'rowLabel' | 'colLabel';
  readonly fieldNames: readonly string[];
  readonly active: boolean;
  readonly sort?: 'asc' | 'desc';
  readonly label: string;
}

export interface PivotInteractiveOverlay {
  readonly pivotId: DocumentId;
  readonly toggles: readonly PivotToggleOverlay[];
  readonly filters: readonly PivotFilterOverlay[];
}

export interface ActivePivotFieldInfo {
  readonly displayName: string;
  readonly fieldName: string;
  readonly area: 'filters' | 'columns' | 'rows' | 'values';
  readonly axisField?: PivotAxisField;
  readonly valueField?: PivotValueField;
  readonly canExpandCollapse: boolean;
}

export interface PivotCellOutput {
  readonly rowOffset: number;
  readonly colOffset: number;
  readonly content: CellContent;
  readonly cache?: CellCache;
  readonly style?: CellStyle;
  readonly role: 'filter' | 'header' | 'rowLabel' | 'colLabel' | 'data' | 'subtotal' | 'grandTotal' | 'blank';
  readonly indent?: number;
  readonly collapsibleKey?: string;
  readonly collapsed?: boolean;
  readonly fieldName?: string;
  readonly valueFieldId?: string;
}

export interface PivotEvaluationResult {
  readonly rows: number;
  readonly columns: number;
  readonly cells: readonly PivotCellOutput[];
  readonly fields: readonly string[];
  readonly recordCount: number;
}

interface ItemEntry {
  readonly label: string;
  readonly displayLabel: string;
  readonly raw: PivotScalar;
}

interface AxisNode {
  readonly level: number;
  readonly fieldName: string;
  readonly item: ItemEntry;
  readonly key: string;
  readonly parentKey: string;
  readonly records: readonly PivotRecord[];
  readonly children: readonly AxisNode[];
  readonly subtotal: 'automatic' | 'none';
  readonly valueIndex?: number;
}

interface AxisLine {
  readonly kind: 'item' | 'subtotal' | 'grandTotal';
  readonly node?: AxisNode;
  readonly path: readonly AxisNode[];
  readonly valueIndex: number;
  readonly label: string;
  readonly level: number;
  readonly hasChildren: boolean;
  readonly collapsed: boolean;
  readonly collapsibleKey?: string;
  readonly records: readonly PivotRecord[];
}

function readCellScalar(cell: CellDocument | undefined): PivotScalar {
  if (!cell) return null;
  switch (cell.content.kind) {
    case 'blank':
      return cell.spillKind === 'formula' && cell.cache ? cell.cache.kind === 'error' ? cell.cache : cell.cache.value : null;
    case 'number':
    case 'text':
    case 'boolean':
      return cell.content.value;
    case 'formula':
      return cell.cache ? (cell.cache.kind === 'error' ? cell.cache : cell.cache.value) : null;
    case 'python': {
      const val = cell.content.block?.hasResult ? cell.content.block.lastValue : undefined;
      return val ? (val.kind === 'error' ? val : val.value) : null;
    }
  }
}

export function formatPivotItemLabel(value: PivotScalar): string {
  if (value === null) return '(blank)';
  if (typeof value === 'object') return value.value;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'string') return value === '' ? '(blank)' : value;
  return String(value);
}

function compareScalars(left: PivotScalar, right: PivotScalar): number {
  if (left === null && right === null) return 0;
  if (left === null) return 1;
  if (right === null) return -1;
  if (typeof left === 'number' && typeof right === 'number') return left - right;
  if (typeof left === 'boolean' && typeof right === 'boolean') return Number(left) - Number(right);
  return formatPivotItemLabel(left).localeCompare(formatPivotItemLabel(right), undefined, { numeric: true, sensitivity: 'base' });
}

export function resolvePivotSourceExtent(document: WorkbookDocument, source: PivotSource): PivotSourceArea | undefined {
  if (source.kind === 'range') {
    return source.area.invalid ? undefined : source.area;
  }
  const table = document.names.find(entry => entry.id === source.tableId && entry.isTable);
  if (!table || table.area.invalid) return undefined;
  const totalOffset = table.design?.totalRow ? 1 : 0;
  return {
    ...table.area,
    rows: Math.max(1, table.area.rows - totalOffset),
  };
}

function cellKeyCoord(coord: { readonly row: number; readonly column: number }): number {
  return (coord.row - 1) * 16_384 + (coord.column - 1);
}

export function extractPivotRecords(
  document: WorkbookDocument,
  source: PivotSource,
): {
  readonly fields: readonly string[];
  readonly records: readonly PivotRecord[];
  readonly fieldFormats: Readonly<Record<string, string>>;
} {
  const extent = resolvePivotSourceExtent(document, source);
  if (!extent) return { fields: [], records: [], fieldFormats: {} };
  const sheet = document.sheets.find(s => s.id === extent.sheetId);
  if (!sheet) return { fields: [], records: [], fieldFormats: {} };

  const byKey = new Map<number, CellDocument>(sheet.cells.map(cell => [cellKeyCoord(cell), cell]));
  const rowFormats = new Map<number, string>();
  for (const dim of sheet.rows) {
    if (dim.style?.numberFormat) rowFormats.set(dim.index, dim.style.numberFormat);
  }
  const colFormats = new Map<number, string>();
  for (const dim of sheet.columns) {
    if (dim.style?.numberFormat) colFormats.set(dim.index, dim.style.numberFormat);
  }

  const rawFields: string[] = [];
  const used = new Map<string, number>();

  for (let c = 0; c < extent.columns; c++) {
    const raw = readCellScalar(byKey.get(cellKeyCoord({ row: extent.row, column: extent.column + c })));
    const base = raw === null || formatPivotItemLabel(raw) === '(blank)' ? `Column${c + 1}` : formatPivotItemLabel(raw).trim();
    const upper = base.toUpperCase();
    const seen = used.get(upper) ?? 0;
    const unique = seen === 0 ? base : `${base}${seen + 1}`;
    used.set(upper, seen + 1);
    rawFields.push(unique);
  }

  const fieldFormats: Record<string, string> = {};
  const records: PivotRecord[] = [];
  for (let r = 1; r < extent.rows; r++) {
    const rowIndex = extent.row + r;
    const record: Record<string, PivotScalar> = {};
    let hasValue = false;
    for (let c = 0; c < extent.columns; c++) {
      const colIndex = extent.column + c;
      const cell = byKey.get(cellKeyCoord({ row: rowIndex, column: colIndex }));
      const val = readCellScalar(cell);
      if (val !== null && val !== '') hasValue = true;
      const fieldName = rawFields[c]!;
      record[fieldName] = val;
      if (fieldFormats[fieldName] === undefined && val !== null && val !== '') {
        const fmt = cell?.style?.numberFormat ?? rowFormats.get(rowIndex) ?? colFormats.get(colIndex);
        if (fmt && fmt !== 'General') fieldFormats[fieldName] = fmt;
      }
    }
    if (hasValue) records.push(record);
  }
  for (let c = 0; c < extent.columns; c++) {
    const fieldName = rawFields[c]!;
    if (fieldFormats[fieldName] === undefined) {
      const colFmt = colFormats.get(extent.column + c);
      if (colFmt && colFmt !== 'General') fieldFormats[fieldName] = colFmt;
    }
  }
  return { fields: rawFields, records, fieldFormats };
}

type RawValue = { readonly kind: 'blank' } | { readonly kind: 'error'; readonly value: Extract<CellCache, { kind: 'error' }>['value'] } | { readonly kind: 'number'; readonly value: number };

function aggregateNumbers(values: readonly PivotScalar[], aggregation: PivotAggregation): RawValue {
  if (values.length === 0) return { kind: 'blank' };
  const nonNull = values.filter(v => v !== null && v !== '');
  if (aggregation === 'count') {
    return nonNull.length === 0 ? { kind: 'blank' } : { kind: 'number', value: nonNull.length };
  }
  for (const val of nonNull) {
    if (val && typeof val === 'object') return { kind: 'error', value: val.value };
  }
  const nums = nonNull.filter((v): v is number => typeof v === 'number' && Number.isFinite(v));
  if (aggregation === 'countNumbers') {
    return { kind: 'number', value: nums.length };
  }
  if (nums.length === 0) {
    if (aggregation === 'average' || aggregation === 'stdDev' || aggregation === 'stdDevp' || aggregation === 'var' || aggregation === 'varp') {
      return { kind: 'error', value: '#DIV/0!' };
    }
    return { kind: 'blank' };
  }
  switch (aggregation) {
    case 'sum':
      return { kind: 'number', value: nums.reduce((acc, n) => acc + n, 0) };
    case 'average':
      return { kind: 'number', value: nums.reduce((acc, n) => acc + n, 0) / nums.length };
    case 'max':
      return { kind: 'number', value: Math.max(...nums) };
    case 'min':
      return { kind: 'number', value: Math.min(...nums) };
    case 'product':
      return { kind: 'number', value: nums.reduce((acc, n) => acc * n, 1) };
    case 'varp':
    case 'stdDevp': {
      const mean = nums.reduce((acc, n) => acc + n, 0) / nums.length;
      const variance = nums.reduce((acc, n) => acc + (n - mean) ** 2, 0) / nums.length;
      return { kind: 'number', value: aggregation === 'stdDevp' ? Math.sqrt(variance) : variance };
    }
    case 'var':
    case 'stdDev': {
      if (nums.length < 2) return { kind: 'error', value: '#DIV/0!' };
      const mean = nums.reduce((acc, n) => acc + n, 0) / nums.length;
      const variance = nums.reduce((acc, n) => acc + (n - mean) ** 2, 0) / (nums.length - 1);
      return { kind: 'number', value: aggregation === 'stdDev' ? Math.sqrt(variance) : variance };
    }
  }
}

function summarizeBucket(records: readonly PivotRecord[], field: PivotValueField): RawValue {
  if (records.length === 0) return { kind: 'blank' };
  return aggregateNumbers(records.map(r => r[field.fieldName] ?? null), field.aggregation);
}

function buildAxisTree(
  records: readonly PivotRecord[],
  fields: readonly PivotAxisField[],
  values: readonly PivotValueField[],
  fieldFormats: Readonly<Record<string, string>>,
  level = 0,
  parentKey = '',
): readonly AxisNode[] {
  if (level >= fields.length) return [];
  const field = fields[level]!;
  const explicitAxisFmt = field.numberFormat;
  const axisFmt = explicitAxisFmt !== undefined
    ? (/^general$/iu.test(explicitAxisFmt) ? undefined : explicitAxisFmt)
    : fieldFormats[field.fieldName];
  const hidden = new Set(field.hiddenItems ?? []);
  const groups = new Map<string, { item: ItemEntry; records: PivotRecord[] }>();
  for (const record of records) {
    const raw = record[field.fieldName] ?? null;
    const label = formatPivotItemLabel(raw);
    if (hidden.has(label)) continue;
    let bucket = groups.get(label);
    if (!bucket) {
      const displayLabel = typeof raw === 'number' && axisFmt && !/^general$/iu.test(axisFmt)
        ? formatNumber(raw, axisFmt).text
        : label;
      bucket = { item: { label, displayLabel, raw }, records: [] };
      groups.set(label, bucket);
    }
    bucket.records.push(record);
  }

  const sortField = field.sort?.byValueFieldId
    ? values.find(v => v.id === field.sort?.byValueFieldId)
    : undefined;
  const entries = [...groups.values()];
  entries.sort((left, right) => {
    let cmp: number;
    if (sortField) {
      const lv = summarizeBucket(left.records, sortField);
      const rv = summarizeBucket(right.records, sortField);
      const ln = lv.kind === 'number' ? lv.value : Number.NEGATIVE_INFINITY;
      const rn = rv.kind === 'number' ? rv.value : Number.NEGATIVE_INFINITY;
      cmp = ln - rn || compareScalars(left.item.raw, right.item.raw);
    } else {
      cmp = compareScalars(left.item.raw, right.item.raw);
    }
    return field.sort?.direction === 'desc' ? -cmp : cmp;
  });

  return entries.map(({ item, records: groupRecords }) => {
    const key = parentKey ? `${parentKey}\u0000${item.label}` : item.label;
    return {
      level,
      fieldName: field.fieldName,
      item,
      key,
      parentKey,
      records: groupRecords,
      children: buildAxisTree(groupRecords, fields, values, fieldFormats, level + 1, key),
      subtotal: field.subtotal ?? 'automatic',
    };
  });
}

function flattenAxis(
  nodes: readonly AxisNode[],
  fields: readonly PivotAxisField[],
  subtotals: PivotTableDocument['subtotals'],
  collapsedKeys: ReadonlySet<string>,
  repeatLabels: boolean,
  reportLayout: PivotTableDocument['reportLayout'] = 'compact',
  path: readonly AxisNode[] = [],
): readonly AxisLine[] {
  const lines: AxisLine[] = [];
  for (const node of nodes) {
    const currentPath = [...path, node];
    const hasChildren = node.children.length > 0 && node.level < fields.length - 1;
    const collapsed = collapsedKeys.has(node.key);
    const showSubtotals = hasChildren && !collapsed && subtotals !== 'off' && node.subtotal !== 'none';
    const childLines = hasChildren && !collapsed
      ? flattenAxis(node.children, fields, subtotals, collapsedKeys, repeatLabels, reportLayout, currentPath)
      : [];

    if (!hasChildren || collapsed) {
      lines.push({
        kind: 'item',
        node,
        path: currentPath,
        valueIndex: 0,
        label: node.item.displayLabel,
        level: node.level,
        hasChildren,
        collapsed,
        ...(hasChildren && { collapsibleKey: node.key }),
        records: node.records,
      });
      continue;
    }

    if (reportLayout === 'tabular') {
      lines.push(...childLines);
      if (showSubtotals) {
        lines.push({
          kind: 'subtotal',
          node,
          path: currentPath,
          valueIndex: 0,
          label: `${node.item.displayLabel} Total`,
          level: node.level,
          hasChildren: false,
          collapsed: false,
          records: node.records,
        });
      }
      continue;
    }

    if (showSubtotals && subtotals === 'top') {
      lines.push({
        kind: 'subtotal',
        node,
        path: currentPath,
        valueIndex: 0,
        label: node.item.displayLabel,
        level: node.level,
        hasChildren: true,
        collapsed: false,
        collapsibleKey: node.key,
        records: node.records,
      });
    } else if (!showSubtotals) {
      // Compact/outline still renders a parent header row when subtotals are off on top or bottom.
      lines.push({
        kind: 'item',
        node,
        path: currentPath,
        valueIndex: 0,
        label: node.item.displayLabel,
        level: node.level,
        hasChildren: true,
        collapsed: false,
        collapsibleKey: node.key,
        records: [],
      });
    }
    lines.push(...childLines);
    if (showSubtotals && subtotals === 'bottom') {
      lines.push({
        kind: 'subtotal',
        node,
        path: currentPath,
        valueIndex: 0,
        label: `${node.item.displayLabel} Total`,
        level: node.level,
        hasChildren: false,
        collapsed: false,
        records: node.records,
      });
    }
  }
  return lines;
}

function expandAxisWithValues(
  lines: readonly AxisLine[],
  valueCount: number,
): readonly AxisLine[] {
  if (valueCount <= 1) return lines;
  const expanded: AxisLine[] = [];
  for (const line of lines) {
    for (let v = 0; v < valueCount; v++) {
      expanded.push({ ...line, valueIndex: v });
    }
  }
  return expanded;
}

function intersectRecords(rowRecords: readonly PivotRecord[], colRecords: readonly PivotRecord[]): readonly PivotRecord[] {
  if (rowRecords.length === 0 || colRecords.length === 0) return [];
  const colSet = new Set(colRecords);
  return rowRecords.filter(r => colSet.has(r));
}

function resolveShowValuesAs(
  mode: PivotShowValuesAs,
  rawMatrix: readonly (readonly RawValue[])[],
  rowLines: readonly AxisLine[],
  colLines: readonly AxisLine[],
  rIdx: number,
  cIdx: number,
  rowTotalColIndex: number | undefined,
  colTotalRowIndex: number | undefined,
  valueField: PivotValueField,
): RawValue {
  const val = rawMatrix[rIdx]![cIdx]!;
  if (val.kind !== 'number' || mode === 'none') return val;
  const gRow = colTotalRowIndex ?? rowLines.length - 1;
  const gCol = rowTotalColIndex ?? colLines.length - 1;
  const grand = rawMatrix[gRow]?.[gCol];
  const div = (num: number, den: RawValue | undefined): RawValue => {
    if (!den || den.kind === 'error') return den ?? { kind: 'error', value: '#DIV/0!' };
    if (den.kind === 'blank' || den.value === 0) return { kind: 'error', value: '#DIV/0!' };
    return { kind: 'number', value: num / den.value };
  };

  switch (mode) {
    case 'percentOfGrandTotal':
      return div(val.value, grand);
    case 'percentOfColumnTotal':
      return div(val.value, rawMatrix[gRow]?.[cIdx]);
    case 'percentOfRowTotal':
      return div(val.value, rawMatrix[rIdx]?.[gCol]);
    case 'percentOfParentRowTotal': {
      const path = rowLines[rIdx]!.path;
      if (path.length <= 1) return div(val.value, rawMatrix[gRow]?.[cIdx]);
      const parentKey = path[path.length - 2]!.key;
      const parentIdx = rowLines.findIndex(l => l.node?.key === parentKey && l.kind === 'subtotal');
      return parentIdx >= 0 ? div(val.value, rawMatrix[parentIdx]?.[cIdx]) : div(val.value, rawMatrix[gRow]?.[cIdx]);
    }
    case 'percentOfParentColumnTotal': {
      const path = colLines[cIdx]!.path;
      if (path.length <= 1) return div(val.value, rawMatrix[rIdx]?.[gCol]);
      const parentKey = path[path.length - 2]!.key;
      const parentIdx = colLines.findIndex(l => l.node?.key === parentKey && l.kind === 'subtotal');
      return parentIdx >= 0 ? div(val.value, rawMatrix[rIdx]?.[parentIdx]) : div(val.value, rawMatrix[rIdx]?.[gCol]);
    }
    case 'differenceFrom':
    case 'percentDifferenceFrom': {
      const baseItem = valueField.baseItem;
      let baseVal: RawValue | undefined;
      if (baseItem === '(previous)') {
        for (let p = rIdx - 1; p >= 0; p--) {
          if (rowLines[p]!.kind === 'item' && rowLines[p]!.valueIndex === rowLines[rIdx]!.valueIndex) {
            baseVal = rawMatrix[p]![cIdx];
            break;
          }
        }
      } else if (baseItem === '(next)') {
        for (let p = rIdx + 1; p < rowLines.length; p++) {
          if (rowLines[p]!.kind === 'item' && rowLines[p]!.valueIndex === rowLines[rIdx]!.valueIndex) {
            baseVal = rawMatrix[p]![cIdx];
            break;
          }
        }
      } else if (baseItem !== undefined) {
        const targetField = valueField.baseField;
        const rowMatch = rowLines.findIndex((l, idx) =>
          l.kind === 'item' && l.valueIndex === rowLines[rIdx]!.valueIndex &&
          (targetField ? l.node?.fieldName === targetField : true) && l.node?.item.label === baseItem && idx >= 0);
        if (rowMatch >= 0) baseVal = rawMatrix[rowMatch]![cIdx];
        else {
          const colMatch = colLines.findIndex(l =>
            l.kind === 'item' && l.valueIndex === colLines[cIdx]!.valueIndex &&
            (targetField ? l.node?.fieldName === targetField : true) && l.node?.item.label === baseItem);
          if (colMatch >= 0) baseVal = rawMatrix[rIdx]![colMatch];
        }
      }
      if (!baseVal || baseVal.kind === 'blank') return { kind: 'blank' };
      if (baseVal.kind === 'error') return baseVal;
      if (mode === 'differenceFrom') return { kind: 'number', value: val.value - baseVal.value };
      if (baseVal.value === 0) return { kind: 'error', value: '#DIV/0!' };
      return { kind: 'number', value: (val.value - baseVal.value) / Math.abs(baseVal.value) };
    }
    case 'runningTotalIn': {
      const targetRow = rowLines[rIdx]!;
      const parentKey = targetRow.path.length > 1 ? targetRow.path[targetRow.path.length - 2]!.key : '';
      let sum = 0;
      for (let i = 0; i <= rIdx; i++) {
        const cur = rowLines[i]!;
        if (cur.kind !== 'item' || cur.valueIndex !== targetRow.valueIndex) continue;
        const curParent = cur.path.length > 1 ? cur.path[cur.path.length - 2]!.key : '';
        if (curParent !== parentKey) continue;
        const cell = rawMatrix[i]![cIdx]!;
        if (cell.kind === 'number') sum += cell.value;
      }
      return { kind: 'number', value: sum };
    }
    case 'rankAscending':
    case 'rankDescending': {
      const targetRow = rowLines[rIdx]!;
      if (targetRow.kind !== 'item') return { kind: 'blank' };
      const peers: number[] = [];
      for (let i = 0; i < rowLines.length; i++) {
        const cur = rowLines[i]!;
        if (cur.kind !== 'item' || cur.valueIndex !== targetRow.valueIndex) continue;
        const cell = rawMatrix[i]![cIdx]!;
        if (cell.kind === 'number') peers.push(cell.value);
      }
      const rank = 1 + peers.filter(p => (mode === 'rankAscending' ? p < val.value : p > val.value)).length;
      return { kind: 'number', value: rank };
    }
    case 'index': {
      const rowTotal = rawMatrix[rIdx]?.[gCol];
      const colTotal = rawMatrix[gRow]?.[cIdx];
      if (!rowTotal || !colTotal || !grand || rowTotal.kind !== 'number' || colTotal.kind !== 'number' || grand.kind !== 'number') {
        return { kind: 'error', value: '#DIV/0!' };
      }
      const den = rowTotal.value * colTotal.value;
      if (den === 0) return { kind: 'error', value: '#DIV/0!' };
      return { kind: 'number', value: (val.value * grand.value) / den };
    }
  }
}

function contentAndCache(raw: RawValue): { content: CellContent; cache?: CellCache } {
  if (raw.kind === 'blank') return { content: { kind: 'blank' } };
  if (raw.kind === 'error') return { content: { kind: 'blank' }, cache: { kind: 'error', value: raw.value } };
  return { content: { kind: 'number', value: raw.value } };
}

interface PivotPalette {
  readonly headerFill: string;
  readonly totalFill: string;
  readonly bandFill: string;
  readonly accentBorder: string;
  readonly filterFill: string;
}

const pivotPalettes: Readonly<Record<PivotTableDocument['style'], PivotPalette>> = {
  Blue: { headerFill: '1E4471', totalFill: '163150', bandFill: '1A2634', accentBorder: '4F81BD', filterFill: '1C2836' },
  Green: { headerFill: '214E36', totalFill: '173826', bandFill: '192920', accentBorder: '548235', filterFill: '1B2A21' },
  Orange: { headerFill: '693714', totalFill: '4B270E', bandFill: '2D2018', accentBorder: 'C65911', filterFill: '2E2118' },
  Purple: { headerFill: '462D66', totalFill: '32204A', bandFill: '251E30', accentBorder: '8064A2', filterFill: '261F31' },
};

interface InternalPivotResult {
  readonly evaluation: PivotEvaluationResult;
  readonly rowHeaderRowOffset: number;
  readonly headerBannerRow: number;
  readonly labelColCount: number;
  readonly dataStartCol: number;
  readonly dataBuckets: ReadonlyMap<string, readonly PivotRecord[]>;
  readonly rowLines: readonly AxisLine[];
  readonly colLines: readonly AxisLine[];
  readonly displayMatrix: readonly (readonly RawValue[])[];
  readonly valuesInRows: boolean;
  readonly hasMultipleValues: boolean;
}

export type PivotEvaluationCache = Map<string, InternalPivotResult>;

function evaluatePivotInternal(document: WorkbookDocument, pivot: PivotTableDocument): InternalPivotResult {
  const { fields, records: allRecords, fieldFormats } = extractPivotRecords(document, pivot.source);
  const filteredRecords = allRecords.filter(record => pivot.filters.every(filter => {
    if (!filter.hiddenItems?.length) return true;
    return !filter.hiddenItems.includes(formatPivotItemLabel(record[filter.fieldName] ?? null));
  }));

  const collapsed = new Set(pivot.collapsedKeys ?? []);
  const valuesInRows = Boolean(pivot.valuesInRows);
  const hasMultipleValues = pivot.values.length > 1;
  const repeatLabels = Boolean(pivot.repeatLabels);
  const palette = pivotPalettes[pivot.style] ?? pivotPalettes.Blue;
  const styleOptions = pivot.styleOptions;
  const showButtons = pivot.showButtons !== false;
  const showFieldHeaders = pivot.showFieldHeaders !== false;

  const headerStyle: CellStyle = {
    bold: true,
    ...(styleOptions.columnHeaders && { fillRgb: palette.headerFill }),
    borderDetails: { bottom: { style: 'thin', rgb: palette.accentBorder } },
  };
  const totalStyleBase: CellStyle = {
    bold: true,
    fillRgb: palette.totalFill,
    borderDetails: { top: { style: 'thin', rgb: palette.accentBorder } },
  };

  const cells: PivotCellOutput[] = [];
  const dataBuckets = new Map<string, readonly PivotRecord[]>();
  let rowCursor = 0;

  if (pivot.filters.length > 0) {
    for (const filter of pivot.filters) {
      const hidden = new Set(filter.hiddenItems ?? []);
      const distinct = [...new Set(allRecords.map(r => formatPivotItemLabel(r[filter.fieldName] ?? null)))];
      const visible = distinct.filter(item => !hidden.has(item));
      const summary = visible.length === distinct.length ? '(All)' : visible.length === 1 ? visible[0]! : '(Multiple Items)';
      cells.push({
        rowOffset: rowCursor, colOffset: 0,
        content: { kind: 'text', value: filter.fieldName },
        style: { bold: true, fillRgb: palette.filterFill }, role: 'filter', fieldName: filter.fieldName,
      });
      cells.push({
        rowOffset: rowCursor, colOffset: 1,
        content: { kind: 'text', value: summary },
        style: { fillRgb: palette.filterFill }, role: 'filter', fieldName: filter.fieldName,
      });
      rowCursor++;
    }
    rowCursor++; // Blank row between Report Filters and PivotTable body.
  }

  if (pivot.rows.length === 0 && pivot.columns.length === 0 && pivot.values.length === 0) {
    cells.push({
      rowOffset: rowCursor, colOffset: 0,
      content: { kind: 'text', value: pivot.name },
      style: headerStyle, role: 'header',
    });
    return {
      evaluation: { rows: rowCursor + 1, columns: 1, cells, fields, recordCount: filteredRecords.length },
      rowHeaderRowOffset: rowCursor,
      headerBannerRow: rowCursor,
      labelColCount: 1,
      dataStartCol: 1,
      dataBuckets,
      rowLines: [],
      colLines: [],
      displayMatrix: [],
      valuesInRows,
      hasMultipleValues,
    };
  }

  const visibleRecords = filteredRecords.filter(record =>
    pivot.rows.every(rf => !rf.hiddenItems?.includes(formatPivotItemLabel(record[rf.fieldName] ?? null))) &&
    pivot.columns.every(cf => !cf.hiddenItems?.includes(formatPivotItemLabel(record[cf.fieldName] ?? null))),
  );

  const rowTree = buildAxisTree(visibleRecords, pivot.rows, pivot.values, fieldFormats);
  const colTree = buildAxisTree(visibleRecords, pivot.columns, pivot.values, fieldFormats);

  const baseRowLines = pivot.rows.length > 0
    ? flattenAxis(rowTree, pivot.rows, pivot.subtotals, collapsed, repeatLabels, pivot.reportLayout)
    : [{ kind: 'item' as const, path: [] as AxisNode[], valueIndex: 0, label: '', level: 0, hasChildren: false, collapsed: false, records: visibleRecords }];
  const baseColLines = pivot.columns.length > 0
    ? flattenAxis(colTree, pivot.columns, pivot.subtotals, collapsed, true, 'compact')
    : [{ kind: 'item' as const, path: [] as AxisNode[], valueIndex: 0, label: '', level: 0, hasChildren: false, collapsed: false, records: visibleRecords }];

  const rowValueCount = valuesInRows && hasMultipleValues ? pivot.values.length : 1;
  const colValueCount = !valuesInRows && hasMultipleValues ? pivot.values.length : 1;

  const rowLines: AxisLine[] = [...expandAxisWithValues(baseRowLines, rowValueCount)];
  if (pivot.grandTotals.rows && (pivot.rows.length > 0 || (valuesInRows && hasMultipleValues))) {
    const gtRows = expandAxisWithValues([{
      kind: 'grandTotal', path: [], valueIndex: 0, label: 'Grand Total', level: 0,
      hasChildren: false, collapsed: false, records: visibleRecords,
    }], rowValueCount);
    rowLines.push(...gtRows);
  }

  const colLines: AxisLine[] = [...expandAxisWithValues(baseColLines, colValueCount)];
  if (pivot.grandTotals.columns && pivot.columns.length > 0) {
    const gtCols = expandAxisWithValues([{
      kind: 'grandTotal', path: [], valueIndex: 0, label: 'Grand Total', level: 0,
      hasChildren: false, collapsed: false, records: visibleRecords,
    }], colValueCount);
    colLines.push(...gtCols);
  }

  const rowTotalColIndex = colLines.findIndex(line => line.kind === 'grandTotal');
  const colTotalRowIndex = rowLines.findIndex(line => line.kind === 'grandTotal');

  // Precompute raw aggregation matrix (rowLines × colLines) and record buckets for drill-down.
  const bucketMatrix: (readonly PivotRecord[])[][] = rowLines.map(rLine => {
    const rowEmptyGroup = rLine.kind === 'item' && rLine.records.length === 0;
    return colLines.map(cLine => {
      if (rowEmptyGroup || (cLine.kind === 'item' && cLine.records.length === 0)) return [];
      const rowBucket = pivot.rows.length > 0 ? rLine.records : visibleRecords;
      const colBucket = pivot.columns.length > 0 ? cLine.records : visibleRecords;
      return pivot.rows.length > 0 && pivot.columns.length > 0
        ? intersectRecords(rowBucket, colBucket)
        : (pivot.rows.length > 0 ? rowBucket : colBucket);
    });
  });

  const rawMatrix: RawValue[][] = rowLines.map((rLine, rIdx) =>
    colLines.map((cLine, cIdx) => {
      if (pivot.values.length === 0) return { kind: 'blank' };
      const bucket = bucketMatrix[rIdx]![cIdx]!;
      if (bucket.length === 0) return { kind: 'blank' };
      const valueIndex = valuesInRows ? rLine.valueIndex : cLine.valueIndex;
      const valueField = pivot.values[valueIndex];
      if (!valueField) return { kind: 'blank' };
      return summarizeBucket(bucket, valueField);
    }),
  );

  const displayMatrix: RawValue[][] = rawMatrix.map((rowArr, rIdx) =>
    rowArr.map((_, cIdx) => {
      const valueIndex = valuesInRows ? rowLines[rIdx]!.valueIndex : colLines[cIdx]!.valueIndex;
      const vf = pivot.values[valueIndex];
      if (!vf || !vf.showValuesAs || vf.showValuesAs === 'none') return rawMatrix[rIdx]![cIdx]!;
      return resolveShowValuesAs(vf.showValuesAs, rawMatrix, rowLines, colLines, rIdx, cIdx,
        rowTotalColIndex >= 0 ? rowTotalColIndex : undefined,
        colTotalRowIndex >= 0 ? colTotalRowIndex : undefined, vf);
    }),
  );

  const labelColCount = pivot.rows.length === 0 && !(valuesInRows && hasMultipleValues)
    ? 0
    : pivot.reportLayout === 'compact'
      ? 1
      : pivot.rows.length + (valuesInRows && hasMultipleValues ? 1 : 0);
  const dataStartCol = labelColCount;

  const hasColumnHierarchy = pivot.columns.length > 0;
  const colHeaderRows = hasColumnHierarchy ? Math.max(1, pivot.columns.length) : 1;
  const singleValueBannerOccupied = pivot.values.length === 1 && hasColumnHierarchy && labelColCount > 0;
  if (singleValueBannerOccupied) {
    const vf = pivot.values[0]!;
    cells.push({
      rowOffset: rowCursor, colOffset: 0,
      content: { kind: 'text', value: vf.customName ?? defaultPivotValueLabel(vf.fieldName, vf.aggregation) },
      style: headerStyle, role: 'header', fieldName: vf.fieldName, valueFieldId: vf.id,
    });
  }
  const headerBannerRow = rowCursor;
  if (hasColumnHierarchy && showFieldHeaders) {
    const colField = pivot.columns[0]?.fieldName;
    cells.push({
      rowOffset: headerBannerRow, colOffset: dataStartCol,
      content: { kind: 'text', value: 'Column Labels' },
      style: headerStyle, role: 'header', ...(colField && { fieldName: colField }),
    });
  }
  const rowHeaderRowOffset = hasColumnHierarchy
    ? (showFieldHeaders ? rowCursor + colHeaderRows : rowCursor + colHeaderRows - 1)
    : rowCursor;
  if (labelColCount > 0) {
    if (pivot.reportLayout === 'compact') {
      const skipTopLeftSingleValue = singleValueBannerOccupied && rowHeaderRowOffset === rowCursor;
      if (!skipTopLeftSingleValue) {
        const vf0 = pivot.rows.length === 0 ? pivot.values[0] : undefined;
        const text = showFieldHeaders
          ? (pivot.rows.length > 0 ? 'Row Labels' : (pivot.values[0]?.customName ?? pivot.name))
          : '';
        cells.push({
          rowOffset: rowHeaderRowOffset, colOffset: 0,
          content: { kind: 'text', value: text },
          style: headerStyle, role: 'header',
          ...(pivot.rows[0] ? { fieldName: pivot.rows[0].fieldName } : vf0 ? { fieldName: vf0.fieldName, valueFieldId: vf0.id } : {}),
        });
      }
    } else {
      for (let i = 0; i < Math.max(1, pivot.rows.length); i++) {
        if (i === 0 && singleValueBannerOccupied && rowHeaderRowOffset === rowCursor) continue;
        cells.push({
          rowOffset: rowHeaderRowOffset, colOffset: i,
          content: { kind: 'text', value: showFieldHeaders ? (pivot.rows[i]?.fieldName ?? 'Row Labels') : '' },
          style: headerStyle, role: 'header', ...(pivot.rows[i] && { fieldName: pivot.rows[i]!.fieldName }),
        });
      }
      if (valuesInRows && hasMultipleValues) {
        cells.push({
          rowOffset: rowHeaderRowOffset, colOffset: labelColCount - 1,
          content: { kind: 'text', value: showFieldHeaders ? 'Values' : '' },
          style: headerStyle, role: 'header',
        });
      }
    }
  }

  const colLabelStartRow = hasColumnHierarchy ? (showFieldHeaders ? rowCursor + 1 : rowCursor) : rowHeaderRowOffset;
  for (let c = 0; c < colLines.length; c++) {
    const colLine = colLines[c]!;
    const colIndex = dataStartCol + c;
    if (colLine.kind === 'grandTotal') {
      const gtVf = hasMultipleValues && !valuesInRows ? pivot.values[colLine.valueIndex] : undefined;
      const text = gtVf
        ? `Total ${gtVf.customName ?? defaultPivotValueLabel(gtVf.fieldName, gtVf.aggregation)}`
        : 'Grand Total';
      cells.push({
        rowOffset: rowHeaderRowOffset,
        colOffset: colIndex,
        content: { kind: 'text', value: text },
        style: headerStyle, role: 'grandTotal',
        ...(gtVf && { fieldName: gtVf.fieldName, valueFieldId: gtVf.id }),
      });
      continue;
    }
    if (hasColumnHierarchy) {
      const topLabel = colLine.node ? colLine.label : '';
      const rowPos = Math.min(rowHeaderRowOffset, colLabelStartRow + (colLine.level ?? 0));
      const cfName = colLine.node?.fieldName ?? pivot.columns[colLine.level ?? 0]?.fieldName;
      const colVf = !valuesInRows && hasMultipleValues ? pivot.values[colLine.valueIndex] : undefined;
      cells.push({
        rowOffset: rowPos, colOffset: colIndex,
        content: { kind: 'text', value: colVf
          ? `${topLabel} - ${colVf.customName ?? colVf.fieldName}`
          : topLabel },
        style: headerStyle, role: 'colLabel',
        ...(cfName && { fieldName: cfName }),
        ...(colVf && { valueFieldId: colVf.id }),
      });
    } else if (hasMultipleValues && !valuesInRows) {
      const vf = pivot.values[colLine.valueIndex]!;
      cells.push({
        rowOffset: rowHeaderRowOffset, colOffset: colIndex,
        content: { kind: 'text', value: vf.customName ?? defaultPivotValueLabel(vf.fieldName, vf.aggregation) },
        style: headerStyle, role: 'colLabel', fieldName: vf.fieldName, valueFieldId: vf.id,
      });
    } else if (pivot.values.length === 1) {
      const vf = pivot.values[0]!;
      cells.push({
        rowOffset: rowHeaderRowOffset, colOffset: colIndex,
        content: { kind: 'text', value: vf.customName ?? defaultPivotValueLabel(vf.fieldName, vf.aggregation) },
        style: headerStyle, role: 'header', fieldName: vf.fieldName, valueFieldId: vf.id,
      });
    }
  }

  let dataRowCursor = rowHeaderRowOffset + 1;
  const maxColCount = dataStartCol + Math.max(1, colLines.length);
  const seenTabularParents = new Set<string>();

  for (let r = 0; r < rowLines.length; r++) {
    const rLine = rowLines[r]!;
    const isTotalRow = rLine.kind === 'subtotal' || rLine.kind === 'grandTotal';
    const rowBand = !isTotalRow && styleOptions.bandedRows && r % 2 === 1;
    const valueSuffix = valuesInRows && hasMultipleValues && pivot.values[rLine.valueIndex]
      ? ` (${pivot.values[rLine.valueIndex]!.customName ?? pivot.values[rLine.valueIndex]!.fieldName})`
      : '';
    if (labelColCount > 0) {
      if (pivot.reportLayout === 'compact') {
        const hasHierarchyButtons = showButtons && pivot.rows.length > 1;
        const indent = rLine.kind === 'grandTotal' ? 0 : rLine.level + (hasHierarchyButtons ? 1 : 0);
        const rowBold = isTotalRow || (styleOptions.rowHeaders && (rLine.level === 0 && pivot.rows.length > 1 || rLine.hasChildren));
        const fill = isTotalRow ? palette.totalFill : rowBand ? palette.bandFill : undefined;
        const cellStyle: CellStyle | undefined = (rowBold || fill || indent > 0 || isTotalRow)
          ? {
            ...(isTotalRow && totalStyleBase),
            ...(rowBold && { bold: true }),
            ...(fill && { fillRgb: fill }),
            ...(indent > 0 && { alignment: 'left', indent }),
          }
          : undefined;
        const compactFieldName = rLine.kind !== 'grandTotal'
          ? (rLine.node?.fieldName ?? rLine.path[rLine.level]?.fieldName ?? pivot.rows[rLine.level]?.fieldName)
          : undefined;
        cells.push({
          rowOffset: dataRowCursor, colOffset: 0,
          content: { kind: 'text', value: `${rLine.label}${valueSuffix}` },
          ...(cellStyle && { style: cellStyle }),
          role: rLine.kind === 'grandTotal' ? 'grandTotal' : rLine.kind === 'subtotal' ? 'subtotal' : 'rowLabel',
          indent: rLine.level,
          ...(rLine.collapsibleKey !== undefined && { collapsibleKey: rLine.collapsibleKey, collapsed: rLine.collapsed }),
          ...(compactFieldName && { fieldName: compactFieldName }),
        });
      } else {
        for (let lvl = 0; lvl < Math.max(1, pivot.rows.length); lvl++) {
          let atLevel = '';
          let attachCollapse: { collapsibleKey: string; collapsed: boolean } | undefined;
          if (rLine.kind === 'grandTotal') {
            atLevel = lvl === 0 ? 'Grand Total' : '';
          } else if (pivot.reportLayout === 'tabular') {
            if (rLine.kind === 'subtotal') {
              atLevel = lvl === rLine.level ? rLine.label : '';
            } else if (lvl === rLine.level) {
              atLevel = rLine.label;
              if (rLine.collapsibleKey !== undefined) {
                attachCollapse = { collapsibleKey: rLine.collapsibleKey, collapsed: rLine.collapsed };
              }
            } else if (lvl < rLine.level && rLine.path[lvl]) {
              const anc = rLine.path[lvl]!;
              const firstSeen = !seenTabularParents.has(anc.key);
              if (firstSeen) {
                seenTabularParents.add(anc.key);
                atLevel = anc.item.displayLabel;
                if (anc.children.length > 0 && pivot.rows.length > lvl + 1) {
                  attachCollapse = { collapsibleKey: anc.key, collapsed: false };
                }
              } else if (repeatLabels) {
                atLevel = anc.item.displayLabel;
              }
            }
          } else {
            atLevel = lvl === rLine.level
              ? rLine.label
              : (repeatLabels && rLine.path[lvl] ? rLine.path[lvl]!.item.displayLabel : '');
            if (lvl === rLine.level && rLine.collapsibleKey !== undefined) {
              attachCollapse = { collapsibleKey: rLine.collapsibleKey, collapsed: rLine.collapsed };
            }
          }
          if (atLevel !== '' || lvl === 0) {
            const rowBold = isTotalRow || (styleOptions.rowHeaders && lvl === 0 && atLevel !== '');
            const indent = showButtons && attachCollapse ? 1 : 0;
            const fill = isTotalRow ? palette.totalFill : rowBand ? palette.bandFill : undefined;
            const cellStyle: CellStyle | undefined = (rowBold || fill || indent > 0 || isTotalRow)
              ? {
                ...(isTotalRow && totalStyleBase),
                ...(rowBold && { bold: true }),
                ...(fill && { fillRgb: fill }),
                ...(indent > 0 && { alignment: 'left', indent }),
              }
              : undefined;
            const rowFieldAtLevel = rLine.kind !== 'grandTotal' ? pivot.rows[lvl]?.fieldName : undefined;
            cells.push({
              rowOffset: dataRowCursor, colOffset: lvl,
              content: { kind: 'text', value: atLevel },
              ...(cellStyle && { style: cellStyle }),
              role: rLine.kind === 'grandTotal' ? 'grandTotal' : rLine.kind === 'subtotal' ? 'subtotal' : 'rowLabel',
              ...attachCollapse,
              ...(rowFieldAtLevel && { fieldName: rowFieldAtLevel }),
            });
          }
        }
        if (valuesInRows && hasMultipleValues) {
          const vf = pivot.values[rLine.valueIndex]!;
          const fill = isTotalRow ? palette.totalFill : rowBand ? palette.bandFill : undefined;
          cells.push({
            rowOffset: dataRowCursor, colOffset: labelColCount - 1,
            content: { kind: 'text', value: vf.customName ?? defaultPivotValueLabel(vf.fieldName, vf.aggregation) },
            ...(isTotalRow || fill ? { style: { ...(isTotalRow && totalStyleBase), ...(fill && { fillRgb: fill }) } } : {}),
            role: 'rowLabel', fieldName: vf.fieldName, valueFieldId: vf.id,
          });
        }
      }
    }

    for (let c = 0; c < colLines.length; c++) {
      const cLine = colLines[c]!;
      const valueIndex = valuesInRows ? rLine.valueIndex : cLine.valueIndex;
      const vf = pivot.values[valueIndex];
      const cellVal = displayMatrix[r]![c]!;
      const isTotal = isTotalRow || cLine.kind === 'grandTotal' || cLine.kind === 'subtotal';
      const colBand = !isTotal && styleOptions.bandedColumns && c % 2 === 1;
      const fill = isTotal ? palette.totalFill : (rowBand || colBand) ? palette.bandFill : undefined;
      const countOrRank = vf?.showValuesAs === 'rankAscending' || vf?.showValuesAs === 'rankDescending' ||
        vf?.showValuesAs === 'index' || vf?.aggregation === 'count' || vf?.aggregation === 'countNumbers';
      const explicitFmt = vf?.numberFormat;
      const numFmt = explicitFmt !== undefined
        ? (/^general$/iu.test(explicitFmt) ? undefined : explicitFmt)
        : (vf?.showValuesAs?.startsWith('percent') ? '0.00%' : countOrRank ? undefined : (vf ? fieldFormats[vf.fieldName] : undefined));
      const style: CellStyle | undefined = isTotal || numFmt || fill
        ? {
          ...(isTotal && totalStyleBase),
          ...(fill && { fillRgb: fill }),
          ...(numFmt && { numberFormat: numFmt }),
        }
        : undefined;
      const { content, cache } = contentAndCache(cellVal);
      const effectiveContent: CellContent = content.kind === 'blank' && !cache && pivot.values.length > 0 && pivot.emptyCellText
        ? { kind: 'text', value: pivot.emptyCellText }
        : content;
      if (pivot.values.length > 0) {
        dataBuckets.set(`${dataRowCursor}:${dataStartCol + c}`, bucketMatrix[r]![c]!);
      }
      if (effectiveContent.kind !== 'blank' || cache || style) {
        cells.push({
          rowOffset: dataRowCursor, colOffset: dataStartCol + c,
          content: effectiveContent, ...(cache && { cache }), ...(style && { style }),
          role: rLine.kind === 'grandTotal' || cLine.kind === 'grandTotal' ? 'grandTotal'
            : rLine.kind === 'subtotal' || cLine.kind === 'subtotal' ? 'subtotal' : 'data',
          ...(vf && { fieldName: vf.fieldName, valueFieldId: vf.id }),
        });
      }
    }
    dataRowCursor++;
    const curRoot = rLine.path[0]?.key;
    const nextLine = rowLines[r + 1];
    const nextRoot = nextLine?.path[0]?.key;
    if (pivot.blankRows && curRoot !== undefined && nextLine && nextLine.kind !== 'grandTotal' && nextRoot !== curRoot) {
      dataRowCursor++;
    }
  }

  return {
    evaluation: {
      rows: Math.max(1, dataRowCursor),
      columns: Math.max(1, maxColCount),
      cells,
      fields,
      recordCount: filteredRecords.length,
    },
    rowHeaderRowOffset,
    headerBannerRow,
    labelColCount,
    dataStartCol,
    dataBuckets,
    rowLines,
    colLines,
    displayMatrix,
    valuesInRows,
    hasMultipleValues,
  };
}

export function evaluatePivotTable(document: WorkbookDocument, pivot: PivotTableDocument): PivotEvaluationResult {
  return evaluatePivotInternal(document, pivot).evaluation;
}

export function computePivotOverlays(
  document: WorkbookDocument,
  sheetId: DocumentId,
): readonly PivotInteractiveOverlay[] {
  const overlays: PivotInteractiveOverlay[] = [];
  for (const pivot of document.pivots ?? []) {
    if (pivot.target.sheetId !== sheetId) continue;
    const internal = evaluatePivotInternal(document, pivot);
    const { cells } = internal.evaluation;
    const toggles: PivotToggleOverlay[] = [];
    const filters: PivotFilterOverlay[] = [];

    if (pivot.showButtons !== false) {
      for (const cell of cells) {
        if (cell.collapsibleKey !== undefined) {
          const text = cell.content.kind === 'text' ? cell.content.value : cell.collapsibleKey;
          toggles.push({
            row: pivot.target.row + cell.rowOffset,
            column: pivot.target.column + cell.colOffset,
            key: cell.collapsibleKey,
            collapsed: Boolean(cell.collapsed),
            label: text,
            level: pivot.reportLayout === 'compact' ? (cell.indent ?? 0) : 0,
          });
        }
      }
    }

    if (pivot.showFieldHeaders !== false) {
      for (const cell of cells) {
        if (cell.role === 'filter' && cell.colOffset === 1 && cell.fieldName) {
          const cfg = pivot.filters.find(f => f.fieldName === cell.fieldName);
          filters.push({
            row: pivot.target.row + cell.rowOffset,
            column: pivot.target.column + cell.colOffset,
            kind: 'reportFilter',
            fieldNames: [cell.fieldName],
            active: Boolean(cfg?.hiddenItems?.length),
            label: cell.fieldName,
          });
        }
      }

      if (pivot.rows.length > 0 && internal.labelColCount > 0) {
        if (pivot.reportLayout === 'compact') {
          const primary = pivot.rows[0]!;
          filters.push({
            row: pivot.target.row + internal.rowHeaderRowOffset,
            column: pivot.target.column,
            kind: 'rowLabel',
            fieldNames: pivot.rows.map(r => r.fieldName),
            active: pivot.rows.some(r => Boolean(r.hiddenItems?.length)),
            ...(primary.sort && { sort: primary.sort.direction }),
            label: 'Row Labels',
          });
        } else {
          for (let i = 0; i < pivot.rows.length; i++) {
            const rf = pivot.rows[i]!;
            filters.push({
              row: pivot.target.row + internal.rowHeaderRowOffset,
              column: pivot.target.column + i,
              kind: 'rowLabel',
              fieldNames: [rf.fieldName],
              active: Boolean(rf.hiddenItems?.length),
              ...(rf.sort && { sort: rf.sort.direction }),
              label: rf.fieldName,
            });
          }
        }
      }

      if (pivot.columns.length > 0) {
        const primary = pivot.columns[0]!;
        filters.push({
          row: pivot.target.row + internal.headerBannerRow,
          column: pivot.target.column + internal.dataStartCol,
          kind: 'colLabel',
          fieldNames: pivot.columns.map(c => c.fieldName),
          active: pivot.columns.some(c => Boolean(c.hiddenItems?.length)),
          ...(primary.sort && { sort: primary.sort.direction }),
          label: 'Column Labels',
        });
      }
    }

    overlays.push({ pivotId: pivot.id, toggles, filters });
  }
  return overlays;
}

export function extractPivotCellDrillDown(
  document: WorkbookDocument,
  pivot: PivotTableDocument,
  target: CellReference,
): { readonly fields: readonly string[]; readonly records: readonly PivotRecord[] } | undefined {
  if (target.sheetId !== pivot.target.sheetId) return undefined;
  const rowOffset = target.row - pivot.target.row;
  const colOffset = target.column - pivot.target.column;
  if (rowOffset < 0 || colOffset < 0) return undefined;
  const internal = evaluatePivotInternal(document, pivot);
  const records = internal.dataBuckets.get(`${rowOffset}:${colOffset}`);
  if (!records) return undefined;
  return { fields: internal.evaluation.fields, records };
}

export function collectFieldCollapsibleKeys(
  document: WorkbookDocument,
  pivot: PivotTableDocument,
  fieldName: string,
): readonly string[] {
  const fieldIndex = pivot.rows.findIndex(rf => rf.fieldName === fieldName);
  if (fieldIndex < 0 || fieldIndex >= pivot.rows.length - 1) return [];
  const { records: allRecords, fieldFormats } = extractPivotRecords(document, pivot.source);
  const visibleRecords = allRecords.filter(record =>
    pivot.filters.every(filter => !filter.hiddenItems?.length || !filter.hiddenItems.includes(formatPivotItemLabel(record[filter.fieldName] ?? null))) &&
    pivot.rows.every(rf => !rf.hiddenItems?.length || !rf.hiddenItems.includes(formatPivotItemLabel(record[rf.fieldName] ?? null))) &&
    pivot.columns.every(cf => !cf.hiddenItems?.length || !cf.hiddenItems.includes(formatPivotItemLabel(record[cf.fieldName] ?? null))),
  );
  const rowTree = buildAxisTree(visibleRecords, pivot.rows, pivot.values, fieldFormats);
  const keys: string[] = [];
  const walk = (nodes: readonly AxisNode[]) => {
    for (const node of nodes) {
      if (node.level === fieldIndex && node.fieldName === fieldName && node.children.length > 0) {
        keys.push(node.key);
      } else if (node.level < fieldIndex) {
        walk(node.children);
      }
    }
  };
  walk(rowTree);
  return keys;
}

export function resolveActivePivotField(
  document: WorkbookDocument,
  pivot: PivotTableDocument,
  target: CellReference,
): ActivePivotFieldInfo | undefined {
  if (target.sheetId !== pivot.target.sheetId) return undefined;
  const rowOffset = target.row - pivot.target.row;
  const colOffset = target.column - pivot.target.column;
  if (rowOffset < 0 || colOffset < 0) return undefined;
  const internal = evaluatePivotInternal(document, pivot);
  if (rowOffset >= internal.evaluation.rows || colOffset >= internal.evaluation.columns) return undefined;

  const cell = internal.evaluation.cells.find(c => c.rowOffset === rowOffset && c.colOffset === colOffset);
  if (cell?.role === 'filter' && cell.fieldName) {
    return {
      displayName: cell.fieldName,
      fieldName: cell.fieldName,
      area: 'filters',
      canExpandCollapse: false,
    };
  }
  if (cell?.valueFieldId) {
    const vf = pivot.values.find(v => v.id === cell.valueFieldId);
    if (vf) {
      return {
        displayName: vf.customName ?? defaultPivotValueLabel(vf.fieldName, vf.aggregation),
        fieldName: vf.fieldName,
        area: 'values',
        valueField: vf,
        canExpandCollapse: false,
      };
    }
  }
  if (cell?.fieldName) {
    const rowIdx = pivot.rows.findIndex(rf => rf.fieldName === cell.fieldName);
    if (rowIdx >= 0 && (cell.role === 'rowLabel' || cell.role === 'subtotal' || (cell.role === 'header' && colOffset < internal.dataStartCol))) {
      const axisField = pivot.rows[rowIdx]!;
      return {
        displayName: axisField.fieldName,
        fieldName: axisField.fieldName,
        area: 'rows',
        axisField,
        canExpandCollapse: rowIdx < pivot.rows.length - 1,
      };
    }
    const colIdx = pivot.columns.findIndex(cf => cf.fieldName === cell.fieldName);
    if (colIdx >= 0) {
      const axisField = pivot.columns[colIdx]!;
      return {
        displayName: axisField.fieldName,
        fieldName: axisField.fieldName,
        area: 'columns',
        axisField,
        canExpandCollapse: false,
      };
    }
  }

  if (colOffset >= internal.dataStartCol && pivot.values.length > 0) {
    const vf = pivot.values[0]!;
    return {
      displayName: vf.customName ?? defaultPivotValueLabel(vf.fieldName, vf.aggregation),
      fieldName: vf.fieldName,
      area: 'values',
      valueField: vf,
      canExpandCollapse: false,
    };
  }
  if (pivot.rows.length > 0) {
    const rowIdx = pivot.reportLayout !== 'compact' && colOffset < pivot.rows.length ? colOffset : 0;
    const axisField = pivot.rows[rowIdx]!;
    return {
      displayName: axisField.fieldName,
      fieldName: axisField.fieldName,
      area: 'rows',
      axisField,
      canExpandCollapse: rowIdx < pivot.rows.length - 1,
    };
  }
  if (pivot.columns.length > 0) {
    const axisField = pivot.columns[0]!;
    return {
      displayName: axisField.fieldName,
      fieldName: axisField.fieldName,
      area: 'columns',
      axisField,
      canExpandCollapse: false,
    };
  }
  if (pivot.filters.length > 0) {
    const filterField = pivot.filters[0]!;
    return {
      displayName: filterField.fieldName,
      fieldName: filterField.fieldName,
      area: 'filters',
      canExpandCollapse: false,
    };
  }
  if (pivot.values.length > 0) {
    const vf = pivot.values[0]!;
    return {
      displayName: vf.customName ?? defaultPivotValueLabel(vf.fieldName, vf.aggregation),
      fieldName: vf.fieldName,
      area: 'values',
      valueField: vf,
      canExpandCollapse: false,
    };
  }
  return undefined;
}

export function findPivotContainingCell(
  document: WorkbookDocument,
  target: CellReference,
): PivotTableDocument | undefined {
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

function matchesAxisItem(node: AxisNode, query: PivotScalar): boolean {
  if (typeof query === 'object' && query !== null) return false;
  if (typeof query === 'string') return query.toLowerCase() === node.item.label.toLowerCase();
  if (typeof query === 'number') {
    return (typeof node.item.raw === 'number' && node.item.raw === query) || String(query) === node.item.label;
  }
  if (typeof query === 'boolean') {
    return (typeof node.item.raw === 'boolean' && node.item.raw === query) ||
      (query ? 'TRUE' : 'FALSE') === node.item.label.toUpperCase();
  }
  return node.item.raw === null || node.item.label === '';
}

function matchesAxisLine(
  line: AxisLine,
  queries: ReadonlyMap<string, PivotScalar>,
  axisLength: number,
): boolean {
  if (queries.size === 0) {
    return axisLength === 0 ? line.path.length === 0 : line.kind === 'grandTotal';
  }
  if (line.kind === 'item' && line.hasChildren && !line.collapsed) return false;
  if (line.path.length !== queries.size) return false;
  return line.path.every(node => {
    const fieldKey = node.fieldName.toLowerCase();
    return queries.has(fieldKey) && matchesAxisItem(node, queries.get(fieldKey)!);
  });
}

export function queryPivotData(
  document: WorkbookDocument,
  pivot: PivotTableDocument,
  dataField: string,
  pairs: readonly (readonly [string, PivotScalar])[],
  cache?: PivotEvaluationCache,
): PivotScalar {
  const refErr: Extract<CellCache, { kind: 'error' }> = { kind: 'error', value: '#REF!' };
  if (pivot.values.length === 0) return refErr;

  const targetName = dataField.trim().toLowerCase();
  if (!targetName) return refErr;
  const valueIdx = pivot.values.findIndex(v =>
    v.customName?.trim().toLowerCase() === targetName ||
    defaultPivotValueLabel(v.fieldName, v.aggregation).toLowerCase() === targetName ||
    v.fieldName.toLowerCase() === targetName,
  );
  if (valueIdx < 0) return refErr;

  const rowFieldKeys = new Set(pivot.rows.map(rf => rf.fieldName.toLowerCase()));
  const colFieldKeys = new Set(pivot.columns.map(cf => cf.fieldName.toLowerCase()));
  const seenFields = new Set<string>();
  const rowQueries = new Map<string, PivotScalar>();
  const colQueries = new Map<string, PivotScalar>();

  for (const [rawField, rawItem] of pairs) {
    if (typeof rawItem === 'object' && rawItem !== null) return rawItem;
    const key = rawField.trim().toLowerCase();
    if (!key || seenFields.has(key)) return refErr;
    seenFields.add(key);
    if (rowFieldKeys.has(key)) rowQueries.set(key, rawItem);
    else if (colFieldKeys.has(key)) colQueries.set(key, rawItem);
    else return refErr;
  }

  let internal = cache?.get(pivot.id);
  if (!internal) {
    internal = evaluatePivotInternal(document, pivot);
    cache?.set(pivot.id, internal);
  }

  const rIdx = internal.rowLines.findIndex(line =>
    (!internal!.hasMultipleValues || !internal!.valuesInRows || line.valueIndex === valueIdx) &&
    matchesAxisLine(line, rowQueries, pivot.rows.length),
  );
  if (rIdx < 0) return refErr;

  const cIdx = internal.colLines.findIndex(line =>
    (!internal!.hasMultipleValues || internal!.valuesInRows || line.valueIndex === valueIdx) &&
    matchesAxisLine(line, colQueries, pivot.columns.length),
  );
  if (cIdx < 0) return refErr;

  const cellVal = internal.displayMatrix[rIdx]?.[cIdx];
  if (!cellVal || cellVal.kind === 'blank') return 0;
  if (cellVal.kind === 'error') return cellVal;
  return cellVal.value;
}
