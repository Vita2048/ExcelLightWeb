import {
  chartCategoryOf,
  type ChartDocument,
  type ChartSeries,
  type ChartSource,
  type ChartType,
} from '../core/chart';
import { displayCellValue, formatCellValue, formatNumber } from '../core/formats/cell-format';
import {
  columnLetters,
  parseFormula,
  quoteSheet,
  referenceBox,
  sameSheet,
  type FormulaReference,
  type StructuredReference,
} from '../core/formula-references';
import { maskRows, RowMask } from '../core/row-mask';
import {
  documentLimits,
  findCell,
  resolvedCellStyle,
  type CellDocument,
  type CellExtent,
  type Coordinate,
  type DocumentId,
  type SheetDocument,
  type WorkbookDocument,
  type WorkbookName,
} from '../core/workbook';
import { lastColumn, lastRow, tableDesign } from '../core/workbook-names';
import { ordinalNameKey } from '../core/workbook-validation';

export interface ResolvedSheetArea {
  readonly sheetId: DocumentId;
  readonly row: number;
  readonly column: number;
  readonly rows: number;
  readonly columns: number;
}

export interface InferredChartSeries {
  readonly name?: string;
  readonly nameRef?: string;
  readonly categoriesRef?: string;
  readonly valuesRef: string;
  readonly xValuesRef?: string;
  readonly bubbleSizesRef?: string;
}

export interface InferredChartData {
  readonly seriesBy: 'columns' | 'rows';
  readonly series: readonly InferredChartSeries[];
  readonly area?: ResolvedSheetArea;
}

export interface ResolvedChartPoint {
  readonly categoryIndex: number;
  readonly value: number | null;
  readonly raw: string | number | boolean | null;
  readonly formatted: string;
  readonly numberFormat?: string;
  readonly x?: number | null;
  readonly size?: number | null;
  readonly isBlank: boolean;
}

export interface ResolvedChartSeries {
  readonly id: DocumentId;
  readonly name: string;
  readonly chartType: ChartType;
  readonly secondaryAxis: boolean;
  readonly hidden: boolean;
  readonly connectGaps: boolean;
  readonly points: readonly ResolvedChartPoint[];
  readonly valuesFormat?: string;
  readonly sourceSeries: ChartSeries;
}

export interface ResolvedChartData {
  readonly title: string;
  readonly categories: readonly string[];
  readonly series: readonly ResolvedChartSeries[];
  readonly visibleCategoryIndices: readonly number[];
}

function sheetById(document: WorkbookDocument, sheetId: DocumentId): SheetDocument | undefined {
  return document.sheets.find(sheet => sheet.id === sheetId);
}

function sheetByName(document: WorkbookDocument, name: string): SheetDocument | undefined {
  return document.sheets.find(sheet => sameSheet(sheet.name, name));
}

export function formatAbsoluteAreaRef(
  sheetName: string,
  row: number,
  column: number,
  rows: number,
  columns: number,
): string {
  const start = `$${columnLetters(column)}$${row}`;
  if (rows <= 1 && columns <= 1) return `=${quoteSheet(sheetName)}${start}`;
  const end = `$${columnLetters(column + columns - 1)}$${row + rows - 1}`;
  return `=${quoteSheet(sheetName)}${start}:${end}`;
}

function areaFromReference(
  document: WorkbookDocument,
  fallbackSheetId: DocumentId,
  ref: FormulaReference,
): ResolvedSheetArea | undefined {
  const targetSheet = ref.sheet ? sheetByName(document, ref.sheet) : sheetById(document, fallbackSheetId);
  if (!targetSheet) return undefined;
  const box = referenceBox(ref);
  const top = Math.max(1, Math.min(documentLimits.rows, box.top));
  const left = Math.max(1, Math.min(documentLimits.columns, box.left));
  const bottom = Math.max(top, Math.min(documentLimits.rows, box.bottom));
  const right = Math.max(left, Math.min(documentLimits.columns, box.right));
  return {
    sheetId: targetSheet.id,
    row: top,
    column: left,
    rows: bottom - top + 1,
    columns: right - left + 1,
  };
}

function tableColumnHeaders(document: WorkbookDocument, table: WorkbookName): readonly string[] {
  const design = tableDesign(table);
  if (!design.headerRow) {
    return table.headerLabels ?? Array.from({ length: table.area.columns }, (_, i) => `Column${i + 1}`);
  }
  const sheet = sheetById(document, table.area.sheetId);
  if (!sheet) return Array.from({ length: table.area.columns }, (_, i) => `Column${i + 1}`);
  return Array.from({ length: table.area.columns }, (_, i) => {
    const cell = findCell(sheet, { row: table.area.row, column: table.area.column + i });
    const text = formatCellValue(cell).text.trim();
    return text || `Column${i + 1}`;
  });
}

function resolveStructuredExtent(
  document: WorkbookDocument,
  ref: StructuredReference,
): ResolvedSheetArea | undefined {
  if (!ref.table) return undefined;
  const key = ordinalNameKey(ref.table);
  const table = document.names.find(n => n.isTable && !n.area.invalid && ordinalNameKey(n.name) === key);
  if (!table) return undefined;
  const design = tableDesign(table);
  const hasHeaders = design.headerRow;
  const hasTotals = design.totalRow;
  const totalRows = table.area.rows;
  let startRow = table.area.row + (hasHeaders ? 1 : 0);
  let endRow = lastRow(table.area) - (hasTotals ? 1 : 0);
  const specs = new Set(ref.specifiers);
  if (specs.has('#All')) {
    startRow = table.area.row;
    endRow = lastRow(table.area);
  } else if (specs.has('#Headers')) {
    if (!hasHeaders) return undefined;
    startRow = table.area.row;
    endRow = table.area.row;
  } else if (specs.has('#Totals')) {
    if (!hasTotals) return undefined;
    startRow = lastRow(table.area);
    endRow = lastRow(table.area);
  } else if (totalRows <= (hasHeaders ? 1 : 0) + (hasTotals ? 1 : 0)) {
    return undefined;
  }
  if (endRow < startRow) return undefined;

  let startCol = table.area.column;
  let endCol = lastColumn(table.area);
  if (ref.startColumn !== undefined) {
    const headers = tableColumnHeaders(document, table);
    const startIdx = headers.findIndex(h => ordinalNameKey(h) === ordinalNameKey(ref.startColumn!));
    if (startIdx < 0) return undefined;
    const endIdx = ref.endColumn !== undefined
      ? headers.findIndex(h => ordinalNameKey(h) === ordinalNameKey(ref.endColumn!))
      : startIdx;
    if (endIdx < 0) return undefined;
    startCol = table.area.column + Math.min(startIdx, endIdx);
    endCol = table.area.column + Math.max(startIdx, endIdx);
  }
  return {
    sheetId: table.area.sheetId,
    row: startRow,
    column: startCol,
    rows: endRow - startRow + 1,
    columns: endCol - startCol + 1,
  };
}

export function resolveFormulaExtent(
  document: WorkbookDocument,
  fallbackSheetId: DocumentId,
  formula: string,
): ResolvedSheetArea | undefined {
  const normalized = formula.trim().startsWith('=') ? formula.trim() : `=${formula.trim()}`;
  if (normalized.length <= 1 || normalized === '=#REF!') return undefined;
  let parsed: ReturnType<typeof parseFormula>;
  try {
    parsed = parseFormula(normalized);
  } catch {
    return undefined;
  }
  const expr = parsed.expression;
  if (expr.kind === 'reference') {
    return areaFromReference(document, fallbackSheetId, expr.reference);
  }
  if (expr.kind === 'spill' && expr.target.kind === 'reference') {
    const anchorArea = areaFromReference(document, fallbackSheetId, expr.target.reference);
    if (!anchorArea) return undefined;
    const sheet = sheetById(document, anchorArea.sheetId);
    const anchor = sheet && findCell(sheet, { row: anchorArea.row, column: anchorArea.column });
    if (anchor?.spillExtent) {
      return {
        sheetId: anchorArea.sheetId,
        row: anchorArea.row,
        column: anchorArea.column,
        rows: anchor.spillExtent.rows,
        columns: anchor.spillExtent.columns,
      };
    }
    if (anchor?.content.kind === 'python' && anchor.content.block?.spillActive && anchor.content.block.lastTable) {
      return {
        sheetId: anchorArea.sheetId,
        row: anchorArea.row,
        column: anchorArea.column,
        rows: anchor.content.block.lastTable.rows,
        columns: anchor.content.block.lastTable.columns,
      };
    }
    return anchorArea;
  }
  if (expr.kind === 'name') {
    const key = ordinalNameKey(expr.name);
    const entry = document.names.find(n => !n.area.invalid && ordinalNameKey(n.name) === key);
    if (!entry) return undefined;
    return {
      sheetId: entry.area.sheetId,
      row: entry.area.row,
      column: entry.area.column,
      rows: entry.area.rows,
      columns: entry.area.columns,
    };
  }
  if (expr.kind === 'structured') {
    return resolveStructuredExtent(document, expr.reference);
  }
  return undefined;
}

export function resolveChartSourceArea(
  document: WorkbookDocument,
  source: ChartSource,
  fallbackSheetId: DocumentId,
): ResolvedSheetArea | undefined {
  if (source.kind === 'range') {
    return resolveFormulaExtent(document, fallbackSheetId, source.formula);
  }
  if (source.kind === 'table') {
    const table = document.names.find(n => n.id === source.tableId && n.isTable && !n.area.invalid);
    if (!table) return undefined;
    const design = tableDesign(table);
    const dataRows = Math.max(1, table.area.rows - (design.totalRow ? 1 : 0));
    return {
      sheetId: table.area.sheetId,
      row: table.area.row,
      column: table.area.column,
      rows: dataRows,
      columns: table.area.columns,
    };
  }
  const pivot = (document.pivots ?? []).find(p => p.id === source.pivotId);
  if (!pivot || !pivot.extent) return undefined;
  const filterOffset = pivot.filters.length > 0 ? pivot.filters.length + 1 : 0;
  const row = pivot.extent.row + filterOffset;
  const rows = Math.max(1, pivot.extent.rows - filterOffset - (pivot.grandTotals.columns ? 1 : 0));
  const columns = Math.max(1, pivot.extent.columns - (pivot.grandTotals.rows && pivot.columns.length > 0 ? 1 : 0));
  return {
    sheetId: pivot.target.sheetId,
    row,
    column: pivot.extent.column,
    rows,
    columns,
  };
}

function readCellNumericValue(cell: CellDocument | undefined): {
  readonly value: number | null;
  readonly raw: string | number | boolean | null;
  readonly isBlank: boolean;
} {
  const displayed = displayCellValue(cell);
  if (!displayed || displayed.kind === 'blank') {
    return { value: null, raw: null, isBlank: true };
  }
  if (displayed.kind === 'number') {
    return Number.isFinite(displayed.value)
      ? { value: displayed.value, raw: displayed.value, isBlank: false }
      : { value: null, raw: null, isBlank: false };
  }
  if (displayed.kind === 'boolean') {
    return { value: displayed.value ? 1 : 0, raw: displayed.value, isBlank: false };
  }
  if (displayed.kind === 'text') {
    return { value: null, raw: displayed.value, isBlank: false };
  }
  if (displayed.kind === 'error') {
    return { value: null, raw: displayed.value, isBlank: false };
  }
  return { value: null, raw: null, isBlank: true };
}

function readCellFormattedText(sheet: SheetDocument, point: Coordinate): string {
  const cell = findCell(sheet, point);
  const style = resolvedCellStyle(sheet, point, cell);
  return formatCellValue(cell, style?.numberFormat).text;
}

function isCellNumericOrBlank(sheet: SheetDocument, point: Coordinate): 'numeric' | 'text' | 'blank' {
  const displayed = displayCellValue(findCell(sheet, point));
  if (!displayed || displayed.kind === 'blank') return 'blank';
  if (displayed.kind === 'number') return 'numeric';
  return 'text';
}

export function inferChartSeriesFromSource(
  document: WorkbookDocument,
  source: ChartSource,
  fallbackSheetId: DocumentId,
  chartType: ChartType = 'clusteredColumn',
  preferredSeriesBy?: 'columns' | 'rows',
): InferredChartData {
  const area = resolveChartSourceArea(document, source, fallbackSheetId);
  if (!area) {
    return { seriesBy: preferredSeriesBy ?? 'columns', series: [] };
  }
  const sheet = sheetById(document, area.sheetId);
  if (!sheet) {
    return { seriesBy: preferredSeriesBy ?? 'columns', series: [] };
  }

  let hasTopHeaders = false;
  if (area.rows >= 2) {
    let topText = 0;
    let bodyNumeric = 0;
    for (let c = 0; c < area.columns; c++) {
      const topKind = isCellNumericOrBlank(sheet, { row: area.row, column: area.column + c });
      const nextKind = isCellNumericOrBlank(sheet, { row: area.row + 1, column: area.column + c });
      if (topKind === 'text') topText++;
      if (nextKind === 'numeric') bodyNumeric++;
    }
    hasTopHeaders = topText > 0 && (bodyNumeric > 0 || topText >= Math.max(1, area.columns - 1));
  }

  let hasLeftLabels = false;
  if (area.columns >= 2) {
    let leftText = 0;
    let rightNumeric = 0;
    const startOffset = hasTopHeaders ? 1 : 0;
    for (let r = startOffset; r < area.rows; r++) {
      const leftKind = isCellNumericOrBlank(sheet, { row: area.row + r, column: area.column });
      const nextKind = isCellNumericOrBlank(sheet, { row: area.row + r, column: area.column + 1 });
      if (leftKind === 'text') leftText++;
      if (nextKind === 'numeric') rightNumeric++;
    }
    hasLeftLabels = leftText > 0 && rightNumeric > 0;
  }

  const dataRowStart = area.row + (hasTopHeaders ? 1 : 0);
  const dataRowCount = Math.max(1, area.rows - (hasTopHeaders ? 1 : 0));
  const dataColStart = area.column + (hasLeftLabels ? 1 : 0);
  const dataColCount = Math.max(1, area.columns - (hasLeftLabels ? 1 : 0));

  const category = chartCategoryOf(chartType);
  const seriesBy: 'columns' | 'rows' = preferredSeriesBy ??
    (category === 'scatter' ? 'columns' : dataRowCount >= dataColCount ? 'columns' : 'rows');

  if (category === 'scatter') {
    const xCol = area.column;
    const firstYCol = area.columns >= 2 ? area.column + 1 : area.column;
    const xValuesRef = formatAbsoluteAreaRef(sheet.name, dataRowStart, xCol, dataRowCount, 1);
    const series: InferredChartSeries[] = [];
    if (chartType === 'bubble') {
      for (let col = firstYCol; col < area.column + area.columns; col += 2) {
        const name = hasTopHeaders ? readCellFormattedText(sheet, { row: area.row, column: col }) : `Series ${series.length + 1}`;
        const nameRef = hasTopHeaders ? formatAbsoluteAreaRef(sheet.name, area.row, col, 1, 1) : undefined;
        const valuesRef = formatAbsoluteAreaRef(sheet.name, dataRowStart, col, dataRowCount, 1);
        const sizeCol = col + 1 < area.column + area.columns ? col + 1 : undefined;
        const bubbleSizesRef = sizeCol !== undefined
          ? formatAbsoluteAreaRef(sheet.name, dataRowStart, sizeCol, dataRowCount, 1)
          : undefined;
        series.push({
          name: name || `Series ${series.length + 1}`,
          ...(nameRef && { nameRef }),
          xValuesRef,
          valuesRef,
          ...(bubbleSizesRef && { bubbleSizesRef }),
        });
      }
    } else {
      for (let col = firstYCol; col < area.column + area.columns; col++) {
        const name = hasTopHeaders ? readCellFormattedText(sheet, { row: area.row, column: col }) : `Series ${series.length + 1}`;
        const nameRef = hasTopHeaders ? formatAbsoluteAreaRef(sheet.name, area.row, col, 1, 1) : undefined;
        const valuesRef = formatAbsoluteAreaRef(sheet.name, dataRowStart, col, dataRowCount, 1);
        series.push({
          name: name || `Series ${series.length + 1}`,
          ...(nameRef && { nameRef }),
          xValuesRef,
          valuesRef,
        });
      }
    }
    return { seriesBy: 'columns', series, area };
  }

  if (seriesBy === 'columns') {
    const categoriesRef = hasLeftLabels
      ? formatAbsoluteAreaRef(sheet.name, dataRowStart, area.column, dataRowCount, 1)
      : undefined;
    const series: InferredChartSeries[] = [];
    for (let offset = 0; offset < dataColCount; offset++) {
      const col = dataColStart + offset;
      const name = hasTopHeaders ? readCellFormattedText(sheet, { row: area.row, column: col }) : '';
      const nameRef = hasTopHeaders ? formatAbsoluteAreaRef(sheet.name, area.row, col, 1, 1) : undefined;
      const valuesRef = formatAbsoluteAreaRef(sheet.name, dataRowStart, col, dataRowCount, 1);
      series.push({
        name: name || `Series ${series.length + 1}`,
        ...(nameRef && { nameRef }),
        ...(categoriesRef && { categoriesRef }),
        valuesRef,
      });
    }
    return { seriesBy, series, area };
  }

  const categoriesRef = hasTopHeaders
    ? formatAbsoluteAreaRef(sheet.name, area.row, dataColStart, 1, dataColCount)
    : undefined;
  const series: InferredChartSeries[] = [];
  for (let offset = 0; offset < dataRowCount; offset++) {
    const row = dataRowStart + offset;
    const name = hasLeftLabels ? readCellFormattedText(sheet, { row, column: area.column }) : '';
    const nameRef = hasLeftLabels ? formatAbsoluteAreaRef(sheet.name, row, area.column, 1, 1) : undefined;
    const valuesRef = formatAbsoluteAreaRef(sheet.name, row, dataColStart, 1, dataColCount);
    series.push({
      name: name || `Series ${series.length + 1}`,
      ...(nameRef && { nameRef }),
      ...(categoriesRef && { categoriesRef }),
      valuesRef,
    });
  }
  return { seriesBy, series, area };
}

function sheetFilterMask(sheet: SheetDocument): RowMask {
  const filter = sheet.filter;
  if (!filter || filter.criteria.length === 0) return RowMask.none();
  const from = filter.row + 1;
  const to = lastRow(filter);
  const columns = new Set(filter.criteria.map(criterion => criterion.column));
  const rules = filter.criteria.map(criterion => ({ column: criterion.column, allowed: new Set(criterion.allowed) }));
  const cells = [];
  for (const cell of sheet.cells) {
    if (cell.row < from || cell.row > to || !columns.has(cell.column)) continue;
    const key = formatCellValue(cell, resolvedCellStyle(sheet, cell, cell)?.numberFormat).text;
    cells.push({ row: cell.row, column: cell.column, key });
  }
  return maskRows(from, to, rules, '', cells);
}

interface SheetVisibility {
  readonly hiddenRows: ReadonlySet<number>;
  readonly hiddenCols: ReadonlySet<number>;
  readonly filterMask: RowMask;
}

function getSheetVisibility(
  cache: Map<DocumentId, SheetVisibility>,
  sheet: SheetDocument,
): SheetVisibility {
  let entry = cache.get(sheet.id);
  if (entry) return entry;
  entry = {
    hiddenRows: new Set(sheet.rows.filter(r => r.hidden).map(r => r.index)),
    hiddenCols: new Set(sheet.columns.filter(c => c.hidden).map(c => c.index)),
    filterMask: sheetFilterMask(sheet),
  };
  cache.set(sheet.id, entry);
  return entry;
}

function enumerateVisibleCoordinates(
  document: WorkbookDocument,
  visCache: Map<DocumentId, SheetVisibility>,
  area: ResolvedSheetArea,
  showHiddenCells: boolean,
): readonly Coordinate[] {
  const sheet = sheetById(document, area.sheetId);
  if (!sheet) return [];
  const vis = showHiddenCells ? undefined : getSheetVisibility(visCache, sheet);
  const coords: Coordinate[] = [];
  for (let r = area.row; r < area.row + area.rows; r++) {
    if (vis && (vis.hiddenRows.has(r) || vis.filterMask.has(r))) continue;
    for (let c = area.column; c < area.column + area.columns; c++) {
      if (vis && vis.hiddenCols.has(c)) continue;
      coords.push({ row: r, column: c });
    }
  }
  return coords;
}

function resolveChartTitle(document: WorkbookDocument, chart: ChartDocument): string {
  const title = chart.elements.title;
  if (!title || !title.show) return '';
  if (title.formulaRef) {
    const area = resolveFormulaExtent(document, chart.placement.sheetId, title.formulaRef);
    if (area) {
      const sheet = sheetById(document, area.sheetId);
      if (sheet) {
        const text = readCellFormattedText(sheet, { row: area.row, column: area.column });
        if (text) return text;
      }
    }
  }
  return title.text ?? '';
}

export function resolveChartData(document: WorkbookDocument, chart: ChartDocument): ResolvedChartData {
  const visCache = new Map<DocumentId, SheetVisibility>();
  const hiddenCatSet = new Set(chart.hiddenCategories ?? []);

  const effectiveSeries: readonly ChartSeries[] = chart.series.length > 0
    ? chart.series
    : inferChartSeriesFromSource(
      document,
      chart.source,
      chart.placement.sheetId,
      chart.chartType,
      chart.seriesBy,
    ).series.map((item, idx) => ({
      id: `${chart.id}:auto:${idx}`,
      ...item,
    }));

  // Resolve categories from the first series with categoriesRef, or numeric 1..N fallback.
  let rawCategories: string[] = [];
  const firstCatRef = effectiveSeries.find(s => s.categoriesRef)?.categoriesRef;
  if (firstCatRef) {
    const catArea = resolveFormulaExtent(document, chart.placement.sheetId, firstCatRef);
    if (catArea) {
      const sheet = sheetById(document, catArea.sheetId);
      if (sheet) {
        const coords = enumerateVisibleCoordinates(document, visCache, catArea, chart.showHiddenCells);
        rawCategories = coords.map(pt => readCellFormattedText(sheet, pt));
      }
    }
  }

  const resolvedSeriesRaw: {
    series: ChartSeries;
    name: string;
    valuesFormat?: string;
    points: ResolvedChartPoint[];
  }[] = [];

  for (let sIdx = 0; sIdx < effectiveSeries.length; sIdx++) {
    const s = effectiveSeries[sIdx]!;
    let name = s.name ?? `Series ${sIdx + 1}`;
    if (s.nameRef) {
      const nameArea = resolveFormulaExtent(document, chart.placement.sheetId, s.nameRef);
      if (nameArea) {
        const sheet = sheetById(document, nameArea.sheetId);
        if (sheet) {
          const text = readCellFormattedText(sheet, { row: nameArea.row, column: nameArea.column });
          if (text) name = text;
        }
      }
    }

    const valArea = resolveFormulaExtent(document, chart.placement.sheetId, s.valuesRef);
    const valSheet = valArea ? sheetById(document, valArea.sheetId) : undefined;
    const valCoords = valArea && valSheet
      ? enumerateVisibleCoordinates(document, visCache, valArea, chart.showHiddenCells)
      : [];

    const xArea = s.xValuesRef ? resolveFormulaExtent(document, chart.placement.sheetId, s.xValuesRef) : undefined;
    const xSheet = xArea ? sheetById(document, xArea.sheetId) : undefined;
    const xCoords = xArea && xSheet
      ? enumerateVisibleCoordinates(document, visCache, xArea, chart.showHiddenCells)
      : [];

    const sizeArea = s.bubbleSizesRef
      ? resolveFormulaExtent(document, chart.placement.sheetId, s.bubbleSizesRef)
      : undefined;
    const sizeSheet = sizeArea ? sheetById(document, sizeArea.sheetId) : undefined;
    const sizeCoords = sizeArea && sizeSheet
      ? enumerateVisibleCoordinates(document, visCache, sizeArea, chart.showHiddenCells)
      : [];

    let firstFormat: string | undefined;
    const points: ResolvedChartPoint[] = [];
    for (let i = 0; i < valCoords.length; i++) {
      if (hiddenCatSet.has(i)) continue;
      const pt = valCoords[i]!;
      const cell = valSheet ? findCell(valSheet, pt) : undefined;
      const style = valSheet ? resolvedCellStyle(valSheet, pt, cell) : undefined;
      if (firstFormat === undefined && style?.numberFormat) firstFormat = style.numberFormat;
      const extracted = readCellNumericValue(cell);
      let value = extracted.value;
      if (extracted.isBlank && chart.hiddenEmptyMode === 'zero') {
        value = 0;
      }
      const formatted = value !== null
        ? formatNumber(value, style?.numberFormat).text
        : extracted.isBlank && chart.hiddenEmptyMode === 'zero'
          ? formatNumber(0, style?.numberFormat).text
          : formatCellValue(cell, style?.numberFormat).text;

      let x: number | null | undefined;
      if (xCoords.length > 0 && xSheet) {
        const xPt = xCoords[i];
        x = xPt ? readCellNumericValue(findCell(xSheet, xPt)).value : null;
      }
      let size: number | null | undefined;
      if (sizeCoords.length > 0 && sizeSheet) {
        const sizePt = sizeCoords[i];
        size = sizePt ? readCellNumericValue(findCell(sizeSheet, sizePt)).value : null;
      }

      points.push({
        categoryIndex: i,
        value,
        raw: extracted.raw,
        formatted,
        ...(style?.numberFormat && { numberFormat: style.numberFormat }),
        ...(x !== undefined && { x }),
        ...(size !== undefined && { size }),
        isBlank: extracted.isBlank,
      });
    }

    resolvedSeriesRaw.push({
      series: s,
      name,
      ...(firstFormat !== undefined && { valuesFormat: firstFormat }),
      points,
    });
  }

  const maxPoints = resolvedSeriesRaw.reduce(
    (max, item) => Math.max(max, item.points.reduce((m, p) => Math.max(m, p.categoryIndex + 1), 0)),
    rawCategories.length,
  );
  const categories: string[] = [];
  const visibleCategoryIndices: number[] = [];
  for (let i = 0; i < maxPoints; i++) {
    if (hiddenCatSet.has(i)) continue;
    visibleCategoryIndices.push(i);
    categories.push(rawCategories[i] ?? String(i + 1));
  }

  return {
    title: resolveChartTitle(document, chart),
    categories,
    visibleCategoryIndices,
    series: resolvedSeriesRaw.map(item => ({
      id: item.series.id,
      name: item.name,
      chartType: item.series.chartType ?? chart.chartType,
      secondaryAxis: item.series.secondaryAxis ?? false,
      hidden: item.series.hidden ?? false,
      connectGaps: chart.hiddenEmptyMode === 'connect',
      points: item.points,
      ...(item.valuesFormat && { valuesFormat: item.valuesFormat }),
      sourceSeries: item.series,
    })),
  };
}
