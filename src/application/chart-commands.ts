import { inferChartSeriesFromSource } from '../charts/chart-data';
import {
  defaultChartElements,
  nextChartName,
  type ChartDocument,
  type ChartElements,
  type ChartHiddenEmptyMode,
  type ChartPlacement,
  type ChartPlacementMode,
  type ChartSeries,
  type ChartShapeStyle,
  type ChartSource,
  type ChartType,
} from '../core/chart';
import type { Coordinate, DocumentId, IdFactory, WorkbookDocument } from '../core/workbook';
import { ordinalNameKey, requireDocument } from '../core/workbook-validation';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { StructureCommand } from './workbook-session';

export interface CreateChartInput {
  readonly chartType?: ChartType;
  readonly name?: string;
  readonly source: ChartSource;
  readonly placement: {
    readonly sheetId: DocumentId;
    readonly row: number;
    readonly column: number;
    readonly offsetX?: number;
    readonly offsetY?: number;
    readonly width?: number;
    readonly height?: number;
    readonly mode?: ChartPlacementMode;
    readonly zIndex?: number;
  };
  readonly seriesBy?: 'columns' | 'rows';
  readonly series?: readonly ChartSeries[];
  readonly elements?: ChartElements;
  readonly hiddenEmptyMode?: ChartHiddenEmptyMode;
  readonly showHiddenCells?: boolean;
  readonly showFieldButtons?: boolean;
  readonly styleId?: number;
  readonly paletteId?: string;
  readonly chartAreaStyle?: ChartShapeStyle;
  readonly plotAreaStyle?: ChartShapeStyle;
}

export type ChartReorderAction = 'bringForward' | 'bringToFront' | 'sendBackward' | 'sendToBack';
export type ChartUpdatePatch = Partial<Omit<ChartDocument, 'id' | 'hiddenCategories' | 'chartAreaStyle' | 'plotAreaStyle'>> & {
  readonly hiddenCategories?: readonly number[] | undefined;
  readonly chartAreaStyle?: ChartShapeStyle | undefined;
  readonly plotAreaStyle?: ChartShapeStyle | undefined;
};

function replaceCharts(document: WorkbookDocument, charts: readonly ChartDocument[]): WorkbookDocument {
  const { charts: _prev, ...rest } = document;
  return charts.length > 0 ? { ...rest, charts } : rest;
}

function nextSheetZIndex(charts: readonly ChartDocument[], sheetId: DocumentId): number {
  let maxZ = 0;
  for (const chart of charts) {
    if (chart.placement.sheetId === sheetId) {
      maxZ = Math.max(maxZ, chart.placement.zIndex ?? 0);
    }
  }
  return maxZ + 1;
}

export function createChart(
  basis: WorkbookSnapshot,
  input: CreateChartInput,
  makeId: IdFactory,
): StructureCommand {
  const { document } = basis;
  const chartType: ChartType = input.chartType ?? 'clusteredColumn';
  const name = (input.name ?? nextChartName(document)).trim();
  requireDocument(name.length > 0 && name.length <= 255, 'Chart requires a valid name.');
  const nameKey = ordinalNameKey(name);
  const existing = document.charts ?? [];
  requireDocument(!existing.some(item => ordinalNameKey(item.name) === nameKey), 'A chart with this name already exists.');
  requireDocument(document.sheets.some(sheet => sheet.id === input.placement.sheetId), 'Chart target worksheet no longer exists.');

  if (input.source.kind === 'table') {
    const tableId = input.source.tableId;
    requireDocument(
      document.names.some(n => n.id === tableId && n.isTable && !n.area.invalid),
      'Chart source table no longer exists.',
    );
  } else if (input.source.kind === 'pivot') {
    const pivotId = input.source.pivotId;
    requireDocument(
      (document.pivots ?? []).some(p => p.id === pivotId),
      'Chart source PivotTable no longer exists.',
    );
  }

  let seriesBy = input.seriesBy ?? 'columns';
  let series: readonly ChartSeries[];
  if (input.series !== undefined) {
    series = input.series;
  } else if (input.source.kind === 'range') {
    const inferred = inferChartSeriesFromSource(
      document,
      input.source,
      input.placement.sheetId,
      chartType,
      input.seriesBy,
    );
    seriesBy = inferred.seriesBy;
    series = inferred.series.map(item => ({ id: makeId(), ...item }));
  } else {
    series = [];
  }

  const zIndex = input.placement.zIndex ?? nextSheetZIndex(existing, input.placement.sheetId);
  const chart: ChartDocument = {
    id: makeId(),
    name,
    chartType,
    placement: {
      sheetId: input.placement.sheetId,
      row: input.placement.row,
      column: input.placement.column,
      offsetX: input.placement.offsetX ?? 0,
      offsetY: input.placement.offsetY ?? 0,
      width: input.placement.width ?? 480,
      height: input.placement.height ?? 300,
      ...(input.placement.mode !== undefined && { mode: input.placement.mode }),
      zIndex,
    },
    source: input.source,
    seriesBy,
    series,
    elements: input.elements ?? defaultChartElements(chartType),
    hiddenEmptyMode: input.hiddenEmptyMode ?? 'gaps',
    showHiddenCells: input.showHiddenCells ?? false,
    ...(input.showFieldButtons !== undefined && { showFieldButtons: input.showFieldButtons }),
    styleId: input.styleId ?? 1,
    paletteId: input.paletteId ?? 'colorful-1',
    ...(input.chartAreaStyle && { chartAreaStyle: input.chartAreaStyle }),
    ...(input.plotAreaStyle && { plotAreaStyle: input.plotAreaStyle }),
  };

  return {
    kind: 'structure',
    label: 'Insert Chart',
    document: replaceCharts(document, [...existing, chart]),
  };
}

export function updateChart(
  basis: WorkbookSnapshot,
  chartId: DocumentId,
  patch: ChartUpdatePatch,
): StructureCommand {
  const existing = basis.document.charts ?? [];
  const current = existing.find(item => item.id === chartId);
  requireDocument(current, 'Chart no longer exists.');
  let nextName = current.name;
  if (patch.name !== undefined) {
    nextName = patch.name.trim();
    requireDocument(nextName.length > 0 && nextName.length <= 255, 'Chart requires a valid name.');
    const key = ordinalNameKey(nextName);
    requireDocument(
      !existing.some(item => item.id !== chartId && ordinalNameKey(item.name) === key),
      'A chart with this name already exists.',
    );
  }
  const merged = { ...current, ...patch, id: current.id, name: nextName } as Record<string, unknown>;
  for (const key of Object.keys(merged)) {
    if (merged[key] === undefined) delete merged[key];
  }
  const updated = merged as unknown as ChartDocument;
  return {
    kind: 'structure',
    label: 'Update Chart',
    document: replaceCharts(basis.document, existing.map(item => (item.id === chartId ? updated : item))),
  };
}

export function moveResizeChart(
  basis: WorkbookSnapshot,
  chartId: DocumentId,
  placementPatch: Partial<ChartPlacement>,
): StructureCommand {
  const existing = basis.document.charts ?? [];
  const current = existing.find(item => item.id === chartId);
  requireDocument(current, 'Chart no longer exists.');
  const updated: ChartDocument = {
    ...current,
    placement: { ...current.placement, ...placementPatch },
  };
  return {
    kind: 'structure',
    label: 'Move/Resize Chart',
    document: replaceCharts(basis.document, existing.map(item => (item.id === chartId ? updated : item))),
  };
}

export function reorderChart(
  basis: WorkbookSnapshot,
  chartId: DocumentId,
  action: ChartReorderAction,
): StructureCommand {
  const existing = basis.document.charts ?? [];
  const current = existing.find(item => item.id === chartId);
  requireDocument(current, 'Chart no longer exists.');
  const sheetId = current.placement.sheetId;

  const sameSheet = existing
    .map((chart, index) => ({ chart, index }))
    .filter(entry => entry.chart.placement.sheetId === sheetId)
    .sort((a, b) => (a.chart.placement.zIndex ?? a.index) - (b.chart.placement.zIndex ?? b.index) || a.index - b.index)
    .map(entry => entry.chart.id);

  const pos = sameSheet.indexOf(chartId);
  if (pos >= 0 && sameSheet.length > 1) {
    sameSheet.splice(pos, 1);
    if (action === 'bringToFront') sameSheet.push(chartId);
    else if (action === 'sendToBack') sameSheet.unshift(chartId);
    else if (action === 'bringForward') sameSheet.splice(Math.min(sameSheet.length, pos + 1), 0, chartId);
    else sameSheet.splice(Math.max(0, pos - 1), 0, chartId);
  }

  const zMap = new Map<DocumentId, number>();
  sameSheet.forEach((id, idx) => zMap.set(id, idx + 1));

  const nextCharts = existing.map(chart => {
    const nextZ = zMap.get(chart.id);
    if (nextZ === undefined || nextZ === chart.placement.zIndex) return chart;
    return { ...chart, placement: { ...chart.placement, zIndex: nextZ } };
  });

  return {
    kind: 'structure',
    label: 'Reorder Chart',
    document: replaceCharts(basis.document, nextCharts),
  };
}

export function moveChartToSheet(
  basis: WorkbookSnapshot,
  chartId: DocumentId,
  targetSheetId: DocumentId,
  anchor?: Partial<Coordinate> & { readonly offsetX?: number; readonly offsetY?: number },
): StructureCommand {
  const existing = basis.document.charts ?? [];
  const current = existing.find(item => item.id === chartId);
  requireDocument(current, 'Chart no longer exists.');
  requireDocument(basis.document.sheets.some(sheet => sheet.id === targetSheetId), 'Destination worksheet no longer exists.');
  const zIndex = nextSheetZIndex(existing.filter(item => item.id !== chartId), targetSheetId);
  const updated: ChartDocument = {
    ...current,
    placement: {
      ...current.placement,
      sheetId: targetSheetId,
      row: anchor?.row ?? (current.placement.sheetId === targetSheetId ? current.placement.row : 2),
      column: anchor?.column ?? (current.placement.sheetId === targetSheetId ? current.placement.column : 2),
      offsetX: anchor?.offsetX ?? (current.placement.sheetId === targetSheetId ? current.placement.offsetX : 0),
      offsetY: anchor?.offsetY ?? (current.placement.sheetId === targetSheetId ? current.placement.offsetY : 0),
      zIndex,
    },
  };
  return {
    kind: 'structure',
    label: 'Move Chart',
    document: replaceCharts(basis.document, existing.map(item => (item.id === chartId ? updated : item))),
  };
}

export function deleteChart(basis: WorkbookSnapshot, chartId: DocumentId): StructureCommand {
  const existing = basis.document.charts ?? [];
  requireDocument(existing.some(item => item.id === chartId), 'Chart no longer exists.');
  return {
    kind: 'structure',
    label: 'Delete Chart',
    document: replaceCharts(basis.document, existing.filter(item => item.id !== chartId)),
  };
}
