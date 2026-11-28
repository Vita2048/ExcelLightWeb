import { defaultPivotValueLabel, type PivotTableDocument } from '../core/pivot';
import type { ChartDocument } from '../core/chart';
import type { WorkbookDocument } from '../core/workbook';

export type PivotChartFieldButtonKind = 'reportFilter' | 'rowLabel' | 'colLabel' | 'value';

export interface PivotChartFieldButtonDescriptor {
  readonly kind: PivotChartFieldButtonKind;
  readonly fieldName: string;
  readonly label: string;
  readonly filtered: boolean;
  readonly sorted?: 'asc' | 'desc';
}

export interface PivotChartFieldSummary {
  readonly pivot: PivotTableDocument;
  readonly reportFilters: readonly PivotChartFieldButtonDescriptor[];
  readonly valueFields: readonly PivotChartFieldButtonDescriptor[];
  readonly axisFields: readonly PivotChartFieldButtonDescriptor[];
  readonly legendFields: readonly PivotChartFieldButtonDescriptor[];
}

export function isPivotChart(chart: ChartDocument): boolean {
  return chart.source.kind === 'pivot';
}

export function describePivotChartFields(
  document: WorkbookDocument,
  chart: ChartDocument,
): PivotChartFieldSummary | undefined {
  if (chart.source.kind !== 'pivot') return undefined;
  const pivotId = chart.source.pivotId;
  const pivot = (document.pivots ?? []).find(item => item.id === pivotId);
  if (!pivot) return undefined;

  const reportFilters: PivotChartFieldButtonDescriptor[] = pivot.filters.map(f => ({
    kind: 'reportFilter',
    fieldName: f.fieldName,
    label: f.fieldName,
    filtered: (f.hiddenItems?.length ?? 0) > 0,
  }));

  const valueFields: PivotChartFieldButtonDescriptor[] = pivot.values.map(v => ({
    kind: 'value',
    fieldName: v.fieldName,
    label: v.customName?.trim() || defaultPivotValueLabel(v.fieldName, v.aggregation),
    filtered: false,
  }));

  const axisFields: PivotChartFieldButtonDescriptor[] = pivot.rows.map(r => ({
    kind: 'rowLabel',
    fieldName: r.fieldName,
    label: r.fieldName,
    filtered: (r.hiddenItems?.length ?? 0) > 0,
    ...(r.sort ? { sorted: r.sort.direction } : {}),
  }));

  const legendFields: PivotChartFieldButtonDescriptor[] = pivot.columns.map(c => ({
    kind: 'colLabel',
    fieldName: c.fieldName,
    label: c.fieldName,
    filtered: (c.hiddenItems?.length ?? 0) > 0,
    ...(c.sort ? { sorted: c.sort.direction } : {}),
  }));

  return {
    pivot,
    reportFilters,
    valueFields,
    axisFields,
    legendFields,
  };
}
