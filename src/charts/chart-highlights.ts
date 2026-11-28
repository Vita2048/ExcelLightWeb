import type { ChartDocument } from '../core/chart';
import type { DocumentId, WorkbookDocument } from '../core/workbook';
import { inferChartSeriesFromSource, resolveFormulaExtent } from './chart-data';

export interface ChartRangeHighlight {
  readonly sheetId: DocumentId;
  readonly area: {
    readonly top: number;
    readonly left: number;
    readonly bottom: number;
    readonly right: number;
  };
  readonly color: string;
  readonly kind: 'categories' | 'name' | 'values';
  readonly seriesIndex: number;
}

const categoryColor = '#a879e8';
const nameColor = '#e06c75';
const valuesColor = '#4ea1ff';

export function chartRangeHighlights(
  document: WorkbookDocument,
  chart: ChartDocument,
): readonly ChartRangeHighlight[] {
  const effectiveSeries = chart.series.length > 0
    ? chart.series
    : inferChartSeriesFromSource(
      document,
      chart.source,
      chart.placement.sheetId,
      chart.chartType,
      chart.seriesBy,
    ).series;

  const highlights: ChartRangeHighlight[] = [];
  const seenKeys = new Set<string>();

  const pushArea = (
    formula: string | undefined,
    color: string,
    kind: ChartRangeHighlight['kind'],
    seriesIndex: number,
  ) => {
    if (!formula) return;
    const area = resolveFormulaExtent(document, chart.placement.sheetId, formula);
    if (!area) return;
    const box = {
      top: area.row,
      left: area.column,
      bottom: area.row + area.rows - 1,
      right: area.column + area.columns - 1,
    };
    const key = `${area.sheetId}:${kind}:${box.top}:${box.left}:${box.bottom}:${box.right}`;
    if (seenKeys.has(key)) return;
    seenKeys.add(key);
    highlights.push({ sheetId: area.sheetId, area: box, color, kind, seriesIndex });
  };

  for (let i = 0; i < effectiveSeries.length; i++) {
    const s = effectiveSeries[i]!;
    pushArea(s.categoriesRef, categoryColor, 'categories', i);
    pushArea(s.xValuesRef, categoryColor, 'categories', i);
    pushArea(s.nameRef, nameColor, 'name', i);
    pushArea(s.valuesRef, valuesColor, 'values', i);
    pushArea(s.bubbleSizesRef, valuesColor, 'values', i);
  }

  return highlights;
}
