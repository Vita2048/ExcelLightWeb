import type { DocumentId, WorkbookDocument } from './workbook';

export const chartTypes = [
  'clusteredColumn',
  'stackedColumn',
  'stackedColumn100',
  'clusteredBar',
  'stackedBar',
  'stackedBar100',
  'line',
  'stackedLine',
  'stackedLine100',
  'lineMarkers',
  'stackedLineMarkers',
  'stackedLine100Markers',
  'area',
  'stackedArea',
  'stackedArea100',
  'pie',
  'doughnut',
  'pieOfPie',
  'barOfPie',
  'scatter',
  'scatterSmoothLines',
  'scatterSmoothLinesMarkers',
  'scatterStraightLines',
  'scatterStraightLinesMarkers',
  'bubble',
  'combo',
  'waterfall',
  'funnel',
  'histogram',
  'pareto',
  'boxWhisker',
  'treemap',
  'sunburst',
  'stockHlc',
  'stockOhlc',
  'stockVhlc',
  'stockVohlc',
  'radar',
  'radarMarkers',
  'filledRadar',
] as const;

export type ChartType = (typeof chartTypes)[number];

export type ChartCategory =
  | 'column'
  | 'bar'
  | 'line'
  | 'area'
  | 'pie'
  | 'scatter'
  | 'combo'
  | 'waterfall'
  | 'funnel'
  | 'histogram'
  | 'boxWhisker'
  | 'hierarchy'
  | 'stock'
  | 'radar';

export type ChartPlacementMode = 'moveAndSize' | 'moveOnly' | 'freeFloating';

export interface ChartPlacement {
  readonly sheetId: DocumentId;
  readonly row: number;
  readonly column: number;
  readonly offsetX: number;
  readonly offsetY: number;
  readonly width: number;
  readonly height: number;
  readonly mode?: ChartPlacementMode;
  readonly zIndex?: number;
}

export type ChartSource =
  | { readonly kind: 'range'; readonly formula: string }
  | { readonly kind: 'table'; readonly tableId: DocumentId }
  | { readonly kind: 'pivot'; readonly pivotId: DocumentId };

export type ChartHiddenEmptyMode = 'gaps' | 'zero' | 'connect';

export type ChartLineDash =
  | 'solid'
  | 'roundDot'
  | 'squareDot'
  | 'dash'
  | 'dashDot'
  | 'longDash'
  | 'longDashDot'
  | 'longDashDotDot';

export interface ChartShapeStyle {
  readonly fillRgb?: string;
  readonly borderRgb?: string;
  readonly borderWidth?: number;
  readonly borderDash?: ChartLineDash;
}

export interface ChartPointOverride {
  readonly index: number;
  readonly isTotal?: boolean;
  readonly explosion?: number;
  readonly fillRgb?: string;
  readonly borderRgb?: string;
}

export type ChartTrendlineType =
  | 'linear'
  | 'exponential'
  | 'logarithmic'
  | 'polynomial'
  | 'power'
  | 'movingAverage';

export interface ChartTrendline {
  readonly id: DocumentId;
  readonly type: ChartTrendlineType;
  readonly polynomialOrder?: number;
  readonly period?: number;
  readonly forward?: number;
  readonly backward?: number;
  readonly intercept?: number;
  readonly displayEquation?: boolean;
  readonly displayRSquared?: boolean;
  readonly name?: string;
  readonly colorRgb?: string;
}

export type ChartErrorBarDirection = 'both' | 'plus' | 'minus';
export type ChartErrorBarType = 'fixed' | 'percentage' | 'stdDev' | 'stdErr' | 'custom';

export interface ChartErrorBars {
  readonly direction: ChartErrorBarDirection;
  readonly type: ChartErrorBarType;
  readonly value?: number;
  readonly plusRef?: string;
  readonly minusRef?: string;
  readonly cap?: boolean;
}

export type ChartDataLabelPosition =
  | 'center'
  | 'insideEnd'
  | 'insideBase'
  | 'outsideEnd'
  | 'bestFit'
  | 'callout'
  | 'left'
  | 'right'
  | 'above'
  | 'below';

export interface ChartDataLabelsConfig {
  readonly show: boolean;
  readonly position?: ChartDataLabelPosition;
  readonly showValue?: boolean;
  readonly showCategoryName?: boolean;
  readonly showSeriesName?: boolean;
  readonly showPercentage?: boolean;
  readonly separator?: string;
  readonly numberFormat?: string;
}

export interface ChartTitleConfig {
  readonly show: boolean;
  readonly text?: string;
  readonly formulaRef?: string;
  readonly overlay?: boolean;
}

export type ChartLegendPosition = 'top' | 'bottom' | 'left' | 'right' | 'overlayRight';

export interface ChartLegendConfig {
  readonly show: boolean;
  readonly position: ChartLegendPosition;
}

export type ChartAxisDisplayUnits = 'none' | 'hundreds' | 'thousands' | 'millions' | 'billions';
export type ChartTickMark = 'none' | 'inside' | 'outside' | 'cross';
export type ChartAxisLabelPosition = 'nextToAxis' | 'high' | 'low' | 'none';

export interface ChartAxisConfig {
  readonly show: boolean;
  readonly title?: ChartTitleConfig;
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  readonly minorUnit?: number;
  readonly logarithmic?: boolean;
  readonly logBase?: number;
  readonly reverseOrder?: boolean;
  readonly crossesAt?: number | 'min' | 'max' | 'auto';
  readonly displayUnits?: ChartAxisDisplayUnits;
  readonly majorTick?: ChartTickMark;
  readonly minorTick?: ChartTickMark;
  readonly labelPosition?: ChartAxisLabelPosition;
  readonly numberFormat?: string;
  readonly sourceLinked?: boolean;
}

export interface ChartAxesConfig {
  readonly category?: ChartAxisConfig;
  readonly value?: ChartAxisConfig;
  readonly secondaryCategory?: ChartAxisConfig;
  readonly secondaryValue?: ChartAxisConfig;
}

export interface ChartGridlinesConfig {
  readonly primaryMajorHorizontal?: boolean;
  readonly primaryMinorHorizontal?: boolean;
  readonly primaryMajorVertical?: boolean;
  readonly primaryMinorVertical?: boolean;
}

export interface ChartDataTableConfig {
  readonly show: boolean;
  readonly showLegendKeys?: boolean;
}

export interface ChartElements {
  readonly title?: ChartTitleConfig;
  readonly axes?: ChartAxesConfig;
  readonly gridlines?: ChartGridlinesConfig;
  readonly legend?: ChartLegendConfig;
  readonly dataLabels?: ChartDataLabelsConfig;
  readonly dataTable?: ChartDataTableConfig;
}

export type ChartBinMode = 'auto' | 'binWidth' | 'binCount' | 'byCategory';
export type ChartQuartileMethod = 'exclusive' | 'inclusive';

export interface ChartSeries {
  readonly id: DocumentId;
  readonly name?: string;
  readonly nameRef?: string;
  readonly categoriesRef?: string;
  readonly valuesRef: string;
  readonly xValuesRef?: string;
  readonly bubbleSizesRef?: string;
  readonly chartType?: ChartType;
  readonly secondaryAxis?: boolean;
  readonly hidden?: boolean;
  readonly fillRgb?: string;
  readonly borderRgb?: string;
  readonly lineWidth?: number;
  readonly lineDash?: ChartLineDash;
  readonly smooth?: boolean;
  readonly markers?: boolean;
  readonly gapWidth?: number;
  readonly overlap?: number;
  readonly explosion?: number;
  readonly holeSize?: number;
  readonly firstSliceAngle?: number;
  readonly binMode?: ChartBinMode;
  readonly binWidth?: number;
  readonly binCount?: number;
  readonly overflowBin?: number;
  readonly underflowBin?: number;
  readonly quartileMethod?: ChartQuartileMethod;
  readonly showInnerPoints?: boolean;
  readonly showOutlierPoints?: boolean;
  readonly showMeanMarkers?: boolean;
  readonly showMeanLine?: boolean;
  readonly showConnectorLines?: boolean;
  readonly pointOverrides?: readonly ChartPointOverride[];
  readonly dataLabels?: ChartDataLabelsConfig;
  readonly trendlines?: readonly ChartTrendline[];
  readonly errorBars?: ChartErrorBars;
}

export interface ChartDocument {
  readonly id: DocumentId;
  readonly name: string;
  readonly chartType: ChartType;
  readonly placement: ChartPlacement;
  readonly source: ChartSource;
  readonly seriesBy: 'columns' | 'rows';
  readonly series: readonly ChartSeries[];
  readonly elements: ChartElements;
  readonly hiddenEmptyMode: ChartHiddenEmptyMode;
  readonly showHiddenCells: boolean;
  readonly hiddenCategories?: readonly number[];
  readonly styleId: number;
  readonly paletteId: string;
  readonly chartAreaStyle?: ChartShapeStyle;
  readonly plotAreaStyle?: ChartShapeStyle;
}

const chartTypeSet: ReadonlySet<string> = new Set(chartTypes);

export function isValidChartType(value: unknown): value is ChartType {
  return typeof value === 'string' && chartTypeSet.has(value);
}

export function chartCategoryOf(type: ChartType): ChartCategory {
  if (type === 'clusteredColumn' || type === 'stackedColumn' || type === 'stackedColumn100') return 'column';
  if (type === 'clusteredBar' || type === 'stackedBar' || type === 'stackedBar100') return 'bar';
  if (
    type === 'line' || type === 'stackedLine' || type === 'stackedLine100' ||
    type === 'lineMarkers' || type === 'stackedLineMarkers' || type === 'stackedLine100Markers'
  ) return 'line';
  if (type === 'area' || type === 'stackedArea' || type === 'stackedArea100') return 'area';
  if (type === 'pie' || type === 'doughnut' || type === 'pieOfPie' || type === 'barOfPie') return 'pie';
  if (
    type === 'scatter' || type === 'scatterSmoothLines' || type === 'scatterSmoothLinesMarkers' ||
    type === 'scatterStraightLines' || type === 'scatterStraightLinesMarkers' || type === 'bubble'
  ) return 'scatter';
  if (type === 'combo') return 'combo';
  if (type === 'waterfall') return 'waterfall';
  if (type === 'funnel') return 'funnel';
  if (type === 'histogram' || type === 'pareto') return 'histogram';
  if (type === 'boxWhisker') return 'boxWhisker';
  if (type === 'treemap' || type === 'sunburst') return 'hierarchy';
  if (type === 'stockHlc' || type === 'stockOhlc' || type === 'stockVhlc' || type === 'stockVohlc') return 'stock';
  return 'radar';
}

export function nextChartName(document: WorkbookDocument): string {
  const used = new Set((document.charts ?? []).map(chart => chart.name.toUpperCase()));
  for (let index = 1; ; index++) {
    const candidate = `Chart ${index}`;
    if (!used.has(candidate.toUpperCase())) return candidate;
  }
}

export function defaultChartElements(chartType: ChartType): ChartElements {
  const category = chartCategoryOf(chartType);
  const hasAxes = category !== 'pie' && category !== 'hierarchy';
  const hasGridlines = hasAxes && category !== 'radar' && category !== 'funnel';
  return {
    title: { show: true, text: 'Chart Title' },
    legend: { show: category !== 'histogram' && category !== 'funnel', position: 'bottom' },
    ...(hasAxes ? {
      axes: {
        category: { show: true, sourceLinked: true },
        value: { show: true, sourceLinked: true },
      },
    } : {}),
    ...(hasGridlines ? {
      gridlines: {
        primaryMajorHorizontal: category !== 'bar',
        primaryMajorVertical: category === 'bar' || category === 'scatter',
      },
    } : {}),
  };
}

function rewriteOptionalRef(ref: string | undefined, rewrite: (source: string) => string): string | undefined {
  if (ref === undefined || !ref.startsWith('=')) return ref;
  const next = rewrite(ref);
  return next === ref ? ref : next;
}

function rewriteTitleConfig(
  title: ChartTitleConfig | undefined,
  rewrite: (source: string) => string,
): ChartTitleConfig | undefined {
  if (!title || title.formulaRef === undefined) return title;
  const formulaRef = rewriteOptionalRef(title.formulaRef, rewrite);
  if (formulaRef === title.formulaRef) return title;
  const { formulaRef: _prev, ...rest } = title;
  return formulaRef !== undefined ? { ...rest, formulaRef } : rest;
}

function rewriteAxisConfig(
  axis: ChartAxisConfig | undefined,
  rewrite: (source: string) => string,
): ChartAxisConfig | undefined {
  if (!axis || !axis.title) return axis;
  const title = rewriteTitleConfig(axis.title, rewrite);
  if (title === axis.title) return axis;
  const { title: _prev, ...rest } = axis;
  return title !== undefined ? { ...rest, title } : rest;
}

export function rewriteChartFormulaRefs(
  chart: ChartDocument,
  rewrite: (source: string) => string,
): ChartDocument {
  let source = chart.source;
  if (source.kind === 'range') {
    const formula = rewriteOptionalRef(source.formula, rewrite) ?? source.formula;
    if (formula !== source.formula) source = { kind: 'range', formula };
  }

  let seriesChanged = false;
  const series = chart.series.map(entry => {
    const nameRef = rewriteOptionalRef(entry.nameRef, rewrite);
    const categoriesRef = rewriteOptionalRef(entry.categoriesRef, rewrite);
    const valuesRef = rewriteOptionalRef(entry.valuesRef, rewrite) ?? entry.valuesRef;
    const xValuesRef = rewriteOptionalRef(entry.xValuesRef, rewrite);
    const bubbleSizesRef = rewriteOptionalRef(entry.bubbleSizesRef, rewrite);
    let errorBars = entry.errorBars;
    if (errorBars) {
      const plusRef = rewriteOptionalRef(errorBars.plusRef, rewrite);
      const minusRef = rewriteOptionalRef(errorBars.minusRef, rewrite);
      if (plusRef !== errorBars.plusRef || minusRef !== errorBars.minusRef) {
        errorBars = {
          ...errorBars,
          ...(plusRef !== undefined && { plusRef }),
          ...(minusRef !== undefined && { minusRef }),
        };
      }
    }
    if (
      nameRef === entry.nameRef &&
      categoriesRef === entry.categoriesRef &&
      valuesRef === entry.valuesRef &&
      xValuesRef === entry.xValuesRef &&
      bubbleSizesRef === entry.bubbleSizesRef &&
      errorBars === entry.errorBars
    ) {
      return entry;
    }
    seriesChanged = true;
    return {
      ...entry,
      ...(nameRef !== undefined && { nameRef }),
      ...(categoriesRef !== undefined && { categoriesRef }),
      valuesRef,
      ...(xValuesRef !== undefined && { xValuesRef }),
      ...(bubbleSizesRef !== undefined && { bubbleSizesRef }),
      ...(errorBars !== undefined && { errorBars }),
    };
  });

  let elements = chart.elements;
  const nextTitle = rewriteTitleConfig(elements.title, rewrite);
  let nextAxes = elements.axes;
  if (nextAxes) {
    const category = rewriteAxisConfig(nextAxes.category, rewrite);
    const value = rewriteAxisConfig(nextAxes.value, rewrite);
    const secondaryCategory = rewriteAxisConfig(nextAxes.secondaryCategory, rewrite);
    const secondaryValue = rewriteAxisConfig(nextAxes.secondaryValue, rewrite);
    if (
      category !== nextAxes.category ||
      value !== nextAxes.value ||
      secondaryCategory !== nextAxes.secondaryCategory ||
      secondaryValue !== nextAxes.secondaryValue
    ) {
      nextAxes = {
        ...(category !== undefined && { category }),
        ...(value !== undefined && { value }),
        ...(secondaryCategory !== undefined && { secondaryCategory }),
        ...(secondaryValue !== undefined && { secondaryValue }),
      };
    }
  }
  if (nextTitle !== elements.title || nextAxes !== elements.axes) {
    elements = {
      ...elements,
      ...(nextTitle !== undefined && { title: nextTitle }),
      ...(nextAxes !== undefined && { axes: nextAxes }),
    };
  }

  if (source === chart.source && !seriesChanged && elements === chart.elements) {
    return chart;
  }
  return {
    ...chart,
    source,
    series: seriesChanged ? series : chart.series,
    elements,
  };
}
