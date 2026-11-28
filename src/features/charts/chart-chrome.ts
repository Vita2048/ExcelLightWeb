import {
  chartCategoryOf,
  type ChartAxesConfig,
  type ChartDataLabelPosition,
  type ChartDocument,
  type ChartElements,
  type ChartErrorBarType,
  type ChartGridlinesConfig,
  type ChartLegendPosition,
  type ChartLineDash,
  type ChartSeries,
  type ChartShapeStyle,
  type ChartTrendline,
  type ChartTrendlineType,
  type ChartType,
} from '../../core/chart';
import type { DocumentId } from '../../core/workbook';
import type { ChartUpdatePatch } from '../../application/chart-commands';
import type { ResolvedChartData } from '../../charts/chart-data';

export type ChartSelectedElement =
  | { readonly kind: 'chartArea' }
  | { readonly kind: 'plotArea' }
  | { readonly kind: 'title' }
  | { readonly kind: 'legend' }
  | { readonly kind: 'axis'; readonly axis: 'category' | 'value' | 'secondaryValue' }
  | { readonly kind: 'series'; readonly seriesId: DocumentId }
  | { readonly kind: 'dataLabels'; readonly seriesId?: DocumentId }
  | { readonly kind: 'trendline'; readonly seriesId: DocumentId; readonly trendlineId: DocumentId };

export interface ChartElementChoice {
  readonly key: string;
  readonly label: string;
  readonly element: ChartSelectedElement;
}

export interface SupportedChartElements {
  readonly axes: boolean;
  readonly axisTitles: boolean;
  readonly chartTitle: boolean;
  readonly dataLabels: boolean;
  readonly dataTable: boolean;
  readonly errorBars: boolean;
  readonly gridlines: boolean;
  readonly legend: boolean;
  readonly trendline: boolean;
}

export interface ChartStylePreset {
  readonly id: number;
  readonly name: string;
  readonly chartAreaBg: string;
  readonly chartAreaBorder: string;
  readonly plotAreaBg: string;
  readonly plotAreaBorder?: string;
  readonly showGridlines: boolean;
  readonly showDataLabels?: boolean;
  readonly barOpacity?: number;
  readonly outlineBars?: boolean;
}

export interface QuickLayoutPreset {
  readonly id: number;
  readonly label: string;
  readonly description: string;
}

export const chartStylePresets: readonly ChartStylePreset[] = [
  {
    id: 1,
    name: 'Style 1 — Classic Dark',
    chartAreaBg: '#181b20',
    chartAreaBorder: '#353b45',
    plotAreaBg: 'transparent',
    showGridlines: true,
  },
  {
    id: 2,
    name: 'Style 2 — Subtle Frame',
    chartAreaBg: '#1d2128',
    chartAreaBorder: '#4c5666',
    plotAreaBg: '#15181d',
    plotAreaBorder: '#2d333d',
    showGridlines: true,
  },
  {
    id: 3,
    name: 'Style 3 — Clean Minimal',
    chartAreaBg: '#14171c',
    chartAreaBorder: '#272c35',
    plotAreaBg: 'transparent',
    showGridlines: false,
    showDataLabels: true,
  },
  {
    id: 4,
    name: 'Style 4 — High Contrast Slate',
    chartAreaBg: '#111317',
    chartAreaBorder: '#69b3e7',
    plotAreaBg: '#191d24',
    showGridlines: true,
  },
  {
    id: 5,
    name: 'Style 5 — Midnight Blue Tint',
    chartAreaBg: '#131b26',
    chartAreaBorder: '#2d4461',
    plotAreaBg: '#0f1620',
    showGridlines: true,
  },
  {
    id: 6,
    name: 'Style 6 — Borderless Canvas',
    chartAreaBg: '#181b20',
    chartAreaBorder: 'none',
    plotAreaBg: '#1e2229',
    showGridlines: true,
    showDataLabels: true,
  },
  {
    id: 7,
    name: 'Style 7 — Bold Plot Frame',
    chartAreaBg: '#16191e',
    chartAreaBorder: '#3f4754',
    plotAreaBg: '#1c2028',
    plotAreaBorder: '#505b6b',
    showGridlines: true,
  },
  {
    id: 8,
    name: 'Style 8 — Compact Values',
    chartAreaBg: '#1b1f26',
    chartAreaBorder: '#39414e',
    plotAreaBg: 'transparent',
    showGridlines: false,
    showDataLabels: true,
  },
];

export const quickLayoutPresets: readonly QuickLayoutPreset[] = [
  { id: 1, label: 'Layout 1', description: 'Title and Legend at Right' },
  { id: 2, label: 'Layout 2', description: 'Title, Data Labels, Legend at Top' },
  { id: 3, label: 'Layout 3', description: 'Title and Legend at Bottom' },
  { id: 4, label: 'Layout 4', description: 'Title, Axis Titles, Legend at Bottom' },
  { id: 5, label: 'Layout 5', description: 'Title, Axis Titles, Data Table, Legend' },
  { id: 6, label: 'Layout 6', description: 'Title and Data Labels Only' },
];

export function getChartStylePreset(styleId: number): ChartStylePreset {
  return chartStylePresets.find(p => p.id === styleId) ?? chartStylePresets[0]!;
}

export function supportedElementsForType(type: ChartType): SupportedChartElements {
  const cat = chartCategoryOf(type);
  const noAxes = cat === 'pie' || cat === 'hierarchy';
  const categorical2D = cat === 'column' || cat === 'bar' || cat === 'line' || cat === 'area' || cat === 'scatter' || cat === 'combo';
  return {
    axes: !noAxes,
    axisTitles: !noAxes && cat !== 'radar' && cat !== 'funnel',
    chartTitle: true,
    dataLabels: cat !== 'stock',
    dataTable: cat === 'column' || cat === 'line' || cat === 'area' || cat === 'combo' || cat === 'waterfall',
    errorBars: cat === 'column' || cat === 'bar' || cat === 'line' || cat === 'area' || cat === 'scatter' || cat === 'combo',
    gridlines: !noAxes && cat !== 'radar' && cat !== 'funnel',
    legend: true,
    trendline: categorical2D,
  };
}

export function serializeSelectedElementKey(el: ChartSelectedElement): string {
  switch (el.kind) {
    case 'chartArea': return 'chartArea';
    case 'plotArea': return 'plotArea';
    case 'title': return 'title';
    case 'legend': return 'legend';
    case 'axis': return `axis:${el.axis}`;
    case 'series': return `series:${el.seriesId}`;
    case 'dataLabels': return el.seriesId ? `dataLabels:${el.seriesId}` : 'dataLabels';
    case 'trendline': return `trendline:${el.seriesId}:${el.trendlineId}`;
  }
}

export function elementMatches(a: ChartSelectedElement | undefined, b: ChartSelectedElement): boolean {
  if (!a) return false;
  return serializeSelectedElementKey(a) === serializeSelectedElementKey(b);
}

export function formatElementTitle(el: ChartSelectedElement, chart?: ChartDocument): string {
  switch (el.kind) {
    case 'chartArea': return 'Format Chart Area';
    case 'plotArea': return 'Format Plot Area';
    case 'title': return 'Format Chart Title';
    case 'legend': return 'Format Legend';
    case 'axis': return 'Format Axis';
    case 'series': {
      const idx = chart?.series.findIndex(s => s.id === el.seriesId) ?? -1;
      const name = idx >= 0 ? (chart?.series[idx]?.name || `Series ${idx + 1}`) : 'Data Series';
      return `Format Data Series — ${name}`;
    }
    case 'dataLabels': return 'Format Data Labels';
    case 'trendline': return 'Format Trendline';
  }
}

export function listChartElements(chart: ChartDocument, data: ResolvedChartData): readonly ChartElementChoice[] {
  const supported = supportedElementsForType(chart.chartType);
  const items: ChartElementChoice[] = [
    { key: 'chartArea', label: 'Chart Area', element: { kind: 'chartArea' } },
    { key: 'plotArea', label: 'Plot Area', element: { kind: 'plotArea' } },
  ];
  if (chart.elements.title?.show !== false) {
    items.push({ key: 'title', label: 'Chart Title', element: { kind: 'title' } });
  }
  if (supported.axes) {
    items.push({ key: 'axis:category', label: 'Horizontal (Category) Axis', element: { kind: 'axis', axis: 'category' } });
    items.push({ key: 'axis:value', label: 'Vertical (Value) Axis', element: { kind: 'axis', axis: 'value' } });
    if (chart.series.some(s => s.secondaryAxis) || chart.chartType === 'pareto' || chart.chartType === 'stockVhlc' || chart.chartType === 'stockVohlc') {
      items.push({
        key: 'axis:secondaryValue',
        label: 'Secondary Vertical (Value) Axis',
        element: { kind: 'axis', axis: 'secondaryValue' },
      });
    }
  }
  if (chart.elements.legend?.show !== false) {
    items.push({ key: 'legend', label: 'Legend', element: { kind: 'legend' } });
  }
  if (chart.elements.dataLabels?.show) {
    items.push({ key: 'dataLabels', label: 'Data Labels', element: { kind: 'dataLabels' } });
  }
  data.series.forEach((s, idx) => {
    const label = s.name || `Series ${idx + 1}`;
    items.push({
      key: `series:${s.id}`,
      label: `Series "${label}"`,
      element: { kind: 'series', seriesId: s.id },
    });
    const src = chart.series.find(cs => cs.id === s.id);
    if (src?.dataLabels?.show) {
      items.push({
        key: `dataLabels:${s.id}`,
        label: `Series "${label}" Data Labels`,
        element: { kind: 'dataLabels', seriesId: s.id },
      });
    }
    for (const tl of src?.trendlines ?? []) {
      items.push({
        key: `trendline:${s.id}:${tl.id}`,
        label: `Series "${label}" Trendline (${tl.type})`,
        element: { kind: 'trendline', seriesId: s.id, trendlineId: tl.id },
      });
    }
  });
  return items;
}

export function dashArrayFor(dash?: ChartLineDash, width = 1): string | undefined {
  const s = Math.max(1, width);
  switch (dash) {
    case 'roundDot': return `${s} ${s * 2}`;
    case 'squareDot': return `${s * 1.5} ${s * 1.5}`;
    case 'dash': return `${s * 4} ${s * 2}`;
    case 'dashDot': return `${s * 4} ${s * 2} ${s} ${s * 2}`;
    case 'longDash': return `${s * 8} ${s * 3}`;
    case 'longDashDot': return `${s * 8} ${s * 2} ${s} ${s * 2}`;
    case 'longDashDotDot': return `${s * 8} ${s * 2} ${s} ${s * 2} ${s} ${s * 2}`;
    default: return undefined;
  }
}

export function resetChartToStyle(chart: ChartDocument): ChartUpdatePatch {
  const cleanedSeries: ChartSeries[] = chart.series.map(s => {
    const { fillRgb: _f, borderRgb: _b, lineWidth: _lw, lineDash: _ld, pointOverrides: _po, ...rest } = s;
    return rest;
  });
  return {
    chartAreaStyle: undefined,
    plotAreaStyle: undefined,
    series: cleanedSeries,
  };
}

export function applyQuickLayout(chart: ChartDocument, layoutId: number): Partial<Omit<ChartDocument, 'id'>> {
  const supported = supportedElementsForType(chart.chartType);
  const current = chart.elements;
  const titleText = current.title?.text || 'Chart Title';
  const baseAxes: ChartAxesConfig = current.axes ?? {
    category: { show: true, sourceLinked: true },
    value: { show: true, sourceLinked: true },
  };
  const baseCat = baseAxes.category ?? { show: true, sourceLinked: true };
  const baseVal = baseAxes.value ?? { show: true, sourceLinked: true };

  switch (layoutId) {
    case 1:
      return {
        elements: {
          ...current,
          title: { show: true, text: titleText },
          legend: { show: true, position: 'right' },
          dataLabels: { show: false },
          dataTable: { show: false },
        },
      };
    case 2:
      return {
        elements: {
          ...current,
          title: { show: true, text: titleText },
          legend: { show: true, position: 'top' },
          dataLabels: { show: true, showValue: true, position: 'outsideEnd' },
          dataTable: { show: false },
        },
      };
    case 3:
      return {
        elements: {
          ...current,
          title: { show: true, text: titleText },
          legend: { show: true, position: 'bottom' },
          dataLabels: { show: false },
          dataTable: { show: false },
        },
      };
    case 4:
      return {
        elements: {
          ...current,
          title: { show: true, text: titleText },
          legend: { show: true, position: 'bottom' },
          ...(supported.axisTitles && {
            axes: {
              ...baseAxes,
              category: { ...baseCat, show: true, title: { show: true, text: baseCat.title?.text || 'Category Axis' } },
              value: { ...baseVal, show: true, title: { show: true, text: baseVal.title?.text || 'Value Axis' } },
            },
          }),
          dataLabels: { show: false },
          dataTable: { show: false },
        },
      };
    case 5:
      return {
        elements: {
          ...current,
          title: { show: true, text: titleText },
          legend: { show: true, position: 'bottom' },
          ...(supported.axisTitles && {
            axes: {
              ...baseAxes,
              category: { ...baseCat, show: true, title: { show: true, text: baseCat.title?.text || 'Category Axis' } },
              value: { ...baseVal, show: true, title: { show: true, text: baseVal.title?.text || 'Value Axis' } },
            },
          }),
          dataTable: supported.dataTable ? { show: true, showLegendKeys: true } : { show: false },
          dataLabels: { show: false },
        },
      };
    case 6:
    default:
      return {
        elements: {
          ...current,
          title: { show: true, text: titleText },
          legend: { show: false, position: current.legend?.position ?? 'bottom' },
          dataLabels: { show: true, showValue: true, position: 'outsideEnd' },
          dataTable: { show: false },
        },
      };
  }
}

export function applyChartStylePatch(styleId: number): ChartUpdatePatch {
  const preset = getChartStylePreset(styleId);
  const chartAreaStyle: ChartShapeStyle = {
    fillRgb: preset.chartAreaBg,
    borderRgb: preset.chartAreaBorder,
  };
  const plotAreaStyle: ChartShapeStyle | undefined = preset.plotAreaBg !== 'transparent' || preset.plotAreaBorder
    ? {
      fillRgb: preset.plotAreaBg,
      ...(preset.plotAreaBorder && { borderRgb: preset.plotAreaBorder }),
    }
    : undefined;
  return {
    styleId: preset.id,
    chartAreaStyle,
    plotAreaStyle,
  };
}

export function patchChartElements(
  chart: ChartDocument,
  patch: Partial<ChartElements>,
): Partial<Omit<ChartDocument, 'id'>> {
  return {
    elements: {
      ...chart.elements,
      ...patch,
    },
  };
}

export function toggleChartAxis(
  chart: ChartDocument,
  which: 'category' | 'value',
  show?: boolean,
): Partial<Omit<ChartDocument, 'id'>> {
  const axes = chart.elements.axes ?? {};
  const existing = axes[which] ?? { show: true, sourceLinked: true };
  const nextShow = show ?? !existing.show;
  return patchChartElements(chart, {
    axes: {
      ...axes,
      [which]: { ...existing, show: nextShow },
    },
  });
}

export function toggleChartAxisTitle(
  chart: ChartDocument,
  which: 'category' | 'value',
  show?: boolean,
): Partial<Omit<ChartDocument, 'id'>> {
  const axes = chart.elements.axes ?? {};
  const existing = axes[which] ?? { show: true, sourceLinked: true };
  const nextShow = show ?? !existing.title?.show;
  const fallback = which === 'category' ? 'Axis Title (Horizontal)' : 'Axis Title (Vertical)';
  return patchChartElements(chart, {
    axes: {
      ...axes,
      [which]: {
        ...existing,
        title: {
          ...existing.title,
          show: nextShow,
          text: existing.title?.text || fallback,
        },
      },
    },
  });
}

export function toggleChartGridline(
  chart: ChartDocument,
  key: keyof ChartGridlinesConfig,
  value?: boolean,
): Partial<Omit<ChartDocument, 'id'>> {
  const grid = chart.elements.gridlines ?? {};
  return patchChartElements(chart, {
    gridlines: {
      ...grid,
      [key]: value ?? !grid[key],
    },
  });
}

export function setChartLegend(
  chart: ChartDocument,
  show: boolean,
  position?: ChartLegendPosition,
): Partial<Omit<ChartDocument, 'id'>> {
  return patchChartElements(chart, {
    legend: {
      show,
      position: position ?? chart.elements.legend?.position ?? 'bottom',
    },
  });
}

export function setChartDataLabels(
  chart: ChartDocument,
  show: boolean,
  position?: ChartDataLabelPosition,
): Partial<Omit<ChartDocument, 'id'>> {
  return patchChartElements(chart, {
    dataLabels: {
      ...chart.elements.dataLabels,
      show,
      showValue: chart.elements.dataLabels?.showValue ?? true,
      ...(position !== undefined && { position }),
    },
  });
}

export function togglePrimaryTrendline(
  chart: ChartDocument,
  type?: ChartTrendlineType,
  makeId: () => DocumentId = () => `tl-${Date.now()}`,
  forward?: number,
): Partial<Omit<ChartDocument, 'id'>> {
  if (chart.series.length === 0) return {};
  const hasAny = chart.series.some(s => (s.trendlines?.length ?? 0) > 0);
  if (type === undefined && hasAny) {
    return {
      series: chart.series.map(s => {
        const { trendlines: _tl, ...rest } = s;
        return rest;
      }),
    };
  }
  const nextType = type ?? 'linear';
  const target = chart.series[0]!;
  const newTrendline: ChartTrendline = {
    id: target.trendlines?.[0]?.id ?? makeId(),
    type: nextType,
    ...(nextType === 'movingAverage' && { period: 2 }),
    ...(nextType === 'polynomial' && { polynomialOrder: 2 }),
    ...(forward !== undefined && { forward }),
    displayRSquared: true,
  };
  return {
    series: chart.series.map((s, idx) =>
      idx === 0 ? { ...s, trendlines: [newTrendline] } : s,
    ),
  };
}

export function togglePrimaryErrorBars(
  chart: ChartDocument,
  type?: ChartErrorBarType,
): Partial<Omit<ChartDocument, 'id'>> {
  if (chart.series.length === 0) return {};
  const hasAny = chart.series.some(s => s.errorBars !== undefined);
  if (type === undefined && hasAny) {
    return {
      series: chart.series.map(s => {
        const { errorBars: _eb, ...rest } = s;
        return rest;
      }),
    };
  }
  const barType = type ?? 'percentage';
  const defaultVal = barType === 'percentage' ? 5 : barType === 'fixed' ? 10 : barType === 'stdDev' ? 1 : undefined;
  return {
    series: chart.series.map((s, idx) =>
      idx === 0
        ? {
          ...s,
          errorBars: {
            direction: 'both',
            type: barType,
            ...(defaultVal !== undefined && { value: defaultVal }),
            cap: true,
          },
        }
        : s,
    ),
  };
}
