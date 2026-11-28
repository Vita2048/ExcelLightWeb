import { describe, expect, it } from 'vitest';
import {
  applyChartStylePatch,
  applyQuickLayout,
  listChartElements,
  resetChartToStyle,
  supportedElementsForType,
  toggleChartAxis,
  toggleChartAxisTitle,
  toggleChartGridline,
  togglePrimaryErrorBars,
  togglePrimaryTrendline,
} from '../../src/features/charts/chart-chrome';
import { resolveChartData } from '../../src/charts/chart-data';
import { defaultChartElements, type ChartDocument } from '../../src/core/chart';
import { newWorkbook, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';

function makeGuidFactory() {
  let n = 0;
  return () => `00000000-0000-4000-8000-${String(++n).padStart(12, '0')}`;
}

function makeSampleWorkbook(): {
  doc: WorkbookDocument;
  chart: ChartDocument;
} {
  const makeId = makeGuidFactory();
  const base = newWorkbook(makeId);
  const sheet = base.sheets[0]!;
  const rows: readonly (readonly (string | number)[])[] = [
    ['Month', 'Sales', 'Target'],
    ['Jan', 120, 100],
    ['Feb', 160, 140],
    ['Mar', 210, 180],
  ];
  const cells: CellDocument[] = [];
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r]!.length; c++) {
      const value = rows[r]![c]!;
      cells.push({
        id: makeId(),
        row: r + 1,
        column: c + 1,
        content: typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
      });
    }
  }
  const doc: WorkbookDocument = { ...base, sheets: [{ ...sheet, cells }] };
  const chart: ChartDocument = {
    id: makeId(),
    name: 'Chart 1',
    chartType: 'clusteredColumn',
    placement: { sheetId: sheet.id, row: 1, column: 5, offsetX: 0, offsetY: 0, width: 480, height: 300, zIndex: 1 },
    source: { kind: 'range', formula: '=Sheet1!$A$1:$C$4' },
    seriesBy: 'columns',
    series: [
      { id: makeId(), nameRef: '=Sheet1!$B$1', categoriesRef: '=Sheet1!$A$2:$A$4', valuesRef: '=Sheet1!$B$2:$B$4' },
      { id: makeId(), nameRef: '=Sheet1!$C$1', categoriesRef: '=Sheet1!$A$2:$A$4', valuesRef: '=Sheet1!$C$2:$C$4' },
    ],
    elements: defaultChartElements('clusteredColumn'),
    hiddenEmptyMode: 'gaps',
    showHiddenCells: false,
    styleId: 1,
    paletteId: 'colorful-1',
  };
  return { doc, chart };
}

describe('Sprint 3.4 chart chrome helpers', () => {
  it('filters supported elements accurately by chart type', () => {
    const colSupport = supportedElementsForType('clusteredColumn');
    expect(colSupport.axes).toBe(true);
    expect(colSupport.axisTitles).toBe(true);
    expect(colSupport.dataTable).toBe(true);
    expect(colSupport.trendline).toBe(true);

    const pieSupport = supportedElementsForType('pie');
    expect(pieSupport.axes).toBe(false);
    expect(pieSupport.gridlines).toBe(false);
    expect(pieSupport.trendline).toBe(false);
    expect(pieSupport.dataLabels).toBe(true);
    expect(pieSupport.legend).toBe(true);

    const treemapSupport = supportedElementsForType('treemap');
    expect(treemapSupport.axes).toBe(false);
    expect(treemapSupport.trendline).toBe(false);
  });

  it('lists selectable elements including series and trendlines', () => {
    const { doc, chart } = makeSampleWorkbook();
    const withTlChart: ChartDocument = {
      ...chart,
      ...togglePrimaryTrendline(chart, 'linear', () => 'tl-1'),
    };
    const data = resolveChartData(doc, withTlChart);
    const items = listChartElements(withTlChart, data);
    expect(items.some(i => i.key === 'chartArea')).toBe(true);
    expect(items.some(i => i.key === 'axis:category')).toBe(true);
    expect(items.some(i => i.key === 'axis:value')).toBe(true);
    expect(items.some(i => i.key === `series:${chart.series[0]!.id}`)).toBe(true);
    expect(items.some(i => i.key === `trendline:${chart.series[0]!.id}:tl-1`)).toBe(true);
  });

  it('applies quick layouts, style patches, element toggles, and reset-to-style cleanly', () => {
    const { chart } = makeSampleWorkbook();

    const q5 = applyQuickLayout(chart, 5);
    expect(q5.elements?.dataTable?.show).toBe(true);
    expect(q5.elements?.axes?.category?.title?.show).toBe(true);

    const stylePatch = applyChartStylePatch(4);
    const styled: ChartDocument = {
      ...chart,
      styleId: stylePatch.styleId ?? chart.styleId,
      ...(stylePatch.chartAreaStyle && { chartAreaStyle: stylePatch.chartAreaStyle }),
      ...(stylePatch.plotAreaStyle && { plotAreaStyle: stylePatch.plotAreaStyle }),
      series: chart.series.map((s, idx) => (idx === 0 ? { ...s, fillRgb: '#ff0000' } : s)),
    };
    expect(styled.styleId).toBe(4);
    expect(styled.chartAreaStyle?.borderRgb).toBe('#69b3e7');

    const reset = resetChartToStyle(styled);
    expect(reset.chartAreaStyle).toBeUndefined();
    expect(reset.plotAreaStyle).toBeUndefined();
    expect(reset.series?.[0]?.fillRgb).toBeUndefined();

    const toggledAxis = toggleChartAxis(chart, 'category', false);
    expect(toggledAxis.elements?.axes?.category?.show).toBe(false);
    const toggledTitle = toggleChartAxisTitle(chart, 'value', true);
    expect(toggledTitle.elements?.axes?.value?.title?.show).toBe(true);
    const toggledGrid = toggleChartGridline(chart, 'primaryMajorVertical', true);
    expect(toggledGrid.elements?.gridlines?.primaryMajorVertical).toBe(true);
    const withEb = togglePrimaryErrorBars(chart, 'percentage');
    expect(withEb.series?.[0]?.errorBars?.type).toBe('percentage');
  });
});
