import { describe, expect, it } from 'vitest';
import { resolveChartData } from '../../src/charts/chart-data';
import { chartRangeHighlights } from '../../src/charts/chart-highlights';
import { computeCore2DChartLayout } from '../../src/charts/layout/core-2d-layout';
import { computeNiceAxisScale, formatAxisTick } from '../../src/charts/layout/scales';
import { defaultChartElements, type ChartDocument, type ChartType } from '../../src/core/chart';
import { newWorkbook, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';

function makeGuidFactory() {
  let n = 0;
  return () => `00000000-0000-4000-8000-${String(++n).padStart(12, '0')}`;
}

function buildWorkbook(rows: readonly (readonly (string | number | null)[])[]): {
  doc: WorkbookDocument;
  makeId: () => string;
  sheetId: string;
} {
  const makeId = makeGuidFactory();
  const base = newWorkbook(makeId);
  const sheet = base.sheets[0]!;
  const cells: CellDocument[] = [];
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r]!.length; c++) {
      const value = rows[r]![c];
      if (value === null || value === undefined) continue;
      cells.push({
        id: makeId(),
        row: r + 1,
        column: c + 1,
        content: typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
      });
    }
  }
  return { doc: { ...base, sheets: [{ ...sheet, cells }] }, makeId, sheetId: sheet.id };
}

function makeChart(
  makeId: () => string,
  sheetId: string,
  chartType: ChartType,
  overrides: Partial<ChartDocument> = {},
): ChartDocument {
  return {
    id: makeId(),
    name: 'Chart 1',
    chartType,
    placement: { sheetId, row: 1, column: 6, offsetX: 0, offsetY: 0, width: 480, height: 300, zIndex: 1 },
    source: { kind: 'range', formula: '=Sheet1!$A$1:$C$4' },
    seriesBy: 'columns',
    series: [
      { id: makeId(), nameRef: '=Sheet1!$B$1', categoriesRef: '=Sheet1!$A$2:$A$4', valuesRef: '=Sheet1!$B$2:$B$4' },
      { id: makeId(), nameRef: '=Sheet1!$C$1', categoriesRef: '=Sheet1!$A$2:$A$4', valuesRef: '=Sheet1!$C$2:$C$4' },
    ],
    elements: defaultChartElements(chartType),
    hiddenEmptyMode: 'gaps',
    showHiddenCells: false,
    styleId: 1,
    paletteId: 'colorful-1',
    ...overrides,
  };
}

describe('2D chart layout engine and range highlights', () => {
  it('computes nice axis scales for positive, negative, zero, and 100% stacked bounds', () => {
    const zeroIncluded = computeNiceAxisScale(12, 87, { includeZero: true });
    expect(zeroIncluded.min).toBe(0);
    expect(zeroIncluded.max).toBeGreaterThanOrEqual(87);
    expect(zeroIncluded.ticks.length).toBeGreaterThanOrEqual(2);

    const negScale = computeNiceAxisScale(-42, 18, { includeZero: true });
    expect(negScale.min).toBeLessThanOrEqual(-42);
    expect(negScale.max).toBeGreaterThanOrEqual(18);

    const allZero = computeNiceAxisScale(0, 0, { includeZero: true });
    expect(allZero.min).toBe(0);
    expect(allZero.max).toBe(1);

    const pct = computeNiceAxisScale(0, 1, { isPercent100: true });
    expect(pct.min).toBe(0);
    expect(pct.max).toBe(1);
    expect(formatAxisTick(0.4, undefined, true)).toBe('40%');
  });

  it('lays out clustered, stacked, and 100% column & bar charts cleanly', () => {
    const { doc, makeId, sheetId } = buildWorkbook([
      ['Qtr', 'East', 'West'],
      ['Q1', 100, 100],
      ['Q2', 150, 50],
      ['Q3', 200, 200],
    ]);

    const clustered = makeChart(makeId, sheetId, 'clusteredColumn');
    const cLayout = computeCore2DChartLayout(resolveChartData(doc, clustered), clustered, 480, 300);
    expect(cLayout.bars).toHaveLength(6);
    expect(cLayout.isEmpty).toBe(false);
    expect(cLayout.legend).toHaveLength(2);

    const stacked100Bar = makeChart(makeId, sheetId, 'stackedBar100');
    const bLayout = computeCore2DChartLayout(resolveChartData(doc, stacked100Bar), stacked100Bar, 480, 300);
    expect(bLayout.bars).toHaveLength(6);
    // Q1 East (100) + Q1 West (100) = 50% + 50%
    const q1Bars = bLayout.bars.filter(b => b.categoryIndex === 0);
    expect(q1Bars[0]!.width).toBeCloseTo(bLayout.plotArea.width * 0.5, 1);
    expect(q1Bars[1]!.width).toBeCloseTo(bLayout.plotArea.width * 0.5, 1);
  });

  it('handles line gap breaking vs connectGaps, area paths, pie/pieOfPie slices, and scatter/bubble', () => {
    const { doc, makeId, sheetId } = buildWorkbook([
      ['X', 'Y', 'Z'],
      ['A', 10, 5],
      ['B', null, 15],
      ['C', 30, 25],
      ['D', 40, 35],
    ]);

    const gapLine = makeChart(makeId, sheetId, 'lineMarkers', {
      source: { kind: 'range', formula: '=Sheet1!$A$1:$C$5' },
      series: [{ id: makeId(), nameRef: '=Sheet1!$B$1', categoriesRef: '=Sheet1!$A$2:$A$5', valuesRef: '=Sheet1!$B$2:$B$5' }],
      hiddenEmptyMode: 'gaps',
    });
    const gapLayout = computeCore2DChartLayout(resolveChartData(doc, gapLine), gapLine, 480, 300);
    // Two separate M ... segments around the null gap at index 1
    expect(gapLayout.lines[0]!.path.match(/M /g)?.length).toBe(2);
    expect(gapLayout.lines[0]!.markers).toHaveLength(3);

    const connectedLine = { ...gapLine, hiddenEmptyMode: 'connect' as const };
    const connLayout = computeCore2DChartLayout(resolveChartData(doc, connectedLine), connectedLine, 480, 300);
    expect(connLayout.lines[0]!.path.match(/M /g)?.length).toBe(1);

    const pieOfPie = makeChart(makeId, sheetId, 'pieOfPie', {
      source: { kind: 'range', formula: '=Sheet1!$A$1:$C$5' },
      series: [{ id: makeId(), nameRef: '=Sheet1!$C$1', categoriesRef: '=Sheet1!$A$2:$A$5', valuesRef: '=Sheet1!$C$2:$C$5' }],
    });
    const pieLayout = computeCore2DChartLayout(resolveChartData(doc, pieOfPie), pieOfPie, 480, 300);
    expect(pieLayout.slices.length).toBeGreaterThanOrEqual(4);
    expect(pieLayout.calloutConnectors).toHaveLength(2);

    const bubbleDoc = buildWorkbook([
      ['X', 'Y', 'Size'],
      [1, 10, 4],
      [2, 20, 16],
      [3, 15, 9],
    ]);
    const bubbleChart = makeChart(bubbleDoc.makeId, bubbleDoc.sheetId, 'bubble', {
      source: { kind: 'range', formula: '=Sheet1!$A$1:$C$4' },
      series: [{
        id: bubbleDoc.makeId(),
        nameRef: '=Sheet1!$B$1',
        xValuesRef: '=Sheet1!$A$2:$A$4',
        valuesRef: '=Sheet1!$B$2:$B$4',
        bubbleSizesRef: '=Sheet1!$C$2:$C$4',
      }],
    });
    const bubbleLayout = computeCore2DChartLayout(resolveChartData(bubbleDoc.doc, bubbleChart), bubbleChart, 480, 300);
    expect(bubbleLayout.bubbles).toHaveLength(3);
    expect(bubbleLayout.bubbles[1]!.r).toBeGreaterThan(bubbleLayout.bubbles[0]!.r);
  });

  it('produces color-coded range highlights for categories, series name, and values', () => {
    const { doc, makeId, sheetId } = buildWorkbook([
      ['Month', 'Sales'],
      ['Jan', 100],
      ['Feb', 200],
    ]);
    const chart = makeChart(makeId, sheetId, 'clusteredColumn', {
      series: [{ id: makeId(), nameRef: '=Sheet1!$B$1', categoriesRef: '=Sheet1!$A$2:$A$3', valuesRef: '=Sheet1!$B$2:$B$3' }],
    });
    const highlights = chartRangeHighlights(doc, chart);
    expect(highlights.map(h => h.kind)).toEqual(['categories', 'name', 'values']);
    expect(highlights.find(h => h.kind === 'values')?.area).toEqual({ top: 2, left: 2, bottom: 3, right: 2 });
  });

  it('sizes left/right legend padding tightly and reserves non-overlapping space for Y-axis title', () => {
    const { doc, makeId, sheetId } = buildWorkbook([
      ['Region', 'Revenue'],
      ['North', 40000],
      ['South', 26000],
    ]);
    const leftLegendChart = makeChart(makeId, sheetId, 'clusteredColumn', {
      series: [{ id: makeId(), nameRef: '=Sheet1!$B$1', categoriesRef: '=Sheet1!$A$2:$A$3', valuesRef: '=Sheet1!$B$2:$B$3' }],
      elements: {
        ...defaultChartElements('clusteredColumn'),
        legend: { show: true, position: 'left' },
      },
    });
    const noTitleLayout = computeCore2DChartLayout(resolveChartData(doc, leftLegendChart), leftLegendChart, 480, 300);
    // "Revenue" is short so left padding shouldn't consume > 140px
    expect(noTitleLayout.plotArea.x).toBeLessThan(140);

    const withAxisTitleChart: ChartDocument = {
      ...leftLegendChart,
      elements: {
        ...leftLegendChart.elements,
        axes: {
          ...leftLegendChart.elements.axes,
          value: { show: true, title: { show: true, text: 'Axis Title' } },
        },
      },
    };
    const withTitleLayout = computeCore2DChartLayout(resolveChartData(doc, withAxisTitleChart), withAxisTitleChart, 480, 300);
    expect(withTitleLayout.plotArea.x).toBeGreaterThan(noTitleLayout.plotArea.x);
  });
});
