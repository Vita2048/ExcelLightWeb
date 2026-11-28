import { describe, expect, it } from 'vitest';
import {
  computeBoxWhiskerStats,
  computeErrorBars,
  computeHistogramBins,
  computeParetoBins,
  computeTrendline,
} from '../../src/charts/analytics';
import { resolveChartData } from '../../src/charts/chart-data';
import { computeChartLayout } from '../../src/charts/layout/advanced-layout';
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
    placement: { sheetId, row: 1, column: 8, offsetX: 0, offsetY: 0, width: 480, height: 300, zIndex: 1 },
    source: { kind: 'range', formula: '=Sheet1!$A$1:$C$5' },
    seriesBy: 'columns',
    series: [
      { id: makeId(), nameRef: '=Sheet1!$B$1', categoriesRef: '=Sheet1!$A$2:$A$5', valuesRef: '=Sheet1!$B$2:$B$5' },
    ],
    elements: defaultChartElements(chartType),
    hiddenEmptyMode: 'gaps',
    showHiddenCells: false,
    styleId: 1,
    paletteId: 'colorful-1',
    ...overrides,
  };
}

describe('chart analytics and advanced Excel chart layouts', () => {
  it('computes histogram bins (Scott auto, binWidth, binCount, overflow/underflow) and Pareto cumulative %', () => {
    const sample = [10, 12, 15, 18, 22, 25, 29, 31, 35, 40, 48, 50];
    const autoBins = computeHistogramBins(sample);
    expect(autoBins.length).toBeGreaterThanOrEqual(2);
    expect(autoBins.reduce((s, b) => s + b.count, 0)).toBe(sample.length);

    const fixedCount = computeHistogramBins(sample, { binMode: 'binCount', binCount: 4 });
    expect(fixedCount).toHaveLength(4);
    expect(fixedCount.reduce((s, b) => s + b.count, 0)).toBe(sample.length);

    const bounded = computeHistogramBins(sample, {
      binMode: 'binWidth',
      binWidth: 10,
      underflowBin: 15,
      overflowBin: 45,
    });
    expect(bounded[0]!.isUnderflow).toBe(true);
    expect(bounded[0]!.count).toBe(3); // 10, 12, 15
    expect(bounded[bounded.length - 1]!.isOverflow).toBe(true);
    expect(bounded[bounded.length - 1]!.count).toBe(2); // 48, 50

    const pareto = computeParetoBins(sample, { binMode: 'binCount', binCount: 4 });
    for (let i = 1; i < pareto.length; i++) {
      expect(pareto[i - 1]!.count).toBeGreaterThanOrEqual(pareto[i]!.count);
      expect(pareto[i]!.cumulativePercent).toBeGreaterThanOrEqual(pareto[i - 1]!.cumulativePercent);
    }
    expect(pareto[pareto.length - 1]!.cumulativePercent).toBeCloseTo(1, 6);
  });

  it('computes Box & Whisker exclusive and inclusive quartiles matching Excel QUARTILE.EXC / QUARTILE.INC', () => {
    const values = [10, 20, 30, 40, 50, 150]; // 150 is an outlier (> Q3 + 1.5*IQR)
    const exc = computeBoxWhiskerStats(values, 'exclusive')!;
    // n=6: rank(0.25)=1.75 -> 10 + 0.75*(20-10) = 17.5; rank(0.75)=5.25 -> 50 + 0.25*(150-50) = 75
    expect(exc.q1).toBeCloseTo(17.5, 6);
    expect(exc.median).toBeCloseTo(35, 6);
    expect(exc.q3).toBeCloseTo(75, 6);

    const inc = computeBoxWhiskerStats([10, 20, 30, 40, 50, 60, 200], 'inclusive')!;
    // n=7: rank(0.25)=2.5 -> 25; median=40; rank(0.75)=5.5 -> 55; IQR=30; fence=100 -> 200 is outlier
    expect(inc.q1).toBeCloseTo(25, 6);
    expect(inc.median).toBeCloseTo(40, 6);
    expect(inc.q3).toBeCloseTo(55, 6);
    expect(inc.outliers).toEqual([200]);
    expect(inc.whiskerHigh).toBe(60);
  });

  it('computes Linear, Exponential, Logarithmic, Polynomial, Power, and Moving Average trendlines + error bars', () => {
    const linearPts = [{ x: 1, y: 3 }, { x: 2, y: 5 }, { x: 3, y: 7 }, { x: 4, y: 9 }];
    const lin = computeTrendline(linearPts, {
      id: 't1', type: 'linear', forward: 2, displayEquation: true, displayRSquared: true,
    });
    expect(lin.rSquared).toBeCloseTo(1, 6);
    expect(lin.points).toEqual([{ x: 1, y: 3 }, { x: 6, y: 13 }]);
    expect(lin.equationText).toContain('2x');

    const fixedInt = computeTrendline(linearPts, { id: 't1b', type: 'linear', intercept: 0 });
    expect(fixedInt.points[0]!.y).toBeCloseTo(fixedInt.points[0]!.x * 2.333333, 2);

    const polyPts = [{ x: 1, y: 1 }, { x: 2, y: 4 }, { x: 3, y: 9 }, { x: 4, y: 16 }];
    const poly = computeTrendline(polyPts, { id: 't2', type: 'polynomial', polynomialOrder: 2 });
    expect(poly.rSquared).toBeCloseTo(1, 5);

    const expPts = [{ x: 1, y: 2 }, { x: 2, y: 4 }, { x: 3, y: 8 }, { x: 4, y: 16 }];
    const exp = computeTrendline(expPts, { id: 't3', type: 'exponential' });
    expect(exp.rSquared).toBeCloseTo(1, 5);

    const logPts = [{ x: 1, y: 0 }, { x: 2, y: Math.log(2) }, { x: 4, y: Math.log(4) }];
    const log = computeTrendline(logPts, { id: 't4', type: 'logarithmic' });
    expect(log.rSquared).toBeCloseTo(1, 5);

    const pow = computeTrendline(polyPts, { id: 't5', type: 'power' });
    expect(pow.rSquared).toBeCloseTo(1, 5);

    const ma = computeTrendline(linearPts, { id: 't6', type: 'movingAverage', period: 2 });
    expect(ma.points).toEqual([{ x: 2, y: 4 }, { x: 3, y: 6 }, { x: 4, y: 8 }]);

    const ebPct = computeErrorBars([100, 200], { direction: 'both', type: 'percentage', value: 10 });
    expect(ebPct).toEqual([{ plus: 10, minus: 10 }, { plus: 20, minus: 20 }]);

    const ebStd = computeErrorBars([10, 20, 30], { direction: 'plus', type: 'stdDev', value: 1 });
    expect(ebStd[0]!.plus).toBeCloseTo(10, 6);
    expect(ebStd[0]!.minus).toBe(0);
  });

  it('lays out Combo (secondary axis), Waterfall, Funnel, Treemap, Sunburst, Stock, Radar, and overlays', () => {
    const { doc, makeId, sheetId } = buildWorkbook([
      ['Item', 'A', 'B', 'C', 'D'],
      ['P1', 100, 120, 90, 110],
      ['P2', 140, 150, 110, 130],
      ['P3', 120, 135, 100, 125],
      ['Total', 360, 405, 300, 365],
    ]);

    const combo = makeChart(makeId, sheetId, 'combo', {
      series: [
        { id: makeId(), nameRef: '=Sheet1!$B$1', categoriesRef: '=Sheet1!$A$2:$A$5', valuesRef: '=Sheet1!$B$2:$B$5', chartType: 'clusteredColumn' },
        { id: makeId(), nameRef: '=Sheet1!$C$1', categoriesRef: '=Sheet1!$A$2:$A$5', valuesRef: '=Sheet1!$C$2:$C$5', chartType: 'line', secondaryAxis: true },
      ],
    });
    const comboLayout = computeChartLayout(resolveChartData(doc, combo), combo, 480, 300);
    expect(comboLayout.bars).toHaveLength(4);
    expect(comboLayout.lines).toHaveLength(1);
    expect(comboLayout.secondaryYAxis).toBeDefined();

    const waterfall = makeChart(makeId, sheetId, 'waterfall', {
      series: [{
        id: makeId(),
        nameRef: '=Sheet1!$B$1',
        categoriesRef: '=Sheet1!$A$2:$A$5',
        valuesRef: '=Sheet1!$B$2:$B$5',
        pointOverrides: [{ index: 3, isTotal: true }],
      }],
    });
    const wfLayout = computeChartLayout(resolveChartData(doc, waterfall), waterfall, 480, 300);
    expect(wfLayout.bars).toHaveLength(4);
    expect(wfLayout.waterfallConnectors).toHaveLength(3);

    const funnel = makeChart(makeId, sheetId, 'funnel');
    expect(computeChartLayout(resolveChartData(doc, funnel), funnel, 480, 300).bars).toHaveLength(4);

    const treemap = makeChart(makeId, sheetId, 'treemap');
    expect(computeChartLayout(resolveChartData(doc, treemap), treemap, 480, 300).treemapRects).toHaveLength(4);

    const sunburst = makeChart(makeId, sheetId, 'sunburst');
    expect(computeChartLayout(resolveChartData(doc, sunburst), sunburst, 480, 300).slices).toHaveLength(4);

    const stockOhlc = makeChart(makeId, sheetId, 'stockOhlc', {
      series: [
        { id: makeId(), name: 'Open', categoriesRef: '=Sheet1!$A$2:$A$5', valuesRef: '=Sheet1!$B$2:$B$5' },
        { id: makeId(), name: 'High', categoriesRef: '=Sheet1!$A$2:$A$5', valuesRef: '=Sheet1!$C$2:$C$5' },
        { id: makeId(), name: 'Low', categoriesRef: '=Sheet1!$A$2:$A$5', valuesRef: '=Sheet1!$D$2:$D$5' },
        { id: makeId(), name: 'Close', categoriesRef: '=Sheet1!$A$2:$A$5', valuesRef: '=Sheet1!$E$2:$E$5' },
      ],
    });
    expect(computeChartLayout(resolveChartData(doc, stockOhlc), stockOhlc, 480, 300).stockStems).toHaveLength(4);

    const radar = makeChart(makeId, sheetId, 'filledRadar');
    const radarLayout = computeChartLayout(resolveChartData(doc, radar), radar, 480, 300);
    expect(radarLayout.radarWeb?.spokes).toHaveLength(4);
    expect(radarLayout.radarWeb?.polygons[0]!.fillOpacity).toBeGreaterThan(0);

    const withOverlays = makeChart(makeId, sheetId, 'clusteredColumn', {
      series: [{
        id: makeId(),
        nameRef: '=Sheet1!$B$1',
        categoriesRef: '=Sheet1!$A$2:$A$5',
        valuesRef: '=Sheet1!$B$2:$B$5',
        trendlines: [{ id: makeId(), type: 'linear', displayEquation: true, displayRSquared: true }],
        errorBars: { direction: 'both', type: 'fixed', value: 15 },
      }],
    });
    const ovLayout = computeChartLayout(resolveChartData(doc, withOverlays), withOverlays, 480, 300);
    expect(ovLayout.trendlines).toHaveLength(1);
    expect(ovLayout.errorBars).toHaveLength(4);
  });
});
