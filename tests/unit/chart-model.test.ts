import { describe, expect, it } from 'vitest';
import { inferChartSeriesFromSource, resolveChartData, resolveFormulaExtent } from '../../src/charts/chart-data';
import { defaultChartElements, nextChartName, type ChartDocument } from '../../src/core/chart';
import { DocumentError, newWorkbook, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';
import { validateWorkbook } from '../../src/core/workbook-validation';

function makeGuidFactory() {
  let n = 0;
  return () => `00000000-0000-4000-8000-${String(++n).padStart(12, '0')}`;
}

function buildMatrixWorkbook(rows: readonly (readonly (string | number | null)[])[]): {
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
  return {
    doc: { ...base, sheets: [{ ...sheet, cells }] },
    makeId,
    sheetId: sheet.id,
  };
}

function baseChart(makeId: () => string, sheetId: string, overrides: Partial<ChartDocument> = {}): ChartDocument {
  return {
    id: makeId(),
    name: 'Chart 1',
    chartType: 'clusteredColumn',
    placement: { sheetId, row: 2, column: 6, offsetX: 0, offsetY: 0, width: 480, height: 300, zIndex: 1 },
    source: { kind: 'range', formula: '=Sheet1!$A$1:$C$4' },
    seriesBy: 'columns',
    series: [{ id: makeId(), name: 'Sales', valuesRef: '=Sheet1!$B$2:$B$4', categoriesRef: '=Sheet1!$A$2:$A$4' }],
    elements: defaultChartElements('clusteredColumn'),
    hiddenEmptyMode: 'gaps',
    showHiddenCells: false,
    styleId: 1,
    paletteId: 'colorful-1',
    ...overrides,
  };
}

describe('chart model and data resolution', () => {
  it('validates chart invariants, unique names, and numeric bounds', () => {
    const { doc, makeId, sheetId } = buildMatrixWorkbook([
      ['Month', 'Sales', 'Cost'],
      ['Jan', 100, 60],
      ['Feb', 140, 80],
    ]);
    const validChart = baseChart(makeId, sheetId);
    expect(() => validateWorkbook({ ...doc, charts: [validChart] })).not.toThrow();
    expect(nextChartName({ ...doc, charts: [validChart] })).toBe('Chart 2');

    // Case-insensitive duplicate name
    const dupName = baseChart(makeId, sheetId, { name: 'chart 1' });
    expect(() => validateWorkbook({ ...doc, charts: [validChart, dupName] })).toThrow(DocumentError);

    // Invalid width
    const tooNarrow = baseChart(makeId, sheetId, {
      placement: { ...validChart.placement, width: 80 },
    });
    expect(() => validateWorkbook({ ...doc, charts: [tooNarrow] })).toThrow(DocumentError);

    // Invalid polynomial trendline order
    const badTrend = baseChart(makeId, sheetId, {
      series: [{
        ...validChart.series[0]!,
        trendlines: [{ id: makeId(), type: 'polynomial', polynomialOrder: 9 }],
      }],
    });
    expect(() => validateWorkbook({ ...doc, charts: [badTrend] })).toThrow(DocumentError);
  });

  it('infers column vs row series orientation and Scatter/Bubble bindings', () => {
    // Tall data -> columns orientation
    const tall = buildMatrixWorkbook([
      ['Quarter', 'North', 'South'],
      ['Q1', 10, 20],
      ['Q2', 15, 25],
      ['Q3', 18, 30],
      ['Q4', 22, 35],
    ]);
    const inferredTall = inferChartSeriesFromSource(
      tall.doc,
      { kind: 'range', formula: '=Sheet1!$A$1:$C$5' },
      tall.sheetId,
      'clusteredColumn',
    );
    expect(inferredTall.seriesBy).toBe('columns');
    expect(inferredTall.series).toHaveLength(2);
    expect(inferredTall.series[0]).toMatchObject({
      name: 'North',
      nameRef: "='Sheet1'!$B$1",
      categoriesRef: "='Sheet1'!$A$2:$A$5",
      valuesRef: "='Sheet1'!$B$2:$B$5",
    });

    // Wide data -> rows orientation
    const wide = buildMatrixWorkbook([
      ['Metric', 'Jan', 'Feb', 'Mar', 'Apr', 'May'],
      ['Revenue', 10, 20, 30, 40, 50],
      ['Expense', 5, 10, 12, 15, 20],
    ]);
    const inferredWide = inferChartSeriesFromSource(
      wide.doc,
      { kind: 'range', formula: '=Sheet1!$A$1:$F$3' },
      wide.sheetId,
      'line',
    );
    expect(inferredWide.seriesBy).toBe('rows');
    expect(inferredWide.series).toHaveLength(2);
    expect(inferredWide.series[1]).toMatchObject({
      name: 'Expense',
      nameRef: "='Sheet1'!$A$3",
      categoriesRef: "='Sheet1'!$B$1:$F$1",
      valuesRef: "='Sheet1'!$B$3:$F$3",
    });

    // Bubble chart -> X, Y, Size
    const bubble = buildMatrixWorkbook([
      ['X', 'Y', 'Size'],
      [1, 10, 5],
      [2, 20, 12],
    ]);
    const inferredBubble = inferChartSeriesFromSource(
      bubble.doc,
      { kind: 'range', formula: '=Sheet1!$A$1:$C$3' },
      bubble.sheetId,
      'bubble',
    );
    expect(inferredBubble.series).toHaveLength(1);
    expect(inferredBubble.series[0]).toMatchObject({
      name: 'Y',
      xValuesRef: "='Sheet1'!$A$2:$A$3",
      valuesRef: "='Sheet1'!$B$2:$B$3",
      bubbleSizesRef: "='Sheet1'!$C$2:$C$3",
    });
  });

  it('resolves dynamic spill ranges (A1#), formatted values, hidden/filtered rows, and hiddenEmptyMode', () => {
    const makeId = makeGuidFactory();
    const base = newWorkbook(makeId);
    const sheet = base.sheets[0]!;
    const cells: CellDocument[] = [
      { id: makeId(), row: 1, column: 1, content: { kind: 'text', value: 'Category' } },
      { id: makeId(), row: 1, column: 2, content: { kind: 'text', value: 'Revenue' } },
      {
        id: makeId(),
        row: 2,
        column: 1,
        content: { kind: 'formula', source: '=SEQUENCE(3)' },
        cache: { kind: 'text', value: 'Alpha' },
        spillExtent: { rows: 3, columns: 2 },
      },
      { id: makeId(), row: 2, column: 2, content: { kind: 'blank' }, cache: { kind: 'number', value: 1200 }, style: { numberFormat: '$#,##0' }, spillOwner: 'anchor', spillKind: 'formula' },
      { id: makeId(), row: 3, column: 1, content: { kind: 'blank' }, cache: { kind: 'text', value: 'Beta' }, spillOwner: 'anchor', spillKind: 'formula' },
      // Blank cell at (3,2) to test gaps vs zero vs connect
      { id: makeId(), row: 4, column: 1, content: { kind: 'blank' }, cache: { kind: 'text', value: 'Gamma' }, spillOwner: 'anchor', spillKind: 'formula' },
      { id: makeId(), row: 4, column: 2, content: { kind: 'blank' }, cache: { kind: 'number', value: 2500 }, style: { numberFormat: '$#,##0' }, spillOwner: 'anchor', spillKind: 'formula' },
      { id: makeId(), row: 6, column: 1, content: { kind: 'text', value: 'Dynamic Chart Title' } },
    ];
    // Manually hide row 3
    const docWithHidden: WorkbookDocument = {
      ...base,
      sheets: [{ ...sheet, cells, rows: [{ index: 3, hidden: true }] }],
    };

    const spillArea = resolveFormulaExtent(docWithHidden, sheet.id, '=Sheet1!$A$2#');
    expect(spillArea).toEqual({ sheetId: sheet.id, row: 2, column: 1, rows: 3, columns: 2 });

    const chart = baseChart(makeId, sheet.id, {
      chartType: 'line',
      source: { kind: 'range', formula: '=Sheet1!$A$1:$B$4' },
      series: [{
        id: makeId(),
        nameRef: '=Sheet1!$B$1',
        categoriesRef: '=Sheet1!$A$2:$A$4',
        valuesRef: '=Sheet1!$B$2:$B$4',
      }],
      elements: {
        ...defaultChartElements('line'),
        title: { show: true, formulaRef: '=Sheet1!$A$6' },
      },
      hiddenEmptyMode: 'gaps',
      showHiddenCells: false,
    });

    // Hidden row 3 excluded by default
    const resExcluded = resolveChartData(docWithHidden, chart);
    expect(resExcluded.title).toBe('Dynamic Chart Title');
    expect(resExcluded.categories).toEqual(['Alpha', 'Gamma']);
    expect(resExcluded.series[0]!.points.map(p => p.value)).toEqual([1200, 2500]);
    expect(resExcluded.series[0]!.points[0]!.formatted).toBe('$1,200');

    // Include hidden row 3 with zero vs connect
    const resZero = resolveChartData(docWithHidden, {
      ...chart,
      showHiddenCells: true,
      hiddenEmptyMode: 'zero',
    });
    expect(resZero.categories).toEqual(['Alpha', 'Beta', 'Gamma']);
    expect(resZero.series[0]!.points.map(p => p.value)).toEqual([1200, 0, 2500]);

    const resConnect = resolveChartData(docWithHidden, {
      ...chart,
      showHiddenCells: true,
      hiddenEmptyMode: 'connect',
    });
    expect(resConnect.series[0]!.connectGaps).toBe(true);
    expect(resConnect.series[0]!.points.map(p => p.value)).toEqual([1200, null, 2500]);
  });

  it('dynamically updates Table-backed charts when Table grows and excludes totalRow', () => {
    const { doc, makeId, sheetId } = buildMatrixWorkbook([
      ['Region', 'Sales'],
      ['East', 100],
      ['West', 200],
      ['Total', 300],
    ]);
    const tableId = makeId();
    const withTable: WorkbookDocument = {
      ...doc,
      names: [{
        id: tableId,
        name: 'SalesTable',
        isTable: true,
        area: { sheetId, row: 1, column: 1, rows: 4, columns: 2 },
        design: { style: 'Blue', headerRow: true, totalRow: true },
      }],
    };
    const chart = baseChart(makeId, sheetId, {
      source: { kind: 'table', tableId },
      series: [],
    });

    const resolved1 = resolveChartData(withTable, chart);
    expect(resolved1.categories).toEqual(['East', 'West']);
    expect(resolved1.series[0]!.points.map(p => p.value)).toEqual([100, 200]);

    // Simulate adding a row to the table before total row
    const sheet = withTable.sheets[0]!;
    const grownDoc: WorkbookDocument = {
      ...withTable,
      sheets: [{
        ...sheet,
        cells: [
          ...sheet.cells.filter(c => c.row < 4),
          { id: makeId(), row: 4, column: 1, content: { kind: 'text', value: 'North' } },
          { id: makeId(), row: 4, column: 2, content: { kind: 'number', value: 350 } },
          { id: makeId(), row: 5, column: 1, content: { kind: 'text', value: 'Total' } },
          { id: makeId(), row: 5, column: 2, content: { kind: 'number', value: 650 } },
        ],
      }],
      names: [{
        ...withTable.names[0]!,
        area: { sheetId, row: 1, column: 1, rows: 5, columns: 2 },
      }],
    };

    const resolved2 = resolveChartData(grownDoc, chart);
    expect(resolved2.categories).toEqual(['East', 'West', 'North']);
    expect(resolved2.series[0]!.points.map(p => p.value)).toEqual([100, 200, 350]);
  });

  it('excludes AutoFilter-filtered rows and resolves PivotTable sources skipping report filters and grand totals', () => {
    const { doc, makeId, sheetId } = buildMatrixWorkbook([
      ['Region', 'Sales'],
      ['East', 100],
      ['West', 200],
      ['North', 300],
    ]);
    const filteredDoc: WorkbookDocument = {
      ...doc,
      sheets: [{
        ...doc.sheets[0]!,
        filter: {
          row: 1,
          column: 1,
          rows: 4,
          columns: 2,
          criteria: [{ column: 1, allowed: ['East', 'North'] }],
        },
      }],
    };
    const filterChart = baseChart(makeId, sheetId, {
      source: { kind: 'range', formula: '=Sheet1!$A$1:$B$4' },
      series: [{
        id: makeId(),
        nameRef: '=Sheet1!$B$1',
        categoriesRef: '=Sheet1!$A$2:$A$4',
        valuesRef: '=Sheet1!$B$2:$B$4',
      }],
    });
    const filteredResolved = resolveChartData(filteredDoc, filterChart);
    expect(filteredResolved.categories).toEqual(['East', 'North']);
    expect(filteredResolved.series[0]!.points.map(p => p.value)).toEqual([100, 300]);

    // Pivot source with filter row + grand total row
    const pivotWorkbook = buildMatrixWorkbook([
      ['Filter', 'All'],
      [null, null],
      ['Region', 'Sum of Sales'],
      ['East', 100],
      ['West', 200],
      ['Grand Total', 300],
    ]);
    const pivotId = makeId();
    const withPivot: WorkbookDocument = {
      ...pivotWorkbook.doc,
      pivots: [{
        id: pivotId,
        name: 'PivotTable1',
        source: { kind: 'range', area: { sheetId: pivotWorkbook.sheetId, row: 1, column: 1, rows: 1, columns: 1 } },
        target: { sheetId: pivotWorkbook.sheetId, row: 1, column: 1 },
        extent: { row: 1, column: 1, rows: 6, columns: 2 },
        filters: [{ fieldName: 'Filter' }],
        columns: [],
        rows: [{ fieldName: 'Region' }],
        values: [{ id: makeId(), fieldName: 'Sales', aggregation: 'sum' }],
        reportLayout: 'compact',
        subtotals: 'top',
        grandTotals: { rows: true, columns: true },
        style: 'Blue',
        styleOptions: { rowHeaders: true, columnHeaders: true, bandedRows: false, bandedColumns: false },
      }],
    };
    const pivotChart = baseChart(makeId, pivotWorkbook.sheetId, {
      source: { kind: 'pivot', pivotId },
      series: [],
    });
    const pivotResolved = resolveChartData(withPivot, pivotChart);
    expect(pivotResolved.categories).toEqual(['East', 'West']);
    expect(pivotResolved.series[0]!.points.map(p => p.value)).toEqual([100, 200]);
  });
});
