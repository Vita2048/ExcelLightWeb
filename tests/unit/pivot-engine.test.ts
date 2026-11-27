import { describe, expect, it } from 'vitest';
import {
  computePivotOverlays,
  evaluatePivotTable,
  extractPivotCellDrillDown,
  extractPivotRecords,
} from '../../src/calculation/pivot-engine';
import { defaultPivotValueLabel, nextPivotName, type PivotTableDocument } from '../../src/core/pivot';
import { newWorkbook, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';

function sampleWorkbook(): WorkbookDocument {
  let next = 0;
  const base = newWorkbook(() => `id-${++next}`);
  const sheet = base.sheets[0]!;
  const rows: (readonly (string | number)[])[] = [
    ['Region', 'Category', 'Product', 'Amount'],
    ['East', 'Hardware', 'Laptop', 100],
    ['East', 'Hardware', 'Monitor', 200],
    ['East', 'Software', 'Editor', 50],
    ['West', 'Hardware', 'Laptop', 150],
    ['West', 'Software', 'Suite', 300],
    ['West', 'Software', 'Editor', 100],
  ];
  const cells: CellDocument[] = [];
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r]!.length; c++) {
      const value = rows[r]![c]!;
      cells.push({
        id: `c-${r}-${c}`,
        row: r + 1,
        column: c + 1,
        content: typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
      });
    }
  }
  return {
    ...base,
    sheets: [{ ...sheet, cells }],
  };
}

function makePivot(doc: WorkbookDocument, overrides: Partial<PivotTableDocument> = {}): PivotTableDocument {
  const sheetId = doc.sheets[0]!.id;
  return {
    id: 'pivot-1',
    name: 'PivotTable1',
    source: { kind: 'range', area: { sheetId, row: 1, column: 1, rows: 7, columns: 4 } },
    target: { sheetId, row: 1, column: 6 },
    filters: [],
    columns: [{ fieldName: 'Region' }],
    rows: [{ fieldName: 'Product' }],
    values: [{ id: 'v1', fieldName: 'Amount', aggregation: 'sum' }],
    reportLayout: 'compact',
    subtotals: 'top',
    grandTotals: { rows: true, columns: true },
    style: 'Blue',
    styleOptions: { rowHeaders: true, columnHeaders: true, bandedRows: false, bandedColumns: false },
    ...overrides,
  };
}

function gridMap(result: ReturnType<typeof evaluatePivotTable>) {
  const map = new Map<string, typeof result.cells[number]>();
  for (const cell of result.cells) map.set(`${cell.rowOffset}:${cell.colOffset}`, cell);
  return map;
}

describe('pivot engine', () => {
  it('computes all 11 Excel aggregations and labels accurately', () => {
    const doc = sampleWorkbook();
    expect(defaultPivotValueLabel('Amount', 'sum')).toBe('Sum of Amount');
    expect(nextPivotName(doc)).toBe('PivotTable1');

    const { fields, records } = extractPivotRecords(doc, {
      kind: 'range', area: { sheetId: doc.sheets[0]!.id, row: 1, column: 1, rows: 7, columns: 4 },
    });
    expect(fields).toEqual(['Region', 'Category', 'Product', 'Amount']);
    expect(records).toHaveLength(6);

    const expected: Record<string, number> = {
      sum: 900,
      count: 6,
      average: 150,
      max: 300,
      min: 50,
      product: 100 * 200 * 50 * 150 * 300 * 100,
      countNumbers: 6,
      varp: ((50 ** 2) * 2 + (0 ** 2) * 0 + (50 ** 2) + (150 ** 2) + (100 ** 2)) / 6,
    };
    const vals = [100, 200, 50, 150, 300, 100];
    const mean = 150;
    const ss = vals.reduce((acc, v) => acc + (v - mean) ** 2, 0);
    expected.varp = ss / 6;
    expected.stdDevp = Math.sqrt(ss / 6);
    expected.var = ss / 5;
    expected.stdDev = Math.sqrt(ss / 5);

    for (const [agg, want] of Object.entries(expected)) {
      const pivot = makePivot(doc, {
        columns: [],
        rows: [],
        grandTotals: { rows: false, columns: true },
        values: [{ id: 'v1', fieldName: 'Amount', aggregation: agg as PivotTableDocument['values'][number]['aggregation'] }],
      });
      const res = evaluatePivotTable(doc, pivot);
      const dataCell = res.cells.find(cell => cell.role === 'data');
      expect(dataCell?.content).toEqual({ kind: 'number', value: expect.closeTo(want, 6) });
    }
  });

  it('evaluates compact layout with Column/Row labels and grand totals', () => {
    const doc = sampleWorkbook();
    const res = evaluatePivotTable(doc, makePivot(doc));
    const grid = gridMap(res);

    expect(grid.get('0:0')?.content).toEqual({ kind: 'text', value: 'Sum of Amount' });
    expect(grid.get('0:1')?.content).toEqual({ kind: 'text', value: 'Column Labels' });
    expect(grid.get('1:0')?.content).toEqual({ kind: 'text', value: 'Row Labels' });
    expect(grid.get('1:1')?.content).toEqual({ kind: 'text', value: 'East' });
    expect(grid.get('1:2')?.content).toEqual({ kind: 'text', value: 'West' });
    expect(grid.get('1:3')?.content).toEqual({ kind: 'text', value: 'Grand Total' });

    // Sorted ascending: Editor, Laptop, Monitor, Suite, Grand Total
    expect(grid.get('2:0')?.content).toEqual({ kind: 'text', value: 'Editor' });
    expect(grid.get('2:1')?.content).toEqual({ kind: 'number', value: 50 });
    expect(grid.get('2:2')?.content).toEqual({ kind: 'number', value: 100 });
    expect(grid.get('2:3')?.content).toEqual({ kind: 'number', value: 150 });

    expect(grid.get('6:0')?.content).toEqual({ kind: 'text', value: 'Grand Total' });
    expect(grid.get('6:1')?.content).toEqual({ kind: 'number', value: 350 });
    expect(grid.get('6:2')?.content).toEqual({ kind: 'number', value: 550 });
    expect(grid.get('6:3')?.content).toEqual({ kind: 'number', value: 900 });
  });

  it('supports hierarchies, subtotals (top/bottom/off), outline/tabular layouts, collapsedKeys, and report filters', () => {
    const doc = sampleWorkbook();
    const multi = makePivot(doc, {
      columns: [],
      rows: [{ fieldName: 'Region' }, { fieldName: 'Category' }],
      filters: [{ fieldName: 'Product', hiddenItems: ['Monitor'] }],
      reportLayout: 'tabular',
      subtotals: 'bottom',
    });
    const res = evaluatePivotTable(doc, multi);
    const texts = res.cells.map(c => c.content.kind === 'text' ? c.content.value : c.content.kind === 'number' ? c.content.value : '');
    expect(texts).toContain('Product');
    expect(texts).toContain('(Multiple Items)');
    expect(texts).toContain('East Total');
    expect(texts).toContain('West Total');
    // Monitor (200) is filtered out -> East = 150, West = 550, Grand Total = 700
    expect(res.cells.find(c => c.role === 'grandTotal' && c.content.kind === 'number')?.content).toEqual({ kind: 'number', value: 700 });

    const collapsed = evaluatePivotTable(doc, { ...multi, collapsedKeys: ['East'] });
    const eastCollapsedTexts = collapsed.cells.map(c => c.content.kind === 'text' ? c.content.value : '');
    expect(eastCollapsedTexts).not.toContain('East Total');
    const eastHeader = collapsed.cells.find(c => c.collapsibleKey === 'East');
    expect(eastHeader?.collapsed).toBe(true);
  });

  it('supports Show Values As calculations and table sources excluding totalRow', () => {
    const doc = sampleWorkbook();
    const sheetId = doc.sheets[0]!.id;
    const withTable: WorkbookDocument = {
      ...doc,
      names: [{
        id: 't1',
        name: 'SalesTable',
        area: { sheetId, row: 1, column: 1, rows: 7, columns: 4 },
        isTable: true,
        design: { style: 'Blue', headerRow: true, totalRow: false, bandedRows: true },
      }],
    };
    const pctPivot = makePivot(withTable, {
      source: { kind: 'table', tableId: 't1' },
      columns: [],
      rows: [{ fieldName: 'Region', sort: { direction: 'desc', byValueFieldId: 'v1' } }],
      values: [{ id: 'v1', fieldName: 'Amount', aggregation: 'sum', showValuesAs: 'percentOfColumnTotal' }],
    });
    const res = evaluatePivotTable(withTable, pctPivot);
    const grid = gridMap(res);
    // West (550) sorted first descending by Amount: 550/900
    expect(grid.get('1:0')?.content).toEqual({ kind: 'text', value: 'West' });
    expect(grid.get('1:1')?.content).toEqual({ kind: 'number', value: expect.closeTo(550 / 900, 6) });
    expect(grid.get('1:1')?.style?.numberFormat).toBe('0.00%');
    expect(grid.get('2:0')?.content).toEqual({ kind: 'text', value: 'East' });
    expect(grid.get('2:1')?.content).toEqual({ kind: 'number', value: expect.closeTo(350 / 900, 6) });
  });

  it('supports Compact indent, Outline repeatLabels/blankRows, and Tabular side-by-side rows', () => {
    const doc = sampleWorkbook();

    const tabular = makePivot(doc, {
      columns: [],
      rows: [{ fieldName: 'Region' }, { fieldName: 'Category' }],
      reportLayout: 'tabular',
      subtotals: 'bottom',
      blankRows: true,
      repeatLabels: true,
    });
    const tabRes = evaluatePivotTable(doc, tabular);
    const tabGrid = gridMap(tabRes);
    // Tabular row 1 starts immediately with East + Hardware on the same row
    expect(tabGrid.get('1:0')?.content).toEqual({ kind: 'text', value: 'East' });
    expect(tabGrid.get('1:1')?.content).toEqual({ kind: 'text', value: 'Hardware' });
    expect(tabGrid.get('1:2')?.content).toEqual({ kind: 'number', value: 300 });
    // Repeat labels keeps East in col 0 on row 2
    expect(tabGrid.get('2:0')?.content).toEqual({ kind: 'text', value: 'East' });
    expect(tabGrid.get('2:1')?.content).toEqual({ kind: 'text', value: 'Software' });
    // Subtotal on row 3, blank separator on row 4, West starts on row 5
    expect(tabGrid.get('3:0')?.content).toEqual({ kind: 'text', value: 'East Total' });
    expect(tabGrid.get('4:0')).toBeUndefined();
    expect(tabGrid.get('5:0')?.content).toEqual({ kind: 'text', value: 'West' });

    const outline = makePivot(doc, {
      columns: [],
      rows: [{ fieldName: 'Region' }, { fieldName: 'Category' }],
      reportLayout: 'outline',
      subtotals: 'off',
      repeatLabels: false,
    });
    const outGrid = gridMap(evaluatePivotTable(doc, outline));
    expect(outGrid.get('1:0')?.content).toEqual({ kind: 'text', value: 'East' });
    expect(outGrid.get('1:1')).toBeUndefined();
    expect(outGrid.get('2:0')?.content).toEqual({ kind: 'text', value: '' });
    expect(outGrid.get('2:1')?.content).toEqual({ kind: 'text', value: 'Hardware' });

    const compact = makePivot(doc, {
      columns: [],
      rows: [{ fieldName: 'Region' }, { fieldName: 'Category' }],
      reportLayout: 'compact',
    });
    const compRes = evaluatePivotTable(doc, compact);
    const eastCell = compRes.cells.find(c => c.content.kind === 'text' && c.content.value === 'East');
    const hardwareCell = compRes.cells.find(c => c.content.kind === 'text' && c.content.value === 'Hardware');
    const grandRowCell = compRes.cells.find(c => c.content.kind === 'text' && c.content.value === 'Grand Total' && c.colOffset === 0);
    expect(eastCell?.style?.alignment).toBe('left');
    expect(eastCell?.style?.indent).toBe(1);
    expect(hardwareCell?.style?.alignment).toBe('left');
    expect(hardwareCell?.style?.indent).toBe(2);
    expect(grandRowCell?.style?.indent).toBeUndefined();
  });

  it('computes interactive overlays and drill-down record buckets', () => {
    const doc = sampleWorkbook();
    const sheetId = doc.sheets[0]!.id;
    const pivot = makePivot(doc, {
      target: { sheetId, row: 1, column: 6 },
      filters: [{ fieldName: 'Category' }],
      columns: [{ fieldName: 'Region' }],
      rows: [{ fieldName: 'Category' }, { fieldName: 'Product' }],
    });
    const withPivot: WorkbookDocument = { ...doc, pivots: [pivot] };

    const overlays = computePivotOverlays(withPivot, sheetId);
    expect(overlays).toHaveLength(1);
    const ov = overlays[0]!;
    expect(ov.toggles.some(t => t.key === 'Hardware' && !t.collapsed && t.level === 0)).toBe(true);
    expect(ov.filters.map(f => f.kind)).toEqual(['reportFilter', 'rowLabel', 'colLabel']);

    const hiddenButtons: WorkbookDocument = {
      ...doc,
      pivots: [{ ...pivot, showButtons: false, showFieldHeaders: false }],
    };
    const suppressed = computePivotOverlays(hiddenButtons, sheetId)[0]!;
    expect(suppressed.toggles).toEqual([]);
    expect(suppressed.filters).toEqual([]);

    // Data area cell drill-down returns matching records; header returns undefined.
    const evaluated = evaluatePivotTable(doc, pivot);
    const headerCell = evaluated.cells.find(c => c.role === 'header')!;
    expect(extractPivotCellDrillDown(doc, pivot, {
      sheetId, row: pivot.target.row + headerCell.rowOffset, column: pivot.target.column + headerCell.colOffset,
    })).toBeUndefined();

    const dataCell = evaluated.cells.find(c => c.role === 'data' && c.content.kind === 'number' && c.content.value === 100)!;
    const drill = extractPivotCellDrillDown(doc, pivot, {
      sheetId, row: pivot.target.row + dataCell.rowOffset, column: pivot.target.column + dataCell.colOffset,
    });
    expect(drill).toBeDefined();
    expect(drill!.fields).toEqual(['Region', 'Category', 'Product', 'Amount']);
    expect(drill!.records).toHaveLength(1);
    expect(drill!.records[0]!.Amount).toBe(100);

    const grandCell = evaluated.cells.find(c => c.role === 'grandTotal' && c.content.kind === 'number' && c.content.value === 900)!;
    const allDrill = extractPivotCellDrillDown(doc, pivot, {
      sheetId, row: pivot.target.row + grandCell.rowOffset, column: pivot.target.column + grandCell.colOffset,
    });
    expect(allDrill?.records).toHaveLength(6);
  });

  it('inherits source column number formats for numeric aggregations and preserves explicit overrides', () => {
    const base = sampleWorkbook();
    const sheet = base.sheets[0]!;
    const formatted: WorkbookDocument = {
      ...base,
      sheets: [{
        ...sheet,
        cells: sheet.cells.map(cell => cell.column === 4 && cell.row > 1
          ? { ...cell, style: { ...cell.style, numberFormat: '$#,##0.00' } }
          : cell),
      }],
    };
    const sumPivot = makePivot(formatted, {
      columns: [],
      rows: [{ fieldName: 'Region' }],
      values: [{ id: 'v1', fieldName: 'Amount', aggregation: 'sum' }],
    });
    const sumRes = evaluatePivotTable(formatted, sumPivot);
    const dataCell = sumRes.cells.find(c => c.role === 'data');
    expect(dataCell?.style?.numberFormat).toBe('$#,##0.00');
    expect(sumRes.cells.find(c => c.role === 'grandTotal' && c.content.kind === 'number')?.style?.numberFormat).toBe('$#,##0.00');

    // Count aggregations do not inherit currency formatting from source numeric column.
    const countPivot = makePivot(formatted, {
      columns: [],
      rows: [{ fieldName: 'Region' }],
      values: [{ id: 'v1', fieldName: 'Amount', aggregation: 'count' }],
    });
    expect(evaluatePivotTable(formatted, countPivot).cells.find(c => c.role === 'data')?.style?.numberFormat).toBeUndefined();

    // Explicit value field format overrides source format.
    const overridePivot = makePivot(formatted, {
      columns: [],
      rows: [{ fieldName: 'Region' }],
      values: [{ id: 'v1', fieldName: 'Amount', aggregation: 'sum', numberFormat: '0.0' }],
    });
    expect(evaluatePivotTable(formatted, overridePivot).cells.find(c => c.role === 'data')?.style?.numberFormat).toBe('0.0');
  });
});
