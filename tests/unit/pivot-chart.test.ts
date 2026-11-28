import { describe, expect, it } from 'vitest';
import { createChart } from '../../src/application/chart-commands';
import { createPivotTable, deletePivotTable, updatePivotLayout } from '../../src/application/pivot-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { resolveChartData } from '../../src/charts/chart-data';
import { describePivotChartFields, isPivotChart } from '../../src/charts/pivot-chart';
import { newWorkbook, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';
import { validateWorkbook } from '../../src/core/workbook-validation';
import { loadNativePackage, saveNativePackage } from '../../src/persistence/native';

function makeGuidFactory() {
  let n = 0;
  return () => `00000000-0000-4000-8000-${String(++n).padStart(12, '0')}`;
}

function seedPivotWorkbook(): {
  doc: WorkbookDocument;
  sheetId: string;
  makeId: () => string;
} {
  const makeId = makeGuidFactory();
  const base = newWorkbook(makeId);
  const sheet = base.sheets[0]!;
  const rows: readonly (readonly [string, string, number])[] = [
    ['Region', 'Category', 0],
    ['North', 'Hardware', 120],
    ['North', 'Software', 80],
    ['South', 'Hardware', 90],
    ['East', 'Software', 150],
  ];
  const cells: CellDocument[] = [
    { id: makeId(), row: 1, column: 1, content: { kind: 'text', value: 'Region' } },
    { id: makeId(), row: 1, column: 2, content: { kind: 'text', value: 'Category' } },
    { id: makeId(), row: 1, column: 3, content: { kind: 'text', value: 'Sales' } },
  ];
  for (let i = 1; i < rows.length; i++) {
    const [r, c, v] = rows[i]!;
    cells.push(
      { id: makeId(), row: i + 1, column: 1, content: { kind: 'text', value: r } },
      { id: makeId(), row: i + 1, column: 2, content: { kind: 'text', value: c } },
      { id: makeId(), row: i + 1, column: 3, content: { kind: 'number', value: v } },
    );
  }
  return {
    doc: { ...base, sheets: [{ ...sheet, cells }] },
    sheetId: sheet.id,
    makeId,
  };
}

describe('PivotCharts model, field buttons, bidirectional sync & .exlt persistence', () => {
  it('binds to a PivotTable, updates live on pivot filter/sort changes, round-trips in .exlt v1, and cleans up on pivot delete', async () => {
    const { doc, sheetId, makeId } = seedPivotWorkbook();
    const session = new WorkbookSession(doc, { makeId });

    const b1 = session.capture();
    session.execute(
      createPivotTable(b1, {
        source: { kind: 'range', area: { sheetId, row: 1, column: 1, rows: 5, columns: 3 } },
        target: { sheetId, row: 1, column: 5 },
        rows: [{ fieldName: 'Region' }],
        columns: [{ fieldName: 'Category' }],
        values: [{ id: makeId(), fieldName: 'Sales', aggregation: 'sum' }],
        filters: [{ fieldName: 'Region' }],
      }, makeId),
    );
    const pivot = session.active.pivots![0]!;

    const b2 = session.capture();
    session.execute(
      createChart(b2, {
        chartType: 'clusteredColumn',
        source: { kind: 'pivot', pivotId: pivot.id },
        showFieldButtons: true,
        placement: { sheetId, row: 1, column: 10, width: 480, height: 300 },
      }, makeId),
    );

    const chart = session.active.charts![0]!;
    expect(isPivotChart(chart)).toBe(true);
    expect(chart.showFieldButtons).toBe(true);

    const fields = describePivotChartFields(session.active, chart)!;
    expect(fields.axisFields.map(f => f.label)).toEqual(['Region']);
    expect(fields.legendFields.map(f => f.label)).toEqual(['Category']);
    expect(fields.reportFilters.map(f => f.label)).toEqual(['Region']);
    expect(fields.valueFields.map(f => f.label)).toEqual(['Sum of Sales']);

    const initialData = resolveChartData(session.active, chart);
    expect(initialData.categories).toEqual(['East', 'North', 'South']);
    expect(initialData.series.map(s => s.name)).toEqual(['Hardware', 'Software']);

    // Filter out 'South' via PivotTable layout update -> PivotChart categories update live & field button marks filtered
    const b3 = session.capture();
    session.execute(
      updatePivotLayout(b3, pivot.id, {
        rows: [{ fieldName: 'Region', hiddenItems: ['South'], sort: { direction: 'desc' } }],
      }, makeId),
    );

    const updatedPivot = session.active.pivots![0]!;
    const updatedFields = describePivotChartFields(session.active, chart)!;
    expect(updatedFields.axisFields[0]!.filtered).toBe(true);
    expect(updatedFields.axisFields[0]!.sorted).toBe('desc');

    const filteredData = resolveChartData(session.active, chart);
    expect(filteredData.categories).toEqual(['North', 'East']);
    expect(updatedPivot.extent).toBeDefined();

    // .exlt v1 round-trip preserves PivotChart + showFieldButtons flag
    const saved = await saveNativePackage(session.active);
    const reloaded = await loadNativePackage(saved);
    expect(() => validateWorkbook(reloaded)).not.toThrow();
    expect(reloaded.charts?.[0]?.source).toEqual({ kind: 'pivot', pivotId: pivot.id });
    expect(reloaded.charts?.[0]?.showFieldButtons).toBe(true);

    // Deleting the PivotTable removes the dependent PivotChart cleanly without dangling pivotId errors
    const b4 = session.capture();
    session.execute(deletePivotTable(b4, pivot.id));
    expect(session.active.pivots ?? []).toHaveLength(0);
    expect(session.active.charts ?? []).toHaveLength(0);
    expect(() => validateWorkbook(session.active)).not.toThrow();
  });
});
