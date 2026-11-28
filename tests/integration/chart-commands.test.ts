import { describe, expect, it } from 'vitest';
import {
  createChart,
  deleteChart,
  moveChartToSheet,
  moveResizeChart,
  reorderChart,
  updateChart,
} from '../../src/application/chart-commands';
import { changeSheet, editAxis } from '../../src/application/structure-commands';
import { createTable, renameTable } from '../../src/application/table-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { resolveChartData } from '../../src/charts/chart-data';
import { newWorkbook, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';
import { sameAuthoredWorkbook } from '../../src/core/workbook-delta';
import { loadNativePackage, saveNativePackage } from '../../src/persistence/native/native-package';

function makeGuidFactory() {
  let n = 0;
  return () => `00000000-0000-4000-8000-${String(++n).padStart(12, '0')}`;
}

function buildSalesWorkbook(): { doc: WorkbookDocument; makeId: () => string } {
  const makeId = makeGuidFactory();
  const base = newWorkbook(makeId);
  const sheet = base.sheets[0]!;
  const rows: (readonly (string | number)[])[] = [
    ['Month', 'Revenue', 'Cost'],
    ['Jan', 100, 60],
    ['Feb', 150, 75],
    ['Mar', 200, 90],
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
  return {
    doc: { ...base, sheets: [{ ...sheet, cells }] },
    makeId,
  };
}

describe('chart commands, structural rewrites, and native .exlt persistence', () => {
  it('executes create, update, move/resize, reorder, moveToSheet, delete, and undo/redo cleanly', () => {
    const { doc, makeId } = buildSalesWorkbook();
    const sheet1 = doc.sheets[0]!.id;
    const session = new WorkbookSession(doc, { makeId });

    // Create two charts on Sheet1
    session.execute(createChart(session.capture(), {
      chartType: 'clusteredColumn',
      source: { kind: 'range', formula: '=Sheet1!$A$1:$C$4' },
      placement: { sheetId: sheet1, row: 2, column: 5, width: 480, height: 300 },
    }, makeId));

    session.execute(createChart(session.capture(), {
      chartType: 'lineMarkers',
      source: { kind: 'range', formula: '=Sheet1!$A$1:$B$4' },
      placement: { sheetId: sheet1, row: 10, column: 5, width: 400, height: 240 },
    }, makeId));

    expect(session.active.charts).toHaveLength(2);
    const c1 = session.active.charts![0]!;
    const c2 = session.active.charts![1]!;
    expect(c1.name).toBe('Chart 1');
    expect(c2.name).toBe('Chart 2');
    expect(c1.placement.zIndex).toBe(1);
    expect(c2.placement.zIndex).toBe(2);
    expect(c1.series).toHaveLength(2);

    // Reorder: bring c1 to front
    session.execute(reorderChart(session.capture(), c1.id, 'bringToFront'));
    expect(session.active.charts![0]!.placement.zIndex).toBe(2);
    expect(session.active.charts![1]!.placement.zIndex).toBe(1);

    // Move/resize c1
    session.execute(moveResizeChart(session.capture(), c1.id, { row: 4, column: 6, width: 600, height: 360 }));
    expect(session.active.charts![0]!.placement.width).toBe(600);
    expect(session.active.charts![0]!.placement.row).toBe(4);

    // Update c1 type & title
    session.execute(updateChart(session.capture(), c1.id, {
      name: 'Revenue Chart',
      chartType: 'stackedColumn',
      elements: { ...c1.elements, title: { show: true, text: 'Quarterly Revenue' } },
    }));
    expect(session.active.charts![0]!.name).toBe('Revenue Chart');
    expect(session.active.charts![0]!.chartType).toBe('stackedColumn');

    // Add Sheet2 and move c2 there
    const addSheet = changeSheet(session.capture(), sheet1, { kind: 'add', after: true }, makeId);
    session.execute(addSheet.command);
    const sheet2 = addSheet.active;
    session.execute(moveChartToSheet(session.capture(), c2.id, sheet2, { row: 3, column: 3 }));
    expect(session.active.charts![1]!.placement.sheetId).toBe(sheet2);
    expect(session.active.charts![1]!.placement.row).toBe(3);

    // Delete c1 and undo
    const beforeDelete = session.active;
    session.execute(deleteChart(session.capture(), c1.id));
    expect(session.active.charts).toHaveLength(1);
    session.undo();
    expect(sameAuthoredWorkbook(session.active, beforeDelete)).toBe(true);
  });

  it('rewrites chart formula references across axis edits, sheet rename/copy/delete, and table rename', () => {
    const { doc, makeId } = buildSalesWorkbook();
    const sheet1 = doc.sheets[0]!.id;
    const session = new WorkbookSession(doc, { makeId });

    // Create a Table on A1:C4 and a chart bound to its range + a second range chart
    session.execute(createTable(session.capture(), {
      id: makeId(),
      name: 'SalesTbl',
      area: { sheetId: sheet1, row: 1, column: 1, rows: 4, columns: 3 },
      isTable: true,
    }, true, makeId));

    session.execute(createChart(session.capture(), {
      chartType: 'clusteredColumn',
      source: { kind: 'range', formula: '=Sheet1!$A$1:$C$4' },
      placement: { sheetId: sheet1, row: 6, column: 2 },
      series: [{
        id: makeId(),
        nameRef: '=Sheet1!$B$1',
        categoriesRef: '=Sheet1!$A$2:$A$4',
        valuesRef: '=SalesTbl[Revenue]',
      }],
    }, makeId));

    // Insert 2 rows at row 1 -> shifts range refs & anchor row
    session.execute(editAxis(session.capture(), sheet1, { axis: 'row', start: 1, count: 2, delete: false }));
    let chart = session.active.charts![0]!;
    expect(chart.source).toEqual({ kind: 'range', formula: '=Sheet1!$A$3:$C$6' });
    expect(chart.series[0]!.categoriesRef).toBe('=Sheet1!$A$4:$A$6');
    expect(chart.placement.row).toBe(8);

    // Rename Table -> rewrites structured reference
    const tableId = session.active.names[0]!.id;
    session.execute(renameTable(session.capture(), tableId, 'OrdersTbl'));
    chart = session.active.charts![0]!;
    expect(chart.series[0]!.valuesRef).toBe('=OrdersTbl[Revenue]');

    // Rename Sheet1 -> RevenueData
    session.execute(changeSheet(session.capture(), sheet1, { kind: 'rename', name: 'RevenueData' }, makeId).command);
    chart = session.active.charts![0]!;
    expect(chart.source).toEqual({ kind: 'range', formula: "='RevenueData'!$A$3:$C$6" });
    expect(chart.series[0]!.nameRef).toBe("='RevenueData'!$B$3");

    // Copy Sheet -> clones chart onto new sheet with rewritten self-sheet refs
    const copyRes = changeSheet(session.capture(), sheet1, { kind: 'place', before: '', copy: true }, makeId);
    session.execute(copyRes.command);
    expect(session.active.charts).toHaveLength(2);
    const copiedChart = session.active.charts![1]!;
    expect(copiedChart.id).not.toBe(chart.id);
    expect(copiedChart.name).toBe('Chart 2');
    expect(copiedChart.placement.sheetId).toBe(copyRes.active);
    expect(copiedChart.source).toEqual({ kind: 'range', formula: "='Sheet1'!$A$3:$C$6" });

    // Verify resolved data on copied sheet still evaluates cleanly
    const resolved = resolveChartData(session.active, copiedChart);
    expect(resolved.categories).toEqual(['Jan', 'Feb', 'Mar']);

    // Delete original sheet -> drops chart placed on it
    session.execute(changeSheet(session.capture(), sheet1, { kind: 'delete' }, makeId).command);
    expect(session.active.charts).toHaveLength(1);
    expect(session.active.charts![0]!.id).toBe(copiedChart.id);
  });

  it('round-trips charts with styles, trendlines, error bars, and overrides through native .exlt v1', async () => {
    const { doc, makeId } = buildSalesWorkbook();
    const sheet1 = doc.sheets[0]!.id;
    const session = new WorkbookSession(doc, { makeId });

    session.execute(createChart(session.capture(), {
      chartType: 'waterfall',
      source: { kind: 'range', formula: '=Sheet1!$A$1:$C$4' },
      placement: { sheetId: sheet1, row: 2, column: 5, offsetX: 12, offsetY: 8, width: 520, height: 320, mode: 'moveOnly', zIndex: 1 },
      hiddenEmptyMode: 'connect',
      showHiddenCells: true,
      styleId: 3,
      paletteId: 'monochrome-2',
      chartAreaStyle: { fillRgb: '#14171C', borderRgb: '#30363D', borderWidth: 1 },
      series: [{
        id: makeId(),
        nameRef: '=Sheet1!$B$1',
        categoriesRef: '=Sheet1!$A$2:$A$4',
        valuesRef: '=Sheet1!$B$2:$B$4',
        showConnectorLines: true,
        pointOverrides: [{ index: 2, isTotal: true, fillRgb: '#4472C4' }],
        trendlines: [{ id: makeId(), type: 'linear', displayEquation: true, displayRSquared: true }],
        errorBars: { direction: 'both', type: 'percentage', value: 5, cap: true },
      }],
    }, makeId));

    const saved = await saveNativePackage(session.active);
    const loaded = await loadNativePackage(saved);
    expect(sameAuthoredWorkbook(loaded, session.active)).toBe(true);
    expect(loaded.charts).toHaveLength(1);
    expect(loaded.charts![0]!.series[0]!.pointOverrides).toEqual([{ index: 2, isTotal: true, fillRgb: '#4472C4' }]);
    expect(loaded.charts![0]!.series[0]!.trendlines).toHaveLength(1);
    expect(loaded.charts![0]!.series[0]!.errorBars).toEqual({ direction: 'both', type: 'percentage', value: 5, cap: true });
  });
});
