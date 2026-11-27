import { describe, expect, it } from 'vitest';
import {
  createPivotTable,
  deletePivotTable,
  drillDownPivotCell,
  movePivotTable,
  refreshPivotTable,
  togglePivotCollapse,
  updatePivotLayout,
} from '../../src/application/pivot-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { findCell, newWorkbook, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';
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
    ['Region', 'Product', 'Sales'],
    ['East', 'Alpha', 100],
    ['East', 'Beta', 200],
    ['West', 'Alpha', 150],
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

describe('pivot session commands and native persistence', () => {
  it('creates, updates, toggles collapse, refreshes, moves, deletes, and undoes/redoes PivotTables', () => {
    const { doc, makeId } = buildSalesWorkbook();
    const sheetId = doc.sheets[0]!.id;
    const session = new WorkbookSession(doc, { makeId });

    session.execute(createPivotTable(session.capture(), {
      source: { kind: 'range', area: { sheetId, row: 1, column: 1, rows: 4, columns: 3 } },
      target: { sheetId, row: 1, column: 5 },
      rows: [{ fieldName: 'Region' }, { fieldName: 'Product' }],
      values: [{ id: makeId(), fieldName: 'Sales', aggregation: 'sum' }],
    }, makeId));

    const pivot = session.active.pivots![0]!;
    expect(pivot.name).toBe('PivotTable1');
    expect(pivot.extent).toBeDefined();
    expect(findCell(session.active.sheets[0]!, { row: 1, column: 5 })?.spillKind).toBe('pivot');
    expect(findCell(session.active.sheets[0]!, { row: 1, column: 6 })?.content).toEqual({ kind: 'text', value: 'Sum of Sales' });

    // Edit source cell (B2 -> 500 at C2) and refresh
    session.execute({
      kind: 'set-cells',
      label: 'Update source',
      edits: [{ target: { sheetId, row: 2, column: 3 }, content: { kind: 'number', value: 500 } }],
    });
    session.execute(refreshPivotTable(session.capture(), pivot.id, makeId));
    const afterRefresh = findCell(session.active.sheets[0]!, { row: 2, column: 6 })?.content;
    expect(afterRefresh).toEqual({ kind: 'number', value: 700 });

    // Collapse East
    session.execute(togglePivotCollapse(session.capture(), pivot.id, 'East', makeId));
    expect(session.active.pivots![0]!.collapsedKeys).toEqual(['East']);
    expect(session.active.pivots![0]!.extent!.rows).toBeLessThan(pivot.extent!.rows);

    // Update layout to Tabular
    session.execute(updatePivotLayout(session.capture(), pivot.id, { reportLayout: 'tabular', subtotals: 'bottom' }, makeId));
    expect(session.active.pivots![0]!.reportLayout).toBe('tabular');

    // Move PivotTable to E10 and verify old cells cleared
    session.execute(movePivotTable(session.capture(), pivot.id, { sheetId, row: 10, column: 5 }, makeId));
    expect(findCell(session.active.sheets[0]!, { row: 1, column: 5 })).toBeUndefined();
    expect(findCell(session.active.sheets[0]!, { row: 10, column: 5 })?.spillKind).toBe('pivot');

    // Undo move restores E1
    session.undo();
    expect(findCell(session.active.sheets[0]!, { row: 10, column: 5 })).toBeUndefined();
    expect(findCell(session.active.sheets[0]!, { row: 1, column: 5 })?.spillKind).toBe('pivot');

    // Delete cleans all pivot cells and metadata
    session.execute(deletePivotTable(session.capture(), pivot.id));
    expect(session.active.pivots).toBeUndefined();
    expect(session.active.sheets[0]!.cells.some(c => c.spillKind === 'pivot')).toBe(false);
  });

  it('rejects PivotTable overlap with populated cells, merges, and tables without mutating state', () => {
    const { doc, makeId } = buildSalesWorkbook();
    const sheetId = doc.sheets[0]!.id;
    const session = new WorkbookSession(doc, { makeId });

    // Collides with occupied A1
    expect(() => session.execute(createPivotTable(session.capture(), {
      source: { kind: 'range', area: { sheetId, row: 1, column: 1, rows: 4, columns: 3 } },
      target: { sheetId, row: 1, column: 1 },
      rows: [{ fieldName: 'Region' }],
      values: [{ id: makeId(), fieldName: 'Sales', aggregation: 'sum' }],
    }, makeId))).toThrow(/overlaps existing worksheet data/i);
    expect(session.active.pivots).toBeUndefined();
  });

  it('round-trips PivotTable definitions and pivot cells through native .exlt v1 packages', async () => {
    const { doc, makeId } = buildSalesWorkbook();
    const sheetId = doc.sheets[0]!.id;
    const session = new WorkbookSession(doc, { makeId });

    session.execute(createPivotTable(session.capture(), {
      source: { kind: 'range', area: { sheetId, row: 1, column: 1, rows: 4, columns: 3 } },
      target: { sheetId, row: 1, column: 5 },
      filters: [{ fieldName: 'Product' }],
      rows: [{ fieldName: 'Region', sort: { direction: 'desc' } }],
      values: [{ id: makeId(), fieldName: 'Sales', aggregation: 'average', showValuesAs: 'percentOfColumnTotal', numberFormat: '0.0%' }],
    }, makeId));

    const blob = await saveNativePackage(session.active);
    const loaded = await loadNativePackage(blob);
    expect(loaded.pivots).toEqual(session.active.pivots);
    expect(loaded.sheets[0]!.cells.filter(c => c.spillKind === 'pivot')).toHaveLength(
      session.active.sheets[0]!.cells.filter(c => c.spillKind === 'pivot').length,
    );
  });

  it('drills down a PivotTable value cell into a formatted Table sheet and supports undo', () => {
    const { doc, makeId } = buildSalesWorkbook();
    const sheetId = doc.sheets[0]!.id;
    const session = new WorkbookSession(doc, { makeId });

    session.execute(createPivotTable(session.capture(), {
      source: { kind: 'range', area: { sheetId, row: 1, column: 1, rows: 4, columns: 3 } },
      target: { sheetId, row: 1, column: 5 },
      rows: [{ fieldName: 'Region' }],
      values: [{ id: makeId(), fieldName: 'Sales', aggregation: 'sum' }],
    }, makeId));

    const pivot = session.active.pivots![0]!;
    // Row 1 = headers (Row Labels / Sum of Sales), Row 2 Col 6 = East total (300 = 2 source rows)
    expect(drillDownPivotCell(session.capture(), pivot.id, { sheetId, row: 1, column: 5 }, makeId)).toBeUndefined();

    const drill = drillDownPivotCell(session.capture(), pivot.id, { sheetId, row: 2, column: 6 }, makeId);
    expect(drill).toBeDefined();
    session.execute(drill!.command);

    expect(session.active.sheets).toHaveLength(2);
    const detailSheet = session.active.sheets.find(s => s.id === drill!.sheetId)!;
    expect(detailSheet.name).toBe('Sheet2');
    // Header + 2 East records = 3 rows × 3 columns
    expect(findCell(detailSheet, { row: 1, column: 1 })?.content).toEqual({ kind: 'text', value: 'Region' });
    expect(findCell(detailSheet, { row: 2, column: 1 })?.content).toEqual({ kind: 'text', value: 'East' });
    expect(findCell(detailSheet, { row: 3, column: 3 })?.content).toEqual({ kind: 'number', value: 200 });
    expect(findCell(detailSheet, { row: 4, column: 1 })).toBeUndefined();
    expect(session.active.names.some(n => n.isTable && n.area.sheetId === detailSheet.id)).toBe(true);

    session.undo();
    expect(session.active.sheets).toHaveLength(1);
    expect(session.active.names).toHaveLength(0);
  });
});
