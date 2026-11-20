import { describe, expect, it } from 'vitest';
import {
  columnValues, compareSortKeys, currentRegion, filterKey, filterRange, filteredRows, headerLabel, localeTextOrder,
  looksLikeHeader, recordSummary, setFilter, sortData, sortKeyFor,
} from '../../src/application/data-commands';
import { changeDimensions } from '../../src/application/structure-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { createTable, changeTableDesign } from '../../src/application/table-commands';
import { calculateNative } from '../../src/calculation/native-calculator';
import { changeSheet } from '../../src/application/structure-commands';
import { findCell, newWorkbook, type CellDocument, type SheetFilter, type WorkbookDocument, type WorkbookName } from '../../src/core/workbook';

const order = localeTextOrder('en-US');

function ids() {
  let n = 0;
  return () => 'id-' + ++n;
}

function cell(row: number, column: number, content: CellDocument['content'], style?: CellDocument['style'],
  cache?: CellDocument['cache']): CellDocument {
  return { id: 'c-' + row + '-' + column, row, column, content, ...(style && { style }), ...(cache && { cache }) };
}

function book(cells: readonly CellDocument[] = [], names: WorkbookDocument['names'] = [],
  rows: WorkbookDocument['sheets'][number]['rows'] = []) {
  const makeId = ids();
  const document = newWorkbook(makeId);
  const sheet = document.sheets[0]!;
  const session = new WorkbookSession({ ...document, sheets: [{ ...sheet, cells, rows }], names }, { makeId });
  return { session, makeId, sheetId: sheet.id };
}

function range(sheetId: string, row: number, column: number, rows: number, columns: number) {
  return { sheetId, row, column, rows, columns };
}

function text(document: WorkbookDocument, row: number, column: number) {
  const sheet = document.sheets[0]!;
  return findCell(sheet, { row, column })?.content;
}

describe('single-key row sort', () => {
  it('orders number, text, boolean, error and blank groups and keeps blanks last in both directions', () => {
    const { session, sheetId } = book([
      cell(1, 1, { kind: 'number', value: 10 }),
      cell(2, 1, { kind: 'formula', source: '=1/0' }, undefined, { kind: 'error', value: '#DIV/0!' }),
      cell(3, 1, { kind: 'boolean', value: true }),
      cell(4, 1, { kind: 'text', value: 'm' }),
      cell(6, 1, { kind: 'number', value: 2 }),
      cell(7, 1, { kind: 'boolean', value: false }),
    ]);
    const target = range(sheetId, 1, 1, 7, 1);
    session.execute(sortData(session.capture(), target, 1, false, true, false, order));
    expect(Array.from({ length: 7 }, (_, index) => text(session.active, index + 1, 1)))
      .toEqual([{ kind: 'number', value: 2 }, { kind: 'number', value: 10 }, { kind: 'text', value: 'm' },
        { kind: 'boolean', value: false }, { kind: 'boolean', value: true },
        { kind: 'formula', source: '=1/0' }, undefined]);
    expect(findCell(session.active.sheets[0]!, { row: 6, column: 1 })?.cache).toEqual({ kind: 'error', value: '#DIV/0!' });

    session.execute(sortData(session.capture(), target, 1, false, false, false, order));
    expect(Array.from({ length: 7 }, (_, index) => text(session.active, index + 1, 1)))
      .toEqual([{ kind: 'formula', source: '=1/0' }, { kind: 'boolean', value: true }, { kind: 'boolean', value: false },
        { kind: 'text', value: 'm' }, { kind: 'number', value: 10 }, { kind: 'number', value: 2 }, undefined]);
  });

  it('keeps equal keys stable when case is ignored and follows the display culture when it is not', () => {
    const cases = (caseSensitive: boolean, ascending: boolean) => {
      const { session, sheetId } = book([
        cell(1, 1, { kind: 'text', value: 'b' }), cell(1, 2, { kind: 'text', value: 'first' }),
        cell(2, 1, { kind: 'text', value: 'B' }), cell(2, 2, { kind: 'text', value: 'second' }),
        cell(3, 1, { kind: 'text', value: 'a' }), cell(3, 2, { kind: 'text', value: 'third' }),
      ]);
      session.execute(sortData(session.capture(), range(sheetId, 1, 1, 3, 2), 1, false, ascending, caseSensitive, order));
      return Array.from({ length: 3 }, (_, index) => text(session.active, index + 1, 2));
    };
    expect(cases(false, true)).toEqual([
      { kind: 'text', value: 'third' }, { kind: 'text', value: 'first' }, { kind: 'text', value: 'second' },
    ]);
    const compare = new Intl.Collator('en-US', { sensitivity: 'variant' }).compare('B', 'b');
    expect(cases(true, true)).toEqual(compare < 0 ? [
      { kind: 'text', value: 'third' }, { kind: 'text', value: 'second' }, { kind: 'text', value: 'first' },
    ] : [
      { kind: 'text', value: 'third' }, { kind: 'text', value: 'first' }, { kind: 'text', value: 'second' },
    ]);
    expect(cases(true, false)).toEqual(compare < 0 ? [
      { kind: 'text', value: 'first' }, { kind: 'text', value: 'second' }, { kind: 'text', value: 'third' },
    ] : [
      { kind: 'text', value: 'second' }, { kind: 'text', value: 'first' }, { kind: 'text', value: 'third' },
    ]);
  });

  it('moves records with their identities and relocates only the formulas that change row', () => {
    const style = { bold: true, fillRgb: 'FFEE00' } as const;
    const { session, sheetId } = book([
      cell(1, 1, { kind: 'number', value: 2 }),
      { ...cell(1, 2, { kind: 'formula', source: '=A1' }, style), id: 'kept' },
      { ...cell(2, 1, { kind: 'number', value: 1 }), id: 'moved' },
      cell(2, 2, { kind: 'formula', source: '=$A$1' }),
      cell(3, 1, { kind: 'number', value: 3 }),
      cell(3, 2, { kind: 'formula', source: '=A3' }),
      cell(1, 3, { kind: 'formula', source: '=A1' }),
      cell(2, 3, { kind: 'formula', source: '=B2' }),
      cell(4, 1, { kind: 'formula', source: '=A1' }),
    ]);
    session.execute(sortData(session.capture(), range(sheetId, 1, 1, 3, 2), 1, false, true, false, order));
    const sheet = session.active.sheets[0]!;
    expect(findCell(sheet, { row: 1, column: 1 })).toMatchObject({ id: 'moved', content: { kind: 'number', value: 1 } });
    expect(findCell(sheet, { row: 1, column: 2 })?.content).toEqual({ kind: 'formula', source: '=$A$1' });
    expect(findCell(sheet, { row: 2, column: 2 })).toMatchObject({
      id: 'kept', content: { kind: 'formula', source: '=A2' }, style,
    });
    expect(findCell(sheet, { row: 3, column: 2 })?.content).toEqual({ kind: 'formula', source: '=A3' });
    expect(findCell(sheet, { row: 1, column: 3 })?.content).toEqual({ kind: 'formula', source: '=A1' });
    expect(findCell(sheet, { row: 2, column: 3 })?.content).toEqual({ kind: 'formula', source: '=B2' });
    expect(findCell(sheet, { row: 4, column: 1 })?.content).toEqual({ kind: 'formula', source: '=A1' });
  });

  it('rejects merged ranges, table header moves, partial table records and malformed formulas without changing the workbook', () => {
    const { session, sheetId } = book([
      cell(1, 1, { kind: 'text', value: 'Name' }), cell(1, 2, { kind: 'text', value: 'Amount' }),
      cell(2, 1, { kind: 'text', value: 'Tea' }), cell(2, 2, { kind: 'number', value: 10 }),
      cell(3, 1, { kind: 'text', value: 'Coffee' }), cell(3, 2, { kind: 'number', value: 20 }),
    ]);
    const table: WorkbookName = {
      id: 'sales', name: 'SalesTable', isTable: true,
      area: { sheetId, row: 1, column: 1, rows: 3, columns: 2 },
    };
    session.execute(createTable(session.capture(), table, true, ids()));
    const before = session.active;
    const undo = session.undoCount;
    expect(() => session.execute(sortData(session.capture(), range(sheetId, 1, 1, 3, 1), 1, true, true, false, order)))
      .toThrow(/all table columns/);
    expect(() => session.execute(sortData(session.capture(), range(sheetId, 1, 1, 3, 2), 1, false, true, false, order)))
      .toThrow(/header/);
    expect(() => session.execute(sortData(session.capture(), range(sheetId, 1, 1, 3, 1), 1, true, true, false, order)))
      .toThrow(/all table columns/);
    expect(session.active).toBe(before);
    expect(session.undoCount).toBe(undo);

    const merged = book([cell(1, 1, { kind: 'number', value: 2 })]);
    const mergedSheet = merged.session.active.sheets[0]!;
    merged.session.execute({ kind: 'structure', label: 'Merge',
      document: { ...merged.session.active, sheets: [{ ...mergedSheet, merges: [{ row: 1, column: 1, rows: 2, columns: 1 }] }] } });
    expect(() => merged.session.execute(sortData(merged.session.capture(), range(merged.sheetId, 1, 1, 2, 1), 1, false, true, false, order)))
      .toThrow(/Unmerge/);

    const malformed = book([
      cell(1, 1, { kind: 'formula', source: '=SUM(' }), cell(2, 1, { kind: 'number', value: 2 }),
    ]);
    expect(() => malformed.session.execute(sortData(malformed.session.capture(), range(malformed.sheetId, 1, 1, 2, 1), 1, false, true, false, order)))
      .toThrow();
    expect(malformed.session.active.sheets[0]!.cells[0]!.content).toEqual({ kind: 'formula', source: '=SUM(' });
  });

  it('sorts a full table record set while keeping its header, extent and total row out of the sort', () => {
    const { session, makeId, sheetId } = book([
      cell(1, 1, { kind: 'text', value: 'Name' }), cell(1, 2, { kind: 'text', value: 'Amount' }),
      cell(2, 1, { kind: 'text', value: 'Tea' }), cell(2, 2, { kind: 'number', value: 10 }),
      cell(3, 1, { kind: 'text', value: 'Coffee' }), cell(3, 2, { kind: 'number', value: 20 }),
    ]);
    const table: WorkbookName = {
      id: 'sales', name: 'SalesTable', isTable: true,
      area: { sheetId, row: 1, column: 1, rows: 3, columns: 2 },
    };
    session.execute(createTable(session.capture(), table, true, makeId));
    session.execute(changeTableDesign(session.capture(), 'sales', { totalRow: true }, makeId));
    const before = session.active.names;
    session.execute(sortData(session.capture(), range(sheetId, 1, 1, 4, 2), 1, true, false, false, order));
    expect(session.active.names).toEqual(before);
    expect(text(session.active, 2, 1)).toEqual({ kind: 'text', value: 'Tea' });
    expect(text(session.active, 3, 1)).toEqual({ kind: 'text', value: 'Coffee' });
    expect(text(session.active, 4, 1)).toEqual({ kind: 'text', value: 'Total' });
    expect(findCell(session.active.sheets[0]!, { row: 4, column: 2 })?.content)
      .toEqual({ kind: 'formula', source: '=SUBTOTAL(109,B2:B3)' });
  });
});

describe('AutoFilter criteria and visibility', () => {
  function filteredBook() {
    const result = book([
      cell(1, 1, { kind: 'text', value: 'Name' }), cell(1, 2, { kind: 'text', value: 'Dept' }),
      cell(1, 3, { kind: 'text', value: 'Amount' }),
      cell(2, 1, { kind: 'text', value: 'zzz' }), cell(2, 2, { kind: 'text', value: 'Sales' }), cell(2, 3, { kind: 'number', value: 5 }),
      cell(3, 1, { kind: 'text', value: 'aaa' }), cell(3, 2, { kind: 'text', value: 'Ops' }), cell(3, 3, { kind: 'number', value: 10 }),
      cell(4, 1, { kind: 'text', value: 'mmm' }), cell(4, 2, { kind: 'text', value: 'Ops' }), cell(4, 3, { kind: 'number', value: 20 }),
      cell(5, 1, { kind: 'text', value: 'outside' }), cell(5, 3, { kind: 'number', value: 99 }),
    ]);
    return result;
  }

  it('combines criteria, spares the header and counts visible records', () => {
    const { session, sheetId } = filteredBook();
    const filter: SheetFilter = { row: 1, column: 1, rows: 4, columns: 3,
      criteria: [{ column: 2, allowed: ['Ops'] }, { column: 3, allowed: ['10', '20'] }] };
    const basis = session.capture();
    const before = session.active;
    session.execute(setFilter(basis, sheetId, filter));
    expect(filteredRows(session.active.sheets[0]!)).toEqual(new Set([2]));
    expect(recordSummary(session.active.sheets[0]!)).toBe('2 of 3 records');
    expect(session.undoCount).toBe(1);
    expect(session.undo()).toBe(true);
    expect(session.active).toEqual(before);
  });

  it('offers other-column values and retains selected keys that another column hides', () => {
    const { session, sheetId } = filteredBook();
    const filter: SheetFilter = { row: 1, column: 1, rows: 4, columns: 3,
      criteria: [{ column: 2, allowed: ['Ops'] }] };
    session.execute(setFilter(session.capture(), sheetId, filter));
    const sheet = session.active.sheets[0]!;
    expect(columnValues(sheet, sheet.filter!, 3, order).map(value => value.key)).toEqual(['10', '20']);
    expect(columnValues(sheet, sheet.filter!, 1, order).map(value => value.label)).toEqual(['aaa', 'mmm']);
    expect(headerLabel(sheet, 1, 2)).toBe('Dept');
    expect(headerLabel(sheet, 1, 5)).toBe('Column E');
  });

  it('formats filter keys with the stored number format so display values select rows', () => {
    const { session, sheetId } = book([
      cell(1, 1, { kind: 'text', value: 'Amount' }),
      cell(2, 1, { kind: 'number', value: 10 }, { numberFormat: '0.00' }),
      cell(3, 1, { kind: 'number', value: 20 }, { numberFormat: '0.00' }),
    ]);
    session.execute(setFilter(session.capture(), sheetId, {
      row: 1, column: 1, rows: 3, columns: 1, criteria: [{ column: 1, allowed: ['10.00'] }],
    }));
    const sheet = session.active.sheets[0]!;
    expect(filteredRows(sheet)).toEqual(new Set([3]));
    expect(columnValues(sheet, sheet.filter!, 1, order).map(value => value.key)).toEqual(['10.00', '20.00']);
  });

  it('sorts filter-hidden and manually hidden rows, retargets criteria and undoes cells and filter together', () => {
    const { session, sheetId } = filteredBook();
    session.execute(setFilter(session.capture(), sheetId, {
      row: 1, column: 1, rows: 4, columns: 3, criteria: [{ column: 2, allowed: ['Ops'] }],
    }));
    session.execute(changeDimensions(session.capture(), sheetId, 'row', [{ index: 3, hidden: true }]));
    const before = session.active;
    const undo = session.undoCount;
    session.execute(sortData(session.capture(), range(sheetId, 1, 1, 4, 3), 1, true, true, false, order));
    const sheet = session.active.sheets[0]!;
    expect(sheet.filter?.sort).toEqual({ column: 1, ascending: true });
    expect(text(session.active, 2, 1)).toEqual({ kind: 'text', value: 'aaa' });
    expect(text(session.active, 3, 1)).toEqual({ kind: 'text', value: 'mmm' });
    expect(text(session.active, 4, 1)).toEqual({ kind: 'text', value: 'zzz' });
    expect(sheet.rows.find(row => row.index === 3)?.hidden).toBe(true);
    expect(filteredRows(sheet)).toEqual(new Set([4]));
    expect(session.undoCount).toBe(undo + 1);
    expect(session.undo()).toBe(true);
    expect(session.active).toEqual(before);
  });

  it('clears criteria without unhiding manual rows and leaves the range in place', () => {
    const { session, sheetId } = filteredBook();
    session.execute(setFilter(session.capture(), sheetId, {
      row: 1, column: 1, rows: 4, columns: 3, criteria: [{ column: 2, allowed: ['Ops'] }],
    }));
    session.execute(changeDimensions(session.capture(), sheetId, 'row', [{ index: 3, hidden: true }]));
    const filter = session.active.sheets[0]!.filter!;
    session.execute(setFilter(session.capture(), sheetId, { ...filter, criteria: [] }));
    const sheet = session.active.sheets[0]!;
    expect(sheet.filter).toMatchObject({ row: 1, column: 1, rows: 4, columns: 3, criteria: [] });
    expect(filteredRows(sheet).size).toBe(0);
    expect(sheet.rows.find(row => row.index === 3)?.hidden).toBe(true);
  });

  it('recalculates SUBTOTAL for visible rows when a filter hides records', () => {
    const { session, sheetId } = book([
      cell(1, 1, { kind: 'text', value: 'Name' }), cell(1, 2, { kind: 'text', value: 'Amount' }),
      cell(2, 1, { kind: 'text', value: 'a' }), cell(2, 2, { kind: 'number', value: 10 }),
      cell(3, 1, { kind: 'text', value: 'b' }), cell(3, 2, { kind: 'number', value: 20 }),
      cell(5, 1, { kind: 'formula', source: '=SUBTOTAL(109,B2:B3)' }),
    ]);
    const before = calculateNative(session.active).updates.find(update => update.target.row === 5)?.cache;
    expect(before).toEqual({ kind: 'number', value: 30 });
    session.execute(setFilter(session.capture(), sheetId, {
      row: 1, column: 1, rows: 3, columns: 2, criteria: [{ column: 1, allowed: ['b'] }],
    }));
    const after = calculateNative(session.active).updates.find(update => update.target.row === 5)?.cache;
    expect(after).toEqual({ kind: 'number', value: 20 });
  });
});

describe('range discovery and header detection', () => {
  it('detects a header row from text over numbers and grows a region around one cell', () => {
    const { session, sheetId } = book([
      cell(1, 1, { kind: 'text', value: 'Name' }), cell(1, 2, { kind: 'text', value: 'Score' }),
      cell(2, 1, { kind: 'text', value: 'b' }), cell(2, 2, { kind: 'number', value: 2 }),
      cell(3, 1, { kind: 'text', value: 'a' }), cell(3, 2, { kind: 'number', value: 1 }),
    ]);
    expect(looksLikeHeader(session.capture(), range(sheetId, 1, 1, 3, 2))).toBe(true);
    expect(currentRegion(session.active, { sheetId, row: 2, column: 2, rows: 1, columns: 1 }))
      .toEqual({ sheetId, row: 1, column: 1, rows: 3, columns: 2 });
    expect(filterRange(session.active, { sheetId, row: 2, column: 2, rows: 1, columns: 1 }))
      .toEqual({ sheetId, row: 1, column: 1, rows: 3, columns: 2 });
    expect(looksLikeHeader(session.capture(), range(sheetId, 2, 1, 2, 2))).toBe(false);
  });

  it('uses a table input rectangle for the filter range', () => {
    const { session, makeId, sheetId } = book([
      cell(1, 1, { kind: 'text', value: 'Name' }), cell(1, 2, { kind: 'text', value: 'Amount' }),
      cell(2, 1, { kind: 'text', value: 'Tea' }), cell(2, 2, { kind: 'number', value: 10 }),
    ]);
    const table: WorkbookName = {
      id: 'sales', name: 'SalesTable', isTable: true,
      area: { sheetId, row: 1, column: 1, rows: 2, columns: 2 },
    };
    session.execute(createTable(session.capture(), table, true, makeId));
    session.execute(changeTableDesign(session.capture(), 'sales', { totalRow: true }, makeId));
    expect(filterRange(session.active, { sheetId, row: 1, column: 2, rows: 1, columns: 1 }))
      .toEqual({ sheetId, row: 1, column: 1, rows: 2, columns: 2 });
  });

  it('copies filter ranges and a sort indicator with a copied sheet and drops them with the table', () => {
    const { session, makeId, sheetId } = book([
      cell(1, 1, { kind: 'text', value: 'Name' }),
      cell(2, 1, { kind: 'text', value: 'a' }),
      cell(3, 1, { kind: 'text', value: 'b' }),
    ]);
    session.execute(setFilter(session.capture(), sheetId, {
      row: 1, column: 1, rows: 3, columns: 1, criteria: [], sort: { column: 1, ascending: true },
    }));
    const copied = changeSheet(session.capture(), sheetId, { kind: 'place', before: '', copy: true }, makeId);
    session.execute(copied.command);
    expect(session.active.sheets.find(sheet => sheet.id === copied.active)?.filter).toEqual({
      row: 1, column: 1, rows: 3, columns: 1, criteria: [], sort: { column: 1, ascending: true },
    });
  });
});

describe('sort keys', () => {
  it('reads blanks, cached formula values and display filter keys', () => {
    expect(sortKeyFor(undefined, 3)).toEqual({ kind: 'blank', number: 0, text: '', flag: false, index: 3 });
    expect(sortKeyFor(cell(1, 1, { kind: 'text', value: '' }), 1).kind).toBe('blank');
    expect(sortKeyFor({ ...cell(1, 1, { kind: 'formula', source: '=A1' }), cache: { kind: 'number', value: 4 } }, 1))
      .toMatchObject({ kind: 'number', number: 4 });
    expect(compareSortKeys({ kind: 'text', number: 0, text: 'b', flag: false, index: 1 },
      { kind: 'text', number: 0, text: 'B', flag: false, index: 2 }, true, false, order)).toBeLessThan(0);
    expect(filterKey(cell(1, 1, { kind: 'boolean', value: true }))).toBe('TRUE');
    expect(filterKey(undefined)).toBe('');
  });
});
