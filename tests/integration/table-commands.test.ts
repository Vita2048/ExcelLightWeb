import { describe, expect, it } from 'vitest';
import { clear, pasteText } from '../../src/application/editing-commands';
import { changeSheet } from '../../src/application/structure-commands';
import {
  changeTableDesign, createTable, defineName, deleteName, deleteTableSelection, nextTableName, renameTable, resizeTable,
} from '../../src/application/table-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { calculateNative } from '../../src/calculation/native-calculator';
import { effectiveCellStyle } from '../../src/grid/cell-appearance';
import { findCell, newWorkbook, type CellDocument, type WorkbookDocument, type WorkbookName } from '../../src/core/workbook';

function ids() {
  let n = 0;
  return () => 'id-' + ++n;
}

function cell(row: number, column: number, content: CellDocument['content'], style?: CellDocument['style']): CellDocument {
  return { id: 'c-' + row + '-' + column, row, column, content, ...(style && { style }) };
}

function book(cells: readonly CellDocument[] = [], names: WorkbookDocument['names'] = []): {
  session: WorkbookSession; makeId: () => string; sheetId: string;
} {
  const makeId = ids();
  const document = newWorkbook(makeId);
  const sheet = document.sheets[0]!;
  const session = new WorkbookSession({ ...document, sheets: [{ ...sheet, cells }], names }, { makeId });
  return { session, makeId, sheetId: sheet.id };
}

function table(sheetId: string, name: string, row: number, column: number, rows: number, columns: number, id = name): WorkbookName {
  return { id, name, area: { sheetId, row, column, rows, columns }, isTable: true };
}

function value(session: WorkbookSession, row: number, column: number) {
  return findCell(session.active.sheets[0]!, { row, column })?.content;
}

describe('workbook names and tables', () => {
  it('rejects reserved identifiers and case-insensitive collisions without changing the workbook', () => {
    const { session, sheetId } = book([cell(1, 1, { kind: 'number', value: 5 })]);
    const define = (name: string) => defineName(session.capture(), {
      id: 'n-' + name, name, area: { sheetId, row: 1, column: 1, rows: 1, columns: 1 },
    });
    for (const name of ['A1', 'R', 'C', 'R1C1', 'r12c3']) expect(() => session.execute(define(name))).toThrow(/Invalid workbook name/);
    session.execute(define('TaxRate'));
    expect(() => session.execute(define('taxrate'))).toThrow(/Duplicate workbook name/);
    expect(session.active.names.map(entry => entry.name)).toEqual(['TaxRate']);
    expect(session.undoCount).toBe(1);
  });

  it('resolves a named cell and gives a copied sheet its own name', () => {
    const { session, makeId, sheetId } = book([
      cell(1, 1, { kind: 'number', value: 5 }),
      cell(2, 1, { kind: 'formula', source: '=TaxRate' }),
    ]);
    session.execute(defineName(session.capture(), {
      id: 'rate', name: 'TaxRate', comment: 'local', area: { sheetId, row: 1, column: 1, rows: 1, columns: 1 },
    }));
    expect(calculateNative(session.active).updates.find(update => update.target.row === 2)?.cache).toEqual({ kind: 'number', value: 5 });
    const before = session.active;
    const copied = changeSheet(session.capture(), sheetId, { kind: 'place', before: '', copy: true }, makeId);
    session.execute(copied.command);
    const clone = session.active.names.find(entry => entry.area.sheetId === copied.active);
    expect(clone?.name).toBe('TaxRate_2');
    expect(clone?.id).not.toBe('rate');
    expect(session.active.names.find(entry => entry.id === 'rate')).toEqual(before.names[0]);
    expect(session.active.sheets.find(sheet => sheet.id === copied.active)?.cells.some(item =>
      item.content.kind === 'formula' && item.content.source === '=TaxRate')).toBe(true);
  });

  it('creates explicit tables, normalizes headers and shifts formulas when headers are added', () => {
    const empty = book();
    const wide = table(empty.sheetId, 'EmptyTable', 1, 1, 11, 6);
    empty.session.execute(createTable(empty.session.capture(), wide, false, empty.makeId));
    expect(empty.session.active.names[0]?.area.rows).toBe(11);
    expect([1, 2, 3, 4, 5, 6].map(column => value(empty.session, 1, column))).toEqual(
      [1, 2, 3, 4, 5, 6].map(column => ({ kind: 'text', value: 'Column' + column })));
    expect(empty.session.active.sheets[0]?.filter).toMatchObject({ row: 1, column: 1, rows: 11, columns: 6 });
    expect(() => empty.session.execute(createTable(empty.session.capture(), table(empty.sheetId, 'Overlap', 10, 6, 2, 2), false, empty.makeId)))
      .toThrow(/overlap/i);
    expect(empty.session.undo() && empty.session.active.names).toEqual([]);

    const populated = book([
      cell(1, 1, { kind: 'text', value: 'Tea' }), cell(1, 2, { kind: 'text', value: '10' }),
      cell(2, 1, { kind: 'text', value: 'Coffee' }), cell(2, 2, { kind: 'number', value: 20 }),
      cell(3, 1, { kind: 'formula', source: '=A1' }),
    ]);
    populated.session.execute(createTable(populated.session.capture(), table(populated.sheetId, 'DataTable', 1, 1, 2, 2), false, populated.makeId));
    expect(populated.session.active.names[0]?.area.rows).toBe(3);
    expect(value(populated.session, 2, 1)).toEqual({ kind: 'text', value: 'Tea' });
    expect(value(populated.session, 3, 2)).toEqual({ kind: 'number', value: 20 });
    expect(value(populated.session, 4, 1)).toEqual({ kind: 'formula', source: '=A2' });

    const labels = book([
      cell(1, 1, { kind: 'text', value: 'Name' }), cell(1, 2, { kind: 'text', value: 'Name' }),
      cell(1, 4, { kind: 'number', value: 12 }),
    ]);
    labels.session.execute(createTable(labels.session.capture(), table(labels.sheetId, 'HeaderTable', 1, 1, 3, 4), true, labels.makeId));
    expect([1, 2, 3, 4].map(column => value(labels.session, 1, column))).toEqual([
      { kind: 'text', value: 'Name' }, { kind: 'text', value: 'Name2' },
      { kind: 'text', value: 'Column3' }, { kind: 'text', value: '12' },
    ]);

    const single = book();
    single.session.execute(createTable(single.session.capture(), table(single.sheetId, 'SingleTable', 1, 1, 1, 1), false, single.makeId));
    expect(single.session.active.names[0]?.area.rows).toBe(2);
    expect(value(single.session, 1, 1)).toEqual({ kind: 'text', value: 'Column1' });
    expect(nextTableName(single.session.active)).toBe('Table_1');
  });

  it('caches hidden headers, refuses a changed restore, and refuses an occupied total row', () => {
    const { session, makeId, sheetId } = book([
      cell(1, 1, { kind: 'text', value: 'Kind' }), cell(1, 2, { kind: 'text', value: 'Amount' }),
      cell(2, 1, { kind: 'text', value: 'a' }), cell(2, 2, { kind: 'number', value: 10 }),
      cell(3, 1, { kind: 'text', value: 'b' }), cell(3, 2, { kind: 'number', value: 20 }),
      cell(4, 1, { kind: 'text', value: 'a' }), cell(4, 2, { kind: 'number', value: 30 }),
    ]);
    const id = 'orders';
    session.execute(createTable(session.capture(), table(sheetId, 'Orders', 1, 1, 4, 2, id), true, makeId));
    const created = session.active;
    session.execute(changeTableDesign(session.capture(), id, { totalRow: true }, makeId));
    expect(value(session, 5, 1)).toEqual({ kind: 'text', value: 'Total' });
    expect(value(session, 5, 2)).toEqual({ kind: 'formula', source: '=SUBTOTAL(109,B2:B4)' });
    expect(session.active.sheets[0]?.filter).toMatchObject({ rows: 4, columns: 2 });
    expect(calculateNative(session.active).updates.find(update => update.target.row === 5 && update.target.column === 2)?.cache)
      .toEqual({ kind: 'number', value: 60 });
    const withTotal = session.active;
    session.execute(changeTableDesign(session.capture(), id, { headerRow: false, style: 'Purple' }, makeId));
    expect(session.active.names[0]?.headerLabels).toEqual(['Kind', 'Amount']);
    expect(session.active.names[0]?.area.row).toBe(2);
    expect(session.active.sheets[0]?.filter).toBeUndefined();
    expect(value(session, 1, 1)?.kind).not.toBe('text');
    const hidden = session.capture();
    session.execute(pasteText(hidden, { sheetId, row: 1, column: 1, rows: 1, columns: 1 }, 'Changed', 'values', 'en-US', makeId));
    expect(() => changeTableDesign(session.capture(), id, { headerRow: true }, makeId)).toThrow(/changed worksheet values/);
    expect(session.active.names[0]?.area.row).toBe(2);
    expect(value(session, 1, 1)).toEqual({ kind: 'text', value: 'Changed' });
    session.undo();
    session.execute(changeTableDesign(session.capture(), id, { headerRow: true }, makeId));
    expect(value(session, 1, 1)).toEqual({ kind: 'text', value: 'Kind' });
    expect(session.active.names[0]?.headerLabels).toBeUndefined();

    const occupied = new WorkbookSession({
      ...created,
      sheets: [{ ...created.sheets[0]!, cells: [...created.sheets[0]!.cells, cell(5, 1, { kind: 'text', value: 'keep' })] }],
    }, { makeId });
    expect(() => changeTableDesign(occupied.capture(), id, { totalRow: true }, makeId)).toThrow(/occupied cells/);
    expect(findCell(occupied.active.sheets[0]!, { row: 5, column: 1 })?.content).toEqual({ kind: 'text', value: 'keep' });
    expect(withTotal.names[0]?.area.rows).toBe(5);
  });

  it('keeps explicit fills over table styles and protects whole-record deletion', () => {
    const styled = cell(1, 1, { kind: 'text', value: 'Kind' }, { fillRgb: 'FFFFFF', textRgb: '112233' });
    const { session, makeId, sheetId } = book([
      styled, cell(1, 2, { kind: 'text', value: 'Amount' }),
      cell(2, 1, { kind: 'text', value: 'a' }), cell(2, 2, { kind: 'number', value: 10 }),
    ]);
    const definition = table(sheetId, 'Orders', 1, 1, 2, 2, 'orders');
    session.execute(createTable(session.capture(), definition, true, makeId));
    const tables = session.active.names;
    expect(effectiveCellStyle(styled.style, { row: 1, column: 1 }, tables)?.fillRgb).toBe('FFFFFF');
    expect(effectiveCellStyle(styled.style, { row: 1, column: 1 }, tables)?.textRgb).toBe('112233');
    expect(effectiveCellStyle(undefined, { row: 1, column: 2 }, tables)?.fillRgb).toBe('4F81BD');
    expect(effectiveCellStyle({ bold: true }, { row: 2, column: 2 }, tables)?.fillRgb).toBe('DCE6F1');
    expect(effectiveCellStyle({ fillDetails: { pattern: 'none' } }, { row: 2, column: 1 }, tables)?.fillRgb).toBeUndefined();

    const basis = session.capture();
    session.execute(clear(basis, [{ sheetId, row: 2, column: 1, rows: 1, columns: 2 }], 'contents'));
    expect(session.active.names).toHaveLength(1);
    session.undo();
    expect(deleteTableSelection(session.capture(), [{ sheetId, row: 2, column: 1, rows: 1, columns: 2 }])).toBeUndefined();
    const removed = deleteTableSelection(session.capture(), [{ sheetId, row: 1, column: 1, rows: 2, columns: 2 }])!;
    session.execute(removed);
    expect(session.active.names).toEqual([]);
    expect(session.active.sheets[0]?.cells.every(item => item.content.kind === 'blank' || item.row > 2)).toBe(true);
    expect(session.active.sheets[0]?.filter).toBeUndefined();
    expect(() => deleteName(basis, 'orders')).toThrow(/entire table/);
    const range = defineName(basis, { id: 'rate', name: 'Rate', area: { sheetId, row: 2, column: 2, rows: 1, columns: 1 } });
    const named = new WorkbookSession(basis.document, { makeId });
    named.execute(range);
    named.execute(deleteName(named.capture(), 'rate'));
    expect(named.active.names.every(entry => entry.id !== 'rate')).toBe(true);
  });

  it('resizes a table without overlapping another and moves its total formulas', () => {
    const { session, makeId, sheetId } = book([
      cell(1, 1, { kind: 'text', value: 'Kind' }), cell(1, 2, { kind: 'text', value: 'Amount' }),
      cell(2, 2, { kind: 'number', value: 10 }),
    ]);
    session.execute(createTable(session.capture(), table(sheetId, 'Orders', 1, 1, 2, 2, 'orders'), true, makeId));
    session.execute(changeTableDesign(session.capture(), 'orders', { totalRow: true }, makeId));
    const current = session.active.names[0]!;
    session.execute(resizeTable(session.capture(), { ...current, area: { ...current.area, rows: 5, columns: 3 } }, makeId));
    expect(value(session, 1, 3)).toEqual({ kind: 'text', value: 'Column3' });
    expect(value(session, 5, 1)).toEqual({ kind: 'text', value: 'Total' });
    expect(value(session, 5, 2)).toEqual({ kind: 'formula', source: '=SUBTOTAL(109,B2:B4)' });
    expect(session.active.sheets[0]?.filter?.columns).toBe(3);
    session.execute(createTable(session.capture(), table(sheetId, 'Right', 1, 5, 2, 1, 'right'), false, makeId));
    expect(() => resizeTable(session.capture(), { ...session.active.names[0]!, area: { ...current.area, columns: 5 } }, makeId))
      .toThrow(/overlap/i);
  });

  it('shifts rows below a shrinking total row instead of deleting them', () => {
    const { session, makeId, sheetId } = book([
      cell(1, 1, { kind: 'text', value: 'Kind' }), cell(1, 2, { kind: 'text', value: 'Qty' }), cell(1, 3, { kind: 'text', value: 'Amount' }),
      cell(2, 1, { kind: 'text', value: 'a' }), cell(2, 2, { kind: 'number', value: 1 }), cell(2, 3, { kind: 'number', value: 10 }),
      cell(3, 1, { kind: 'text', value: 'b' }), cell(3, 2, { kind: 'formula', source: '=A3' }), cell(3, 3, { kind: 'number', value: 4 }),
      cell(4, 1, { kind: 'text', value: 'c' }), cell(4, 2, { kind: 'number', value: 3 }), cell(4, 3, { kind: 'number', value: 7 }),
      cell(3, 4, { kind: 'text', value: 'keep' }),
      cell(6, 1, { kind: 'text', value: 'below' }),
    ]);
    session.execute(createTable(session.capture(), table(sheetId, 'Orders', 1, 1, 4, 3, 'orders'), true, makeId));
    session.execute(changeTableDesign(session.capture(), 'orders', { totalRow: true }, makeId));
    const current = session.active.names[0]!;
    expect(current.area.rows).toBe(5);
    expect(() => resizeTable(session.capture(), { ...current, area: { ...current.area, rows: 3, columns: 4 } }, makeId))
      .toThrow(/occupied cells/);
    expect(value(session, 3, 1)).toEqual({ kind: 'text', value: 'b' });

    session.execute(resizeTable(session.capture(), { ...current, area: { ...current.area, rows: 3 } }, makeId));
    expect(session.active.names[0]?.area.rows).toBe(3);
    expect(value(session, 2, 1)).toEqual({ kind: 'text', value: 'a' });
    expect(value(session, 3, 1)).toEqual({ kind: 'text', value: 'Total' });
    expect(value(session, 3, 2)).toBeUndefined();
    expect(value(session, 3, 3)).toEqual({ kind: 'formula', source: '=SUBTOTAL(109,C2:C2)' });
    expect(value(session, 4, 1)).toEqual({ kind: 'text', value: 'b' });
    expect(value(session, 4, 2)).toEqual({ kind: 'formula', source: '=A4' });
    expect(value(session, 4, 3)).toEqual({ kind: 'number', value: 4 });
    expect(value(session, 5, 1)).toEqual({ kind: 'text', value: 'c' });
    expect(value(session, 5, 3)).toEqual({ kind: 'number', value: 7 });
    expect(value(session, 3, 4)).toEqual({ kind: 'text', value: 'keep' });
    expect(value(session, 6, 1)).toEqual({ kind: 'text', value: 'below' });
  });

  it('pulls occupied rows below a total into the table when the table grows', () => {
    const { session, makeId, sheetId } = book([
      cell(1, 1, { kind: 'text', value: 'Kind' }), cell(1, 2, { kind: 'text', value: 'Amount' }),
      cell(2, 1, { kind: 'text', value: 'a' }), cell(2, 2, { kind: 'number', value: 10 }),
      cell(4, 1, { kind: 'text', value: 'next' }), cell(4, 2, { kind: 'formula', source: '=A4' }),
      cell(5, 1, { kind: 'text', value: 'last' }), cell(5, 2, { kind: 'number', value: 8 }),
    ]);
    session.execute(createTable(session.capture(), table(sheetId, 'Orders', 1, 1, 2, 2, 'orders'), true, makeId));
    session.execute(changeTableDesign(session.capture(), 'orders', { totalRow: true }, makeId));
    const current = session.active.names[0]!;
    session.execute(resizeTable(session.capture(), { ...current, area: { ...current.area, rows: 5 } }, makeId));
    expect(session.active.names[0]?.area.rows).toBe(5);
    expect(value(session, 2, 2)).toEqual({ kind: 'number', value: 10 });
    expect(value(session, 3, 1)).toEqual({ kind: 'text', value: 'next' });
    expect(value(session, 3, 2)).toEqual({ kind: 'formula', source: '=A3' });
    expect(value(session, 4, 1)).toEqual({ kind: 'text', value: 'last' });
    expect(value(session, 4, 2)).toEqual({ kind: 'number', value: 8 });
    expect(value(session, 5, 1)).toEqual({ kind: 'text', value: 'Total' });
    expect(value(session, 5, 2)).toEqual({ kind: 'formula', source: '=SUBTOTAL(109,B2:B4)' });
  });

  it('rewrites native structured table formulas when a table is renamed', () => {
    const { session, makeId, sheetId } = book([
      cell(1, 1, { kind: 'text', value: 'Item' }), cell(1, 2, { kind: 'text', value: 'Amount' }),
      cell(2, 1, { kind: 'text', value: 'Pen' }), cell(2, 2, { kind: 'number', value: 12 }),
      cell(3, 1, { kind: 'text', value: 'Pad' }), cell(3, 2, { kind: 'number', value: 18 }),
      cell(1, 4, { kind: 'formula', source: '=SUM(Orders[Amount])' }),
    ]);
    session.execute(createTable(session.capture(), table(sheetId, 'Orders', 1, 1, 3, 2, 'orders'), true, makeId));
    session.execute(renameTable(session.capture(), 'orders', 'Invoices'));
    expect(value(session, 1, 4)).toEqual({ kind: 'formula', source: '=SUM(Invoices[Amount])' });
    expect(calculateNative(session.active).updates.find(u => u.target.row === 1 && u.target.column === 4)?.cache)
      .toEqual({ kind: 'number', value: 30 });
  });

});
