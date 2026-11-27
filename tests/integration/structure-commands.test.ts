import { describe, expect, it } from 'vitest';
import { changeDimensions, changeSheet, editAxis } from '../../src/application/structure-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { calculateNative } from '../../src/calculation/native-calculator';
import { shiftReferences } from '../../src/core/formula-references';
import { documentLimits, emptySheet, findCell, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';

function fixture() {
  let id = 0;
  const makeId = () => 'new-' + ++id;
  const cells: CellDocument[] = Array.from({ length: 16 }, (_, i) => {
    const row = Math.floor(i / 4) + 1, column = i % 4 + 1;
    return { row, column, id: 'cell-' + i, style: { bold: true }, content: { kind: 'number', value: row * 10 + column } };
  });
  const formula = (source: string, row: number, column: number): CellDocument =>
    ({ id: 'formula-' + row + '-' + column, row, column, content: { kind: 'formula', source } });
  const document: WorkbookDocument = {
    id: 'book', formatVersion: 1, dateSystem: '1900',
    sheets: [
      { ...emptySheet('data', 'Data'), cells: [...cells, formula('=$B$2', 1, 6)] },
      { ...emptySheet('totals', 'Totals'), cells: [formula('=SUM(Data!A1:C4)', 1, 1),
        formula('=Data!$B$2', 2, 1), formula('=Data!D3', 3, 1)] },
    ],
    names: [{ id: 'named', name: 'SalesRange', area: { sheetId: 'data', row: 1, column: 1, rows: 4, columns: 3 } }],
  };
  const session = new WorkbookSession(document, { makeId });
  const data = (row: number, column: number) => findCell(session.active.sheets.find(sheet => sheet.id === 'data')!, { row, column });
  const totals = () => calculateNative(session.active).updates.filter(update => update.target.sheetId === 'totals').map(update => update.cache?.value);
  return { session, makeId, data, totals };
}
describe('structural workbook commands', () => {
  it.each([
    ['row', false, [2, 2], 4, 2, [324, 22, 34]],
    ['row', true, [2, 2], 2, 2, [270, '#REF!', 34]],
    ['column', false, [2, 3], 2, 4, [324, 22, 34]],
    ['column', true, [2, 3], 2, 2, [214, '#REF!', 34]],
  ] as const)('shifts %s lanes with delete=%s and rewrites dependent ranges in one undo', (axis, deleting, lane, row, column, values) => {
    const { session, data, totals } = fixture();
    const before = session.capture();
    const prior = deleting ? data(axis === 'row' ? 4 : 2, axis === 'row' ? 2 : 4) : data(2, 2);
    expect(session.tryExecute(before, editAxis(before, 'data', { axis, start: 2, count: 2, lane, delete: deleting }))).toBe(true);
    expect(data(row, column)?.id).toBe(prior?.id);
    expect(data(row, column)?.style).toEqual(prior?.style);
    expect(totals()).toEqual(values);
    expect(data(1, 6)?.content).toEqual({ kind: 'formula', source: deleting ? '=#REF!' : axis === 'row' ? '=$B$4' : '=$D$2' });
    expect(session.active.names[0]?.area.invalid).toBe(true);
    expect(session.undoCount).toBe(1);
    expect(session.undo()).toBe(true);
    expect(session.active).toEqual(before.document);
    expect(session.redo()).toBe(true);
    expect(totals()).toEqual(values);
  });

  it('shifts absolute references, preserves whole-column syntax and handles quoted sheet names', () => {
    const insert = { axis: 'row', start: 2, count: 2, delete: false } as const;
    expect(shiftReferences('=SUM(\'O\'\'Brien\'!$A$1:B4)+SUM(A:A)+"A2"', "O'Brien", "O'Brien", insert))
      .toBe('=SUM(\'O\'\'Brien\'!$A$1:B6)+SUM(A:A)+"A2"');
    expect(shiftReferences('=SUM(\'O\'\'Brien\'!$4:$2)+$A$2', 'Totals', "O'Brien", { ...insert, delete: true }))
      .toBe("=SUM('O''Brien'!$2:$2)+$A$2");
  });

  it('moves dimensions, filters, merges and names with an entire axis while retaining stable IDs', () => {
    const { session, data } = fixture();
    const basis = session.capture();
    const sheet = basis.document.sheets[0]!;
    session.execute({ kind: 'structure', label: 'fixture', document: { ...basis.document, sheets: [
      { ...sheet, rows: [{ index: 2, size: 36, hidden: true }], merges: [{ row: 1, column: 8, rows: 2, columns: 1 }],
        filter: { row: 1, column: 1, rows: 4, columns: 3, criteria: [{ column: 2, allowed: ['x'] }] } },
      basis.document.sheets[1]!,
    ] } });
    const id = data(2, 2)?.id;
    session.execute(editAxis(session.capture(), 'data', { axis: 'row', start: 2, count: 1, delete: false }));
    expect(data(3, 2)?.id).toBe(id);
    const shifted = session.active.sheets[0]!;
    expect(shifted.rows).toEqual([{ index: 3, size: 36, hidden: true }]);
    expect(shifted.merges[0]?.rows).toBe(3);
    expect(shifted.filter?.rows).toBe(5);
    expect(session.active.names[0]?.area.rows).toBe(5);
    session.execute(editAxis(session.capture(), 'data', { axis: 'column', start: 2, count: 1, delete: true }));
    expect(session.active.sheets[0]?.filter?.criteria).toEqual([]);
  });

  it('rejects overflow and stale proposals without partial changes', () => {
    const { session } = fixture();
    session.execute({ kind: 'set-cells', edits: [{ target: { sheetId: 'data', row: documentLimits.rows, column: 1 },
      content: { kind: 'text', value: 'edge' } }] });
    const basis = session.capture();
    expect(() => editAxis(basis, 'data', { axis: 'row', start: 1, count: 1, delete: false })).toThrow(/boundary/);
    expect(session.capture()).toBe(basis);
    const pending = changeDimensions(basis, 'data', 'row', [{ index: 2, size: 50 }]);
    session.undo();
    expect(session.tryExecute(basis, pending)).toBe(false);
  });

  it('renames, copies independent records, reorders and deletes with reference and name integrity', () => {
    const { session, makeId, totals } = fixture();
    session.execute({ kind: 'set-cells', edits: [
      { target: { sheetId: 'data', row: 1, column: 7 }, content: { kind: 'formula', source: '=Data!A1' } },
      { target: { sheetId: 'data', row: 1, column: 8 }, content: { kind: 'python', source: '42', blockId: 'block' } },
    ] });
    const original = session.capture();
    const execute = (id: string, change: Parameters<typeof changeSheet>[2]) => {
      const next = changeSheet(session.capture(), id, change, makeId);
      session.execute(next.command);
      return next.active;
    };
    execute('data', { kind: 'rename', name: "O'Brien data" });
    expect(session.active.sheets[1]?.cells[1]?.content).toEqual({ kind: 'formula', source: "='O''Brien data'!$B$2" });
    const copy = execute('data', { kind: 'place', before: '', copy: true });
    expect(session.active.sheets.at(-1)?.id).toBe(copy);
    expect(session.active.sheets.at(-1)?.cells[0]?.id).not.toBe(original.document.sheets[0]?.cells[0]?.id);
    const copiedSheet = session.active.sheets.at(-1)!;
    expect(findCell(copiedSheet, { row: 1, column: 7 })?.content).toEqual({ kind: 'formula', source: "='Sheet1'!A1" });
    expect(findCell(copiedSheet, { row: 1, column: 8 })?.content).toMatchObject({ kind: 'python', source: '42' });
    expect(findCell(copiedSheet, { row: 1, column: 8 })?.content).not.toMatchObject({ blockId: 'block' });
    expect(session.active.names.at(-1)).toMatchObject({ name: 'SalesRange_2', area: { sheetId: copy } });
    expect(session.active.names.at(-1)?.id).not.toBe('named');
    execute(copy, { kind: 'place', before: 'data', copy: false });
    expect(session.active.sheets[0]?.id).toBe(copy);
    execute('data', { kind: 'delete' });
    expect(totals()).toEqual(['#REF!', '#REF!', '#REF!']);
    expect(session.active.names[0]?.area.invalid).toBe(true);
    expect(session.undo()).toBe(true);
    expect(totals()).toEqual([324, 22, 34]);
  });

  it('sizes selected tracks in one transaction and retains those sizes through hide and unhide', () => {
    const { session } = fixture();
    session.execute(changeDimensions(session.capture(), 'data', 'column', [{ index: 1, size: 180 }, { index: 2, size: 180 }]));
    expect(session.undoCount).toBe(1);
    session.execute(changeDimensions(session.capture(), 'data', 'column', [{ index: 2, hidden: true }]));
    session.execute(changeDimensions(session.capture(), 'data', 'column', [{ index: 2, hidden: false }]));
    expect(session.active.sheets[0]?.columns.map(track => track.size)).toEqual([180, 180]);
    const basis = session.capture();
    session.execute(changeDimensions(basis, 'data', 'column',
      Array.from({ length: 26 }, (_, i) => ({ index: i + 1, hidden: true }))));
    expect(session.active.sheets[0]?.columns.filter(track => track.hidden)).toHaveLength(26);
    expect(session.undo()).toBe(true);
    expect(session.active).toEqual(basis.document);
  });

  it('rewrites and shrinks 3D sheet references across rename, axis shift, and delete', () => {
    let id = 0;
    const makeId = () => 's3d-' + ++id;
    const document: WorkbookDocument = {
      id: 'book-3d',
      formatVersion: 1,
      dateSystem: '1900',
      sheets: [
        emptySheet('q1', 'Q1'),
        emptySheet('q2', 'Q2'),
        emptySheet('q3', 'Q3'),
        {
          ...emptySheet('sum', 'Summary'),
          cells: [{ id: 'f1', row: 1, column: 1, content: { kind: 'formula', source: '=SUM(Q1:Q3!B2:B4)' } }],
        },
      ],
      names: [],
    };
    const session = new WorkbookSession(document, { makeId });
    session.execute(changeSheet(session.capture(), 'q1', { kind: 'rename', name: 'Jan' }, makeId).command);
    expect(findCell(session.active.sheets[3]!, { row: 1, column: 1 })?.content)
      .toEqual({ kind: 'formula', source: '=SUM(Jan:Q3!B2:B4)' });
    session.execute(editAxis(session.capture(), 'q2', { axis: 'row', start: 2, count: 1, delete: false }));
    expect(findCell(session.active.sheets[3]!, { row: 1, column: 1 })?.content)
      .toEqual({ kind: 'formula', source: '=SUM(Jan:Q3!B3:B5)' });
    session.execute(changeSheet(session.capture(), 'q1', { kind: 'delete' }, makeId).command);
    expect(findCell(session.active.sheets[2]!, { row: 1, column: 1 })?.content)
      .toEqual({ kind: 'formula', source: '=SUM(Q2:Q3!B3:B5)' });
  });
});
