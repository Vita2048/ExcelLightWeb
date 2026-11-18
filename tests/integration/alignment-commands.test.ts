import { expect, it } from 'vitest';
import { changeAlignment, mergeState, proposeMerge, type MergeMode } from '../../src/application/alignment-commands';
import { captureClipboard, fill, pasteBlock } from '../../src/application/editing-commands';
import { editAxis } from '../../src/application/structure-commands';
import { formatStyles } from '../../src/application/style-commands';
import { WorkbookSession, type CellEdit } from '../../src/application/workbook-session';
import { emptySheet, findCell, type WorkbookName } from '../../src/core/workbook';

const extent = { row: 2, column: 2, rows: 2, columns: 2 };
const range = { ...extent, sheetId: 'sheet' };
const at = (row: number, column: number) => ({ row, column, sheetId: 'sheet' });
function setup(names: WorkbookName[] = []) {
  let id = 0;
  const makeId = () => 'id-' + ++id;
  const session = new WorkbookSession({ id: 'book', formatVersion: 1, dateSystem: '1900', names,
    sheets: [emptySheet('sheet', 'Sheet')] }, { makeId });
  const set = (...edits: CellEdit[]) => session.execute({ kind: 'set-cells', edits });
  const sheet = () => session.active.sheets[0]!;
  const cell = (row: number, column: number) => findCell(sheet(), { row, column });
  const propose = (mode: MergeMode, target = range) => proposeMerge(session.capture(), target, mode, makeId);
  return { session, set, sheet, cell, propose };
}

it('moves the first populated value to a stable anchor and combines all edited formatting with one undo', () => {
  const { session, set, cell, sheet, propose } = setup();
  set({ target: at(2, 2), content: { kind: 'blank' }, style: { italic: true }, preserveStyle: false },
    { target: at(2, 3), content: { kind: 'number', value: 7 }, style: { numberFormat: '0.00' }, preserveStyle: false },
    { target: at(3, 2), content: { kind: 'text', value: 'discarded' } });
  const before = session.active;
  const basis = session.capture();
  const proposal = propose('cells');
  expect(proposal.discards).toBe(true);
  expect(session.active).toBe(before);
  const edit = formatStyles({ ...basis, document: proposal.command.document }, [range], { bold: true }, undefined, { top: { style: 'thin' } },
    { alignment: 'right', textRotation: -45, indent: 3, readingOrder: 'rtl', verticalAlignment: 'top' });
  expect(session.tryExecute(basis, { ...proposal.command, edits: edit.edits })).toBe(true);
  expect(mergeState(sheet(), extent)).toBe(true);
  expect(cell(2, 2)).toMatchObject({ id: before.sheets[0]!.cells[0]!.id, content: { value: 7 },
    style: { bold: true, italic: true, numberFormat: '0.00', alignment: 'right', textRotation: -45, indent: 3,
      readingOrder: 'rtl', verticalAlignment: 'top', borderDetails: { top: { style: 'thin' } } } });
  expect(cell(2, 3)?.content.kind).toBe('blank');
  expect(cell(3, 2)?.content.kind).toBe('blank');
  expect(session.undo()).toBe(true);
  expect(session.active).toEqual(before);
  expect(session.redo()).toBe(true);
  session.execute(propose('unmerge').command);
  expect(sheet().merges).toEqual([]);
  expect(cell(2, 2)?.content).toEqual({ kind: 'number', value: 7 });
  expect(cell(3, 2)?.content.kind).toBe('blank');
});

it('merges across independently, replaces intersecting merges after expanding, and copies and shifts merged areas', () => {
  const { session, set, cell, sheet, propose } = setup();
  set({ target: at(2, 3), content: { kind: 'text', value: 'first row' } },
    { target: at(3, 3), content: { kind: 'text', value: 'next row' } });
  expect(propose('across').discards).toBe(false);
  session.execute(propose('across').command);
  expect(sheet().merges).toEqual([{ ...extent, rows: 1 }, { ...extent, row: 3, rows: 1 }]);
  expect(mergeState(sheet(), extent)).toBeUndefined();
  expect(cell(3, 2)?.content).toEqual({ kind: 'text', value: 'next row' });
  const clipboard = captureClipboard(session.capture(), range, false);
  session.execute(pasteBlock(session.capture(), { ...range, row: 6 }, clipboard, 'all'));
  expect(sheet().merges).toContainEqual({ ...extent, row: 7, rows: 1 });
  expect(() => fill(session.capture(), range, 'down')).toThrow(/Unmerge/);
  expect(() => editAxis(session.capture(), 'sheet', { axis: 'row', start: 2, count: 1, delete: false, lane: [3, 3] })).toThrow(/Unmerge/);
  session.execute(editAxis(session.capture(), 'sheet', { axis: 'row', start: 2, count: 1, delete: false }));
  expect(sheet().merges).toContainEqual({ ...extent, row: 8, rows: 1 });
  const overlap = propose('center', { ...range, row: 3, column: 3, columns: 1 });
  expect(overlap.discards).toBe(true);
  expect(overlap.range).toMatchObject({ row: 3, column: 2, rows: 2, columns: 2 });
  session.execute(overlap.command);
  expect(cell(3, 2)?.style?.alignment).toBe('center');
  expect(sheet().merges[0]).toEqual({ ...extent, row: 3 });
  expect(() => set({ target: at(4, 3), content: { kind: 'text', value: 'covered' } })).toThrow(/merged range/);
});

it('rejects table intersections and Python cells before building a merge, and applies ribbon alignment transitions', () => {
  const table = setup([{ id: 'table', name: 'Inventory', area: range, isTable: true }]);
  expect(() => table.propose('cells')).toThrow(/table/);
  const python = setup();
  python.set({ target: at(3, 3), content: { kind: 'python', blockId: 'block', source: '1' } });
  expect(() => python.propose('center')).toThrow(/Python/);
  const { session, cell } = setup();
  session.execute(changeAlignment(session.capture(), [range], { kind: 'indent', value: 1 }));
  expect(cell(2, 2)?.style).toMatchObject({ alignment: 'left', indent: 1 });
  session.execute(changeAlignment(session.capture(), [range], { kind: 'horizontal', value: 'left' }, cell(2, 2)?.style));
  expect(cell(2, 2)?.style).toBeUndefined();
  session.execute(changeAlignment(session.capture(), [range], { kind: 'rotation', value: 255 }));
  expect(cell(2, 2)?.style?.textRotation).toBe(255);
});
