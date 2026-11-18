import { expect, it } from 'vitest';
import { captureClipboard, clipboardText, editorEntryText, parseEntry, pasteText, type EditRange } from '../../src/application/editing-commands';
import { applyNumberFormat } from '../../src/application/number-format-command';
import { WorkbookSession } from '../../src/application/workbook-session';
import { findCell, newWorkbook } from '../../src/core/workbook';

function fixture() {
  let id = 0;
  const makeId = () => 'id-' + ++id;
  const session = new WorkbookSession(newWorkbook(makeId), { makeId });
  const target = (row: number, column = 1) => ({ sheetId: session.active.sheets[0]!.id, row, column });
  const cell = (row: number, column = 1) => findCell(session.active.sheets[0]!, { row, column });
  const range: EditRange = { ...target(1), rows: 2, columns: 1 };
  return { session, range, target, cell };
}

it('changes only the selected style field, keeps the current formula cache, and undoes overlapping edits in one step', () => {
  const { session, range, target, cell } = fixture();
  session.execute({ kind: 'set-cells', edits: [
    { target: target(1), content: { kind: 'formula', source: '=1.25' }, preserveStyle: false,
      style: { bold: true, numberFormat: '0.00%' } },
    { target: target(2), content: { kind: 'number', value: 1.25 }, preserveStyle: false,
      style: { fontSize: 11, numberFormat: '#,##0.00;[Red](#,##0.00)' } },
  ] });
  const { kind: _kind, document: _document, ...basis } = session.beginNativeCalculation('format-test');
  expect(session.tryCommitNativeCalculation({ ...basis, kind: 'native-calculated',
    updates: [{ target: target(1), cache: { kind: 'number', value: 1.25 } }] })).toBe(true);
  const prior = session.active;
  const revision = session.revision;
  const first = cell(1)!;
  const second = cell(2)!;
  session.execute(applyNumberFormat(session.capture(), [range, { ...range, row: 2 }], { decimals: 1 }));
  expect(session.capture().authoredRevision).toBe(revision + 1);
  expect(cell(1)).toEqual({ ...first, style: { bold: true, numberFormat: '0.000%' } });
  expect(cell(2)).toEqual({ ...second, style: { fontSize: 11, numberFormat: '#,##0.000;[Red](#,##0.000)' } });
  expect(cell(3)?.style?.numberFormat).toBe('0.0');
  expect(session.undo()).toBe(true);
  expect(session.active).toEqual(prior);
  expect(session.redo()).toBe(true);
  session.execute(applyNumberFormat(session.capture(), [range], { code: 'General' }));
  expect(cell(1)?.style).toEqual({ bold: true });
  expect(cell(2)?.style).toEqual({ fontSize: 11 });
});

it('rejects invalid formats atomically and can preformat empty cells for dates and text', () => {
  const { session, range, target, cell } = fixture();
  const initial = session.active;
  expect(() => applyNumberFormat(session.capture(), [range], { code: '0.0"' })).toThrow(/valid Excel/);
  expect(session.active).toBe(initial);
  expect(session.canUndo).toBe(false);
  const date = '[$-405]dd.mm.yyyy';
  session.execute(applyNumberFormat(session.capture(), [range], { code: date }));
  const serial = 10.123456789;
  session.execute({ kind: 'set-cells', edits: [{ target: target(1), content: { kind: 'number', value: serial } }] });
  expect(editorEntryText(cell(1))).toBe('10.01.1900 2:57:46.6665696');
  expect(clipboardText(captureClipboard(session.capture(), { ...range, rows: 1 }, false))).toBe(String(serial));
  session.execute(pasteText(session.capture(), { ...range, row: 2, rows: 1 }, '29.02.1900 12:00:00', 'all'));
  expect(cell(2)?.content).toEqual({ kind: 'number', value: 60.5 });
  expect(parseEntry('60.5', { numberFormat: date })).toEqual({ kind: 'number', value: 60.5 });
  expect(parseEntry('29.02.1900', { numberFormat: '@' })).toEqual({ kind: 'text', value: '29.02.1900' });
  expect(parseEntry('=A1', { numberFormat: date })).toEqual({ kind: 'formula', source: '=A1' });
});
