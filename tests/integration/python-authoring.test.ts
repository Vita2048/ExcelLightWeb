import { describe, expect, it } from 'vitest';
import { captureClipboard, fill, parseEntry, pasteBlock } from '../../src/application/editing-commands';
import { sortData } from '../../src/application/data-commands';
import { changeSheet, editAxis } from '../../src/application/structure-commands';
import { defineName } from '../../src/application/table-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { defaultPythonSource, emptySheet, findCell, type WorkbookDocument } from '../../src/core/workbook';
import { pythonSourceCommand, pythonAnchorAt } from '../../src/features/python/python-editor-model';

function fixture() {
  let n = 0;
  const makeId = () => 'id-' + ++n;
  const document: WorkbookDocument = { id: 'book', formatVersion: 1, dateSystem: '1900', sheets: [
    emptySheet('data', 'Data'), emptySheet('other', "O'Brien"),
  ], names: [{ id: 'tax', name: 'TaxRate', area: { sheetId: 'data', row: 1, column: 1, rows: 1, columns: 1 } }] };
  const session = new WorkbookSession(document, { makeId });
  const at = (row: number, column: number, sheetId = 'data') => ({ sheetId, row, column });
  const range = (row: number, column: number, rows = 1, columns = 1, sheetId = 'data') => ({ sheetId, row, column, rows, columns });
  const cell = (row: number, column: number, sheetId = 'data') => findCell(session.active.sheets.find(sheet => sheet.id === sheetId)!, { row, column });
  const source = (row: number, column: number, value: string, sheetId = 'data') => session.execute({
    kind: 'set-cells', label: 'anchor', edits: [{ target: at(row, column, sheetId), content: {
      kind: 'python', blockId: 'block-' + row + '-' + column,
      source: value, block: { id: 'block-' + row + '-' + column, source: value,
        bindings: [{ name: 'legacy', reference: '=A1' }], status: 'never' } } }],
  });
  return { session, makeId, at, range, cell, source };
}

describe('S20 typed Python anchors', () => {
  it('creates one empty, inert anchor for =PY or =PY() and preserves an existing block', () => {
    const { makeId } = fixture();
    const created = parseEntry(' =Py() ', undefined, 'en-US', { makeId });
    expect(created).toMatchObject({ kind: 'python', source: defaultPythonSource, block: {
      source: defaultPythonSource, bindings: [], status: 'never' } });
    expect(created.kind === 'python' && created.source).toBe('');
    expect(created.kind === 'python' && created.block?.id).toBe(created.kind === 'python' && created.blockId);
    expect(parseEntry('=PY', undefined, 'en-US', { makeId, existing: created })).toBe(created);
    expect(parseEntry('=PY()', { numberFormat: '@' })).toEqual({ kind: 'text', value: '=PY()' });
    expect(() => parseEntry('=PY()')).toThrow(/identity/);
  });

  it('commits source without executing and keeps it separate from workbook history', () => {
    const { session, at, cell, source } = fixture();
    source(2, 1, defaultPythonSource);
    const target = pythonAnchorAt(session.capture(), 'data', { row: 2, column: 1 })!;
    session.execute(pythonSourceCommand(session.capture(), target, 'xl("A1:A3")\n')!);
    expect(cell(2, 1)?.content).toMatchObject({ kind: 'python', source: 'xl("A1:A3")\n',
      block: { source: 'xl("A1:A3")\n', status: 'never', bindings: [{ name: 'legacy', reference: '=A1' }] } });
    expect(cell(2, 1)?.cache).toBeUndefined();
    expect(session.undoCount).toBe(2);
    expect(session.undo()).toBe(true);
    expect(cell(2, 1)?.content).toMatchObject({ kind: 'python', source: defaultPythonSource });
    session.undo();
    expect(findCell(session.active.sheets[0]!, { row: 2, column: 1 })).toBeUndefined();
    expect(at(2, 1)).toEqual({ sheetId: 'data', row: 2, column: 1 });
  });
});

describe('S20 xl reference rewriting', () => {
  it('translates copied source and legacy bindings, and gives the copy independent identity', () => {
    const { session, makeId, range, cell, source } = fixture();
    source(1, 3, 'xl("B1:B2") + xl("TaxRate")\n');
    const original = cell(1, 3)!.content;
    const clipboard = captureClipboard(session.capture(), range(1, 3), false);
    session.execute(pasteBlock(session.capture(), range(2, 4), clipboard, 'all', makeId));
    const copied = cell(2, 4)!.content;
    expect(copied).toMatchObject({ kind: 'python', source: 'xl("C2:C3") + xl("TaxRate")\n' });
    expect(copied).toMatchObject({ block: { source: 'xl("C2:C3") + xl("TaxRate")\n',
      bindings: [{ name: 'legacy', reference: '=B2' }], status: 'never' } });
    const copiedPython = 'blockId' in copied ? copied : undefined;
    expect(copiedPython?.blockId).not.toBe('blockId' in original ? original.blockId : undefined);
    expect(copiedPython?.block?.id).toBe(copiedPython?.blockId);
  });

  it('moves anchors with stable identity and keeps relative and absolute literals correct', () => {
    const { session, makeId, range, cell, source } = fixture();
    source(2, 2, 'xl("$B$2") + xl("C3")\n');
    const content = cell(2, 2)!.content;
    const id = content.kind === 'python' ? content.blockId : '';
    const clipboard = captureClipboard(session.capture(), range(2, 2), true);
    session.execute(pasteBlock(session.capture(), range(3, 3), clipboard, 'all', makeId));
    expect(cell(2, 2)).toBeUndefined();
    expect(cell(3, 3)?.content).toMatchObject({ kind: 'python', blockId: id,
      source: 'xl("$C$3") + xl("C3")\n', block: { id, bindings: [{ reference: '=A1' }] } });
  });

  it('translates filled and sorted anchors', () => {
    const { session, makeId, range, cell, source } = fixture();
    source(1, 2, 'xl("A1")\n');
    session.execute(fill(session.capture(), range(1, 2, 2), 'down', makeId));
    session.execute({
      kind: 'set-cells', label: 'sort keys', edits: [1, 2, 3].map(row => ({
        target: { sheetId: 'data', row, column: 1 }, content: { kind: 'number' as const, value: row === 1 ? 3 : row === 2 ? 1 : 2 },
      })),
    });
    session.execute(sortData(session.capture(), range(1, 1, 3, 3), 1, false, true, false));
    expect(cell(1, 2)?.content).toMatchObject({ kind: 'python', source: 'xl("A1")\n' });
    expect(cell(2, 2)).toBeUndefined();
    expect(cell(3, 2)?.content).toMatchObject({ kind: 'python', source: 'xl("A3")\n' });
  });

  it('rewrites structural edits, sheet renames and workbook name renames in authored source', () => {
    const { session, makeId, cell, source } = fixture();
    source(3, 2, 'xl("A1") + xl("A2:A3") + xl("\'O\'\'Brien\'!B2")\n');
    session.execute(editAxis(session.capture(), 'data', { axis: 'row', start: 2, count: 1, delete: false }));
    expect(cell(4, 2)?.content).toMatchObject({ kind: 'python',
      source: 'xl("A1") + xl("A3:A4") + xl("\'O\'\'Brien\'!B2")\n' });
    session.execute(editAxis(session.capture(), 'data', { axis: 'row', start: 1, count: 1, delete: true }));
    expect(cell(3, 2)?.content).toMatchObject({ kind: 'python',
      source: 'xl("#REF!") + xl("A2:A3") + xl("\'O\'\'Brien\'!B2")\n' });
    const renamed = changeSheet(session.capture(), 'other', { kind: 'rename', name: 'O Brien' }, makeId);
    session.execute(renamed.command);
    expect(cell(3, 2)?.content).toMatchObject({ kind: 'python',
      source: 'xl("#REF!") + xl("A2:A3") + xl("\'O Brien\'!B2")\n' });

    source(3, 4, 'xl("TaxRate")\n');
    const name = session.active.names.find(entry => entry.name === 'TaxRate')!;
    session.execute(defineName(session.capture(), { ...name, name: 'Rate' }));
    expect(cell(3, 4)?.content).toMatchObject({ kind: 'python', source: 'xl("Rate")\n' });
  });
});
