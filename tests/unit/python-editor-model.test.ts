import { expect, test } from 'vitest';
import type { WorkbookSnapshot } from '../../src/application/workbook-protocol';
import { WorkbookSession } from '../../src/application/workbook-session';
import { defaultPythonSource, emptySheet, findCell, type WorkbookDocument } from '../../src/core/workbook';
import {
  acceptsPythonEditorMessage, currentPythonSource, pythonAnchorAt, pythonSourceCommand, samePythonTarget,
} from '../../src/features/python/python-editor-model';

function session() {
  let n = 0;
  const makeId = () => 'id-' + ++n;
  const document: WorkbookDocument = { id: 'book', formatVersion: 1, dateSystem: '1900',
    sheets: [emptySheet('sheet', 'Sheet1')], names: [] };
  return { session: new WorkbookSession(document, { makeId }), makeId };
}

test('resolves anchors and owned spill children to their source anchor', () => {
  const { session: s } = session();
  s.execute({ kind: 'set-cells', label: 'anchor', edits: [{ target: { sheetId: 'sheet', row: 2, column: 3 },
    content: { kind: 'python', blockId: 'block-a', source: '40 + 2\n', block: {
      id: 'block-a', source: '40 + 2\n', bindings: [{ name: 'legacy', reference: '=A1' }], status: 'never' } } }] });
  const target = pythonAnchorAt(s.capture(), 'sheet', { row: 2, column: 3 })!;
  expect(target).toMatchObject({ blockId: 'block-a', cellId: expect.any(String), row: 2, column: 3, source: '40 + 2\n' });
  const snapshot = s.capture();
  const spill: WorkbookSnapshot = { ...snapshot, document: { ...snapshot.document, sheets: snapshot.document.sheets.map(sheet => ({
    ...sheet, cells: [...sheet.cells, { id: 'child', row: 3, column: 3, content: { kind: 'number', value: 42 },
      spillOwner: 'block-a' }],
  })) } };
  expect(pythonAnchorAt(spill, 'sheet', { row: 3, column: 3 })).toEqual(target);
  expect(pythonAnchorAt(spill, 'sheet', { row: 1, column: 1 })).toBeUndefined();
  expect(samePythonTarget(target, { ...target })).toBe(true);
  expect(samePythonTarget(target, { ...target, row: 4 })).toBe(false);
});

test('commits a source draft without changing block identity or legacy bindings and no execution state', () => {
  const { session: s } = session();
  s.execute({ kind: 'set-cells', label: 'anchor', edits: [{ target: { sheetId: 'sheet', row: 1, column: 1 },
    content: { kind: 'python', blockId: 'block-a', source: defaultPythonSource, block: {
      id: 'block-a', source: defaultPythonSource, bindings: [{ name: 'legacy', reference: '=B1' }], status: 'never' } } }] });
  const target = pythonAnchorAt(s.capture(), 'sheet', { row: 1, column: 1 })!;
  const command = pythonSourceCommand(s.capture(), target, 'import pandas as pd\nxl("A1:A3")\n')!;
  s.execute(command);
  const content = findCell(s.active.sheets[0]!, { row: 1, column: 1 })!.content;
  expect(content).toEqual({ kind: 'python', blockId: 'block-a', source: 'import pandas as pd\nxl("A1:A3")\n',
    block: { id: 'block-a', source: 'import pandas as pd\nxl("A1:A3")\n',
      bindings: [{ name: 'legacy', reference: '=B1' }], status: 'never', error: null } });
  expect(s.undo()).toBe(true);
  expect(currentPythonSource(s.capture(), target)).toBe(defaultPythonSource);
  expect(currentPythonSource(s.capture(), { ...target, sessionId: 'other' })).toBeUndefined();
  expect(pythonSourceCommand(s.capture(), target, defaultPythonSource)).toBeUndefined();
});

test('rejects stale editor contexts and moved targets', () => {
  const { session: s } = session();
  s.execute({ kind: 'set-cells', label: 'anchor', edits: [{ target: { sheetId: 'sheet', row: 1, column: 1 },
    content: { kind: 'python', blockId: 'block-a', source: '1\n', block: { id: 'block-a', source: '1\n', bindings: [] } } }] });
  const target = pythonAnchorAt(s.capture(), 'sheet', { row: 1, column: 1 })!;
  expect(acceptsPythonEditorMessage('context-a', { contextId: 'context-a', text: '2' })).toBe(true);
  expect(acceptsPythonEditorMessage('context-b', { contextId: 'context-a', text: '2' })).toBe(false);
  expect(acceptsPythonEditorMessage(undefined, { contextId: 'context-a', text: '2' })).toBe(false);
  s.execute({ kind: 'set-cells', label: 'move', edits: [{
    target: { sheetId: 'sheet', row: 1, column: 2 }, content: { kind: 'blank' } },
  ] });
  expect(currentPythonSource(s.capture(), target)).toBe('1\n');
  expect(pythonSourceCommand(s.capture(), { ...target, cellId: 'missing' }, '2')).toBeUndefined();
});
