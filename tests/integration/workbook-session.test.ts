import { once } from 'node:events';
import { Worker } from 'node:worker_threads';
import { describe, expect, it } from 'vitest';
import { WorkbookSession, defaultHistoryBudget, type CellEdit } from '../../src/application/workbook-session';
import { applyDelta, invertDelta } from '../../src/core/workbook-delta';
import { DocumentError, documentLimits, emptySheet, findReferencedCell, newWorkbook } from '../../src/core/workbook';
import type { NativeCalculationReply, NativeCalculationRequest } from '../../src/application/workbook-protocol';

function setup(budget?: { transactions?: number; bytes?: number }) {
  let id = 0;
  const makeId = () => `id-${++id}`;
  const book = newWorkbook(makeId);
  const sheet = book.sheets[0]!;
  const session = new WorkbookSession(book, { makeId, ...(budget && { budget }) });
  const at = (row: number, column = 1) => ({ sheetId: sheet.id, row, column });
  const read = (row: number, column = 1) => findReferencedCell(session.active, at(row, column));
  const edit = (row: number, content: CellEdit['content'], extras: Partial<CellEdit> = {}) =>
    session.execute({ kind: 'set-cells', edits: [{ target: at(row), content, ...extras }] });
  const number = (row: number, value: number) => edit(row, { kind: 'number', value });
  return { session, book, sheet, at, read, edit, number };
}

describe('workbook transactions and history', () => {
  it('preserves typed empties, styled blanks, cell IDs, and immutable snapshots', () => {
    const { session, read, edit } = setup();
    const style = { fillRgb: 'FFFFFF', borderDetails: { bottom: { style: 'thick', rgb: '113355' } } };
    const content = { kind: 'text', value: '' } as const;
    edit(1, content, { style, preserveStyle: false });
    const initial = read(1)!;
    const snapshot = session.capture();
    style.borderDetails.bottom.rgb = '000000';
    expect(initial.style?.borderDetails?.bottom?.rgb).toBe('113355');
    expect(Object.isFrozen(initial.style?.borderDetails?.bottom)).toBe(true);
    expect(initial.content).toEqual(content);
    edit(1, { kind: 'boolean', value: false });
    expect(read(1)).toMatchObject({ id: initial.id, style: initial.style, content: { kind: 'boolean', value: false } });
    edit(1, { kind: 'blank' });
    expect(read(1)).toMatchObject({ id: initial.id, content: { kind: 'blank' } });
    edit(1, { kind: 'blank' }, { preserveStyle: false });
    expect(read(1)).toBeUndefined();
    expect(snapshot.document.sheets[0]!.cells[0]).toBe(initial);
    expect(session.undo()).toBe(true);
    expect(read(1)?.id).toBe(initial.id);
  });

  it('applies compound edits as one indexed delta and advances revision on undo and redo', () => {
    const { session, sheet, at, number, read } = setup();
    number(2, 2);
    number(1, 1);
    const before = session.capture();
    const outcome = session.execute({
      kind: 'set-cells', edits: [
        { target: { ...at(2), absoluteRow: true }, content: { kind: 'blank' } },
        { target: at(3), content: { kind: 'formula', source: '=A1' }, requiresExternalCalculation: true },
        { target: { sheetId: sheet.id, row: 3, column: 2 }, content: { kind: 'text', value: '' } },
      ],
    });
    const after = session.capture();
    expect(outcome.changed).toBe(true);
    expect(session.undoCount).toBe(3);
    expect(read(3)?.requiresExternalCalculation).toBe(true);
    expect(applyDelta(before.document, outcome.delta)).toEqual(after.document);
    expect(applyDelta(after.document, invertDelta(outcome.delta))).toEqual(before.document);
    expect(session.undo()).toBe(true);
    expect(session.active).toEqual(before.document);
    expect(session.revision).toBe(after.revision + 1);
    expect(session.redo()).toBe(true);
    expect(session.active).toEqual(after.document);
    expect(session.revision).toBe(after.revision + 2);
  });

  it('clears an unstyled blank carrying only an imported cache and restores it on undo', () => {
    const { session, book, sheet, read, edit } = setup();
    const cachedBlank = {
      id: 'cached-blank', row: 1, column: 1, content: { kind: 'blank' },
      cache: { kind: 'text', value: '' },
    } as const;
    session.adopt({ ...book, sheets: [{ ...sheet, cells: [cachedBlank] }] });
    expect(edit(1, { kind: 'blank' }).changed).toBe(true);
    expect(read(1)).toBeUndefined();
    expect(session.undo()).toBe(true);
    expect(read(1)).toEqual(cachedBlank);
    expect(session.isDirty).toBe(false);
  });

  it('rejects the full compound edit and preserves an existing redo branch when any edit is invalid', () => {
    const { session, at, number } = setup();
    number(1, 10);
    number(2, 20);
    session.undo();
    const basis = session.capture();
    const bytes = session.historyBytes;
    const invalid: CellEdit[] = [
      { target: at(3), content: { kind: 'formula', source: '=' } },
      { target: at(3), content: { kind: 'number', value: Infinity } },
      { target: at(3), content: { kind: 'text', value: 'x'.repeat(32_768) } },
      { target: at(0), content: { kind: 'number', value: 1 } },
      { target: at(3, documentLimits.columns + 1), content: { kind: 'number', value: 1 } },
      { target: { ...at(3), invalid: true }, content: { kind: 'number', value: 1 } },
      { target: { ...at(3), sheetId: 'foreign' }, content: { kind: 'number', value: 1 } },
      { target: at(3), content: { kind: 'number', value: 1 }, cellId: basis.document.sheets[0]!.cells[0]!.id },
      { target: at(3), content: { kind: 'blank' }, preserveStyle: false, style: { doubleUnderline: true } },
      { target: { ...at(4), absoluteColumn: true }, content: { kind: 'number', value: 2 } },
    ];
    for (const bad of invalid) {
      expect(() => session.execute({
        kind: 'set-cells', edits: [{ target: at(4), content: { kind: 'number', value: 7 } }, bad],
      })).toThrow(DocumentError);
      expect(session.capture()).toBe(basis);
      expect(session.historyBytes).toBe(bytes);
      expect(session.canRedo).toBe(true);
      expect(session.isDirty).toBe(true);
    }
    expect(number(1, 10).changed).toBe(false);
    expect(session.execute({ kind: 'set-cells', edits: [] }).changed).toBe(false);
    expect(session.capture()).toBe(basis);
    number(4, 7);
    expect(session.canRedo).toBe(false);
  });

  it('commits detached sheet/metadata changes without changing stable identities', () => {
    const { session, at, read, number } = setup();
    number(1, 50);
    const basis = session.capture();
    const oldSheet = basis.document.sheets[0]!;
    const draft = { ...basis.document, sheets: [
      emptySheet('another', 'Another'),
      { ...oldSheet, name: 'Renamed', rows: [{ index: 1, size: 50 }] },
    ] };
    expect(session.tryExecute(basis, { kind: 'structure', document: draft, label: 'Rename and add' })).toBe(true);
    expect(session.active.sheets[1]?.id).toBe(at(1).sheetId);
    expect(read(1)).toBe(oldSheet.cells[0]);
    draft.sheets.reverse();
    expect(session.active.sheets[0]?.id).toBe('another');
    expect(session.undo()).toBe(true);
    expect(session.active).toBe(basis.document);
    expect(session.redo()).toBe(true);
    expect(session.tryExecute(basis, { kind: 'structure', document: draft, label: 'Stale' })).toBe(false);
    expect(() => session.execute({ kind: 'structure', document: { ...draft, id: 'different' }, label: 'Bad' })).toThrow(DocumentError);
  });

  it('evicts oldest entries at 100, retains accounting during undo/redo, and keeps the saved basis', () => {
    const { session, number, read } = setup();
    for (let value = 1; value <= 103; value++) number(1, value);
    expect(session.undoCount).toBe(defaultHistoryBudget.transactions);
    const saved = session.capture();
    expect(session.markSaved(saved)).toBe(true);
    const bytes = session.historyBytes;
    for (let i = 0; i < 100; i++) expect(session.undo()).toBe(true);
    expect(session.undo()).toBe(false);
    expect(read(1)?.content).toEqual({ kind: 'number', value: 3 });
    expect(session.isDirty).toBe(true);
    expect(session.historyBytes).toBe(bytes);
    for (let i = 0; i < 100; i++) expect(session.redo()).toBe(true);
    expect(session.isDirty).toBe(false);
  });

  it('evicts by bytes and rejects oversized compound transactions before changing history or state', () => {
    const { session, edit } = setup({ bytes: 12_000 });
    for (let row = 1; row <= 10; row++) edit(row, { kind: 'text', value: 'x'.repeat(300) });
    expect(session.undoCount).toBeGreaterThan(0);
    expect(session.undoCount).toBeLessThan(10);
    expect(session.historyBytes).toBeLessThanOrEqual(12_000);
    session.undo();
    const basis = session.capture();
    const bytes = session.historyBytes;
    expect(() => edit(20, { kind: 'text', value: 'x'.repeat(10_000) })).toThrow(/Split/);
    expect(session.capture()).toBe(basis);
    expect(session.historyBytes).toBe(bytes);
    expect(session.canRedo).toBe(true);
  });
});

function completion(request: NativeCalculationRequest, updates: NativeCalculationReply['updates']): NativeCalculationReply {
  const { document: _document, kind: _kind, ...basis } = request;
  return { ...basis, kind: 'native-calculated', updates };
}

describe('workbook completion guards', () => {
  it('records the saved snapshot even when an edit wins the race and ignores saves from a prior session', async () => {
    const { session, number, book } = setup();
    number(1, 1);
    const saveBasis = session.capture();
    const finishSave = Promise.resolve().then(() => session.markSaved(saveBasis));
    number(1, 2);
    expect(await finishSave).toBe(true);
    expect(session.savedIdentity).toBe(saveBasis.sourceFingerprint);
    expect(session.isDirty).toBe(true);
    expect(session.markSaved({ ...session.capture() })).toBe(false);
    expect(session.undo()).toBe(true);
    expect(session.isDirty).toBe(false);
    expect(session.redo()).toBe(true);
    number(1, 1);
    expect(session.isDirty).toBe(false);
    const revision = session.revision;
    const previous = session.capture();
    expect(() => session.adopt({ ...book, sheets: [] })).toThrow(DocumentError);
    expect(session.capture()).toBe(previous);
    expect(session.tryAdopt(saveBasis, book)).toBe(false);
    expect(session.tryAdopt(previous, book, false)).toBe(true);
    expect(session.sessionId).not.toBe(previous.sessionId);
    expect(session.markSaved(previous)).toBe(false);
    expect(session.revision).toBe(revision + 1);
    expect(session.isDirty).toBe(true);
    expect(session.canUndo || session.canRedo).toBe(false);
  });

  it('accepts only current, once-only replies and commits typed native caches outside authored history', () => {
    const { session, at, edit, number, read } = setup();
    edit(1, { kind: 'formula', source: '=FALSE()' });
    number(2, 7);
    session.markSaved(session.capture());
    const request = session.beginNativeCalculation('runtime-a');
    const reply = completion(request, [{ target: at(1), cache: { kind: 'error', value: '#NAME?' } }]);
    for (const key of ['sessionId', 'requestId', 'runtimeFingerprint', 'sourceFingerprint', 'protocolVersion', 'revision', 'generation'] as const) {
      expect(session.tryCommitNativeCalculation({ ...reply, [key]: 'wrong' })).toBe(false);
    }
    const before = session.capture();
    expect(() => session.tryCommitNativeCalculation({ ...reply, updates: [...reply.updates, { target: at(2) }] })).toThrow(/native formula/);
    expect(() => session.tryCommitNativeCalculation({ ...reply, updates: [...reply.updates, ...reply.updates] })).toThrow(/Duplicate/);
    expect(() => session.tryCommitNativeCalculation({ ...reply, updates: [
      { target: at(1), cache: { kind: 'number', value: NaN } },
    ] })).toThrow(/finite/);
    expect(session.capture()).toBe(before);
    expect(session.tryCommitNativeCalculation(reply)).toBe(true);
    expect(read(1)?.cache).toEqual({ kind: 'error', value: '#NAME?' });
    expect(session.sourceFingerprint).toBe(before.sourceFingerprint);
    expect(session.revision).toBe(before.revision + 1);
    expect(session.undoCount).toBe(2);
    expect(session.isDirty).toBe(false);
    expect(session.tryCommitNativeCalculation(reply)).toBe(false);
    const clear = completion(session.beginNativeCalculation('runtime-a'), [{ target: at(1) }]);
    expect(session.tryCommitNativeCalculation(clear)).toBe(true);
    expect(read(1)?.cache).toBeUndefined();
    const current = session.revision;
    const noOp = completion(session.beginNativeCalculation('runtime-a'), []);
    expect(session.tryCommitNativeCalculation(noOp)).toBe(true);
    expect(session.revision).toBe(current);
    expect(session.tryCommitNativeCalculation(noOp)).toBe(false);
  });

  it('invalidates replies on worker replacement, cancel, edit/undo with identical authored state, and adoption', () => {
    const { session, number, book } = setup();
    const replaced = completion(session.beginNativeCalculation('old'), []);
    const cancelled = completion(session.beginNativeCalculation('new'), []);
    expect(session.tryCommitNativeCalculation(replaced)).toBe(false);
    session.cancelNativeCalculation();
    expect(session.tryCommitNativeCalculation(cancelled)).toBe(false);
    const obsolete = completion(session.beginNativeCalculation('new'), []);
    number(1, 10);
    expect(session.undo()).toBe(true);
    expect(session.sourceFingerprint).toBe(obsolete.sourceFingerprint);
    expect(session.isCurrent(obsolete)).toBe(false);
    expect(session.tryCommitNativeCalculation(obsolete)).toBe(false);
    const fromOldSession = completion(session.beginNativeCalculation('new'), []);
    session.adopt(book);
    expect(session.tryCommitNativeCalculation(fromOldSession)).toBe(false);
  });

  it('round-trips the worker DTO and rejects an in-flight reply after a synchronous edit', async () => {
    const { session, edit, read, at, number } = setup();
    edit(1, { kind: 'formula', source: '=1' });
    const worker = new Worker(new URL('./workbook-dto.worker.mjs', import.meta.url));
    try {
      function send() {
        const request = session.beginNativeCalculation('echo-runtime');
        const reply = completion(request, [{ target: at(1), cache: { kind: 'number', value: 1 } }]);
        const received = once(worker, 'message');
        worker.postMessage({ request, reply });
        return received;
      }
      const stale = send();
      number(2, 5);
      const [staleReply] = await stale;
      expect(session.tryCommitNativeCalculation(staleReply)).toBe(false);
      expect(read(1)?.cache).toBeUndefined();
      const [currentReply] = await send();
      expect(session.tryCommitNativeCalculation(currentReply)).toBe(true);
      expect(read(1)?.cache).toEqual({ kind: 'number', value: 1 });
    } finally {
      await worker.terminate();
    }
  });
});
