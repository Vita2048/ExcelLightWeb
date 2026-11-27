import { describe, expect, it } from 'vitest';
import { createTable } from '../../src/application/table-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { calculateMixedGeneration } from '../../src/calculation/mixed-generation';
import { runNativeGeneration } from '../../src/calculation/native-generation';
import { applyPythonResult } from '../../src/core/python-spills';
import { emptySheet, findCell, newWorkbook, type CellDocument } from '../../src/core/workbook';
import { NativePackage } from '../../src/persistence/native/native-package';

function setupSession() {
  let n = 0;
  const makeId = () => `00000000-0000-4000-8000-${String(++n).padStart(12, '0')}`;
  const session = new WorkbookSession(newWorkbook(makeId), { makeId });
  const sheetId = session.active.sheets[0]!.id;
  const at = (row: number, column: number) => ({ sheetId, row, column });
  const calc = () => {
    const req = session.beginNativeCalculation('runtime-1');
    const res = runNativeGeneration(req.document, 'automatic');
    expect(session.tryCommitNativeCalculation({
      ...req,
      kind: 'native-calculated',
      updates: res.updates,
    })).toBe(true);
  };
  return { session, sheetId, at, makeId, calc };
}

describe('Sprint 1.3 dynamic array spills and shared spill ownership', () => {
  it('commits native spills without advancing authoredRevision, undo history, or dirty state', () => {
    const { session, sheetId, at, calc } = setupSession();
    session.execute({
      kind: 'set-cells',
      edits: [{ target: at(1, 1), content: { kind: 'formula', source: '=MAKEARRAY(2,2,LAMBDA(r,c,r*10+c))' } }],
    });
    session.markSaved(session.capture());
    expect(session.isDirty).toBe(false);
    const authoredBefore = session.capture().authoredRevision;
    const undoBefore = session.undoCount;

    calc();

    expect(session.capture().authoredRevision).toBe(authoredBefore);
    expect(session.undoCount).toBe(undoBefore);
    expect(session.isDirty).toBe(false);
    const sheet = session.active.sheets[0]!;
    expect(findCell(sheet, at(1, 1))?.spillExtent).toEqual({ rows: 2, columns: 2 });
    expect(findCell(sheet, at(1, 1))?.cache).toEqual({ kind: 'number', value: 11 });
    expect(findCell(sheet, at(1, 2))?.content).toEqual({ kind: 'number', value: 12 });
    expect(findCell(sheet, at(2, 1))?.content).toEqual({ kind: 'number', value: 21 });
    expect(findCell(sheet, at(2, 2))?.content).toEqual({ kind: 'number', value: 22 });
    expect(findCell(sheet, at(2, 2))?.spillKind).toBe('formula');
    expect(findCell(sheet, at(2, 2))?.spillOwner).toBe(findCell(sheet, at(1, 1))!.id);

    // Generated spill child is read-only.
    expect(() => session.execute({
      kind: 'set-cells',
      edits: [{ target: at(2, 2), content: { kind: 'number', value: 99 } }],
    })).toThrow(/read-only/i);

    // Clearing the anchor removes owned spill children, and undo restores cleanly.
    session.execute({
      kind: 'set-cells',
      edits: [{ target: at(1, 1), content: { kind: 'blank' } }],
    });
    expect(findCell(session.active.sheets[0]!, at(2, 2))).toBeUndefined();
    expect(session.undo()).toBe(true);
    expect(findCell(session.active.sheets[0]!, at(1, 1))?.content).toEqual({
      kind: 'formula',
      source: '=MAKEARRAY(2,2,LAMBDA(r,c,r*10+c))',
    });
    expect(session.active.sheets[0]!.id).toBe(sheetId);
  });

  it('produces #SPILL! on populated cell, merge, Table, and Python spill collisions and recovers automatically', () => {
    const { session, sheetId, at, makeId, calc } = setupSession();

    // 1. Populated cell blocker + automatic recovery when cleared.
    session.execute({
      kind: 'set-cells',
      edits: [
        { target: at(1, 1), content: { kind: 'formula', source: '=MAKEARRAY(2,2,LAMBDA(r,c,r+c))' } },
        { target: at(2, 2), content: { kind: 'text', value: 'blocker' } },
      ],
    });
    calc();
    expect(findCell(session.active.sheets[0]!, at(1, 1))?.cache).toEqual({ kind: 'error', value: '#SPILL!' });
    expect(findCell(session.active.sheets[0]!, at(1, 1))?.spillExtent).toBeUndefined();

    session.execute({
      kind: 'set-cells',
      edits: [{ target: at(2, 2), content: { kind: 'blank' } }],
    });
    calc();
    expect(findCell(session.active.sheets[0]!, at(1, 1))?.cache).toEqual({ kind: 'number', value: 2 });
    expect(findCell(session.active.sheets[0]!, at(2, 2))?.content).toEqual({ kind: 'number', value: 4 });

    // 2. Merged cells collision.
    session.execute({
      kind: 'structure',
      label: 'merge blocker',
      document: {
        ...session.active,
        sheets: [{
          ...emptySheet(sheetId, 'Sheet1'),
          merges: [{ row: 2, column: 2, rows: 1, columns: 2 }],
          cells: [{ id: makeId(), row: 1, column: 1, content: { kind: 'formula', source: '=MAKEARRAY(2,2,LAMBDA(r,c,r+c))' } }],
        }],
      },
    });
    calc();
    expect(findCell(session.active.sheets[0]!, at(1, 1))?.cache).toEqual({ kind: 'error', value: '#SPILL!' });

    // 3. Table collision.
    session.execute({
      kind: 'structure',
      label: 'reset for table',
      document: {
        ...session.active,
        sheets: [{
          ...emptySheet(sheetId, 'Sheet1'),
          cells: [
            { id: makeId(), row: 1, column: 1, content: { kind: 'formula', source: '=MAKEARRAY(2,2,LAMBDA(r,c,r+c))' } },
            { id: makeId(), row: 2, column: 2, content: { kind: 'text', value: 'Col' } },
            { id: makeId(), row: 3, column: 2, content: { kind: 'number', value: 5 } },
          ],
        }],
      },
    });
    session.execute(createTable(session.capture(), {
      id: makeId(),
      name: 'Tbl',
      isTable: true,
      area: { sheetId, row: 2, column: 2, rows: 2, columns: 1 },
    }, true, makeId));
    calc();
    expect(findCell(session.active.sheets[0]!, at(1, 1))?.cache).toEqual({ kind: 'error', value: '#SPILL!' });

    // 4. Python spill collision.
    const pyBlockId = makeId();
    session.execute({
      kind: 'structure',
      label: 'python spill setup',
      document: {
        ...session.active,
        names: [],
        sheets: [{
          ...emptySheet(sheetId, 'Sheet1'),
          cells: [
            { id: makeId(), row: 1, column: 1, content: { kind: 'formula', source: '=MAKEARRAY(2,2,LAMBDA(r,c,r+c))' } },
            { id: makeId(), row: 1, column: 3, content: { kind: 'python', blockId: pyBlockId, source: '[[10],[20]]' } },
          ],
        }],
      },
    });
    const withPySpill = applyPythonResult(session.active, at(1, 3), {
      id: pyBlockId,
      source: '[[10],[20]]',
      bindings: [],
      hasResult: true,
      status: 'succeeded',
      lastValue: { kind: 'number', value: 10 },
      spillActive: true,
      lastTable: { rows: 2, columns: 1, cells: [{ kind: 'number', value: 10 }, { kind: 'number', value: 20 }], notices: [] },
    }, makeId);
    session.execute({ kind: 'structure', label: 'apply python spill', document: withPySpill });
    // Now spill A1 across C1:C2 by using 2x3 array.
    session.execute({
      kind: 'set-cells',
      edits: [{ target: at(1, 1), content: { kind: 'formula', source: '=MAKEARRAY(2,3,LAMBDA(r,c,r+c))' } }],
    });
    calc();
    expect(findCell(session.active.sheets[0]!, at(1, 1))?.cache).toEqual({ kind: 'error', value: '#SPILL!' });
  });

  it('chains native formula spills into Python inputs and round-trips through native .exlt', async () => {
    const { session, sheetId, at, makeId, calc } = setupSession();
    const pyId = makeId();
    const pyBlockId = makeId();
    const pySource = 'xl("A1:B2", headers=False)';
    session.execute({
      kind: 'set-cells',
      edits: [
        { target: at(1, 1), content: { kind: 'formula', source: '=MAKEARRAY(2,2,LAMBDA(r,c,r*10+c))' } },
        { target: at(1, 4), content: { kind: 'formula', source: '=SUM(A1#)' } },
        {
          target: at(1, 5),
          cellId: pyId,
          content: {
            kind: 'python',
            blockId: pyBlockId,
            source: pySource,
            block: {
              id: pyBlockId,
              source: pySource,
              bindings: [],
              hasResult: false,
              lastValue: null,
              inputFingerprint: null,
              runtimeFingerprint: null,
              status: 'never',
              error: null,
              lastTable: null,
              spillActive: false,
              lastFigure: null,
              placements: [],
            },
          },
        },
      ],
    });
    calc();
    expect(findCell(session.active.sheets[0]!, at(1, 4))?.cache).toEqual({ kind: 'number', value: 66 });

    let receivedInput: unknown;
    const mixed = await calculateMixedGeneration({
      basis: session.capture(),
      requestedPythonIds: new Set([pyId]),
      signal: new AbortController().signal,
    }, {
      runtimeFingerprint: 'py-rt',
      makeId,
      calculateNative: async (doc) => {
        const res = runNativeGeneration(doc, 'automatic');
        return { updates: res.updates, staleCellIds: res.staleCellIds };
      },
      runPython: async (_source, input) => {
        receivedInput = input;
        return {
          runtimeFingerprint: 'py-rt',
          output: {
            status: 'succeeded',
            result: { kind: 'scalar', value: 66 },
            stdout: '',
            stderr: '',
            outputTruncated: false,
          },
        };
      },
    });
    expect(receivedInput).toMatchObject({
      values: { __xl_0: { rows: [[11, 12], [21, 22]] } },
    });
    expect(mixed.completedPythonIds.has(pyId)).toBe(true);
    expect(session.tryCommitGeneration(session.capture(), mixed.document, 'Mixed calculation')).toBe(true);

    // Native .exlt round-trip preserves spillKind and spillExtent cleanly.
    const blob = await NativePackage.save(session.active);
    const loaded = await NativePackage.load(blob);
    const loadedSheet = loaded.sheets.find(s => s.id === sheetId)!;
    const anchor: CellDocument | undefined = findCell(loadedSheet, at(1, 1));
    const child: CellDocument | undefined = findCell(loadedSheet, at(2, 2));
    expect(anchor?.spillExtent).toEqual({ rows: 2, columns: 2 });
    expect(child?.spillKind).toBe('formula');
    expect(child?.spillOwner).toBe(anchor?.id);
  });
});
