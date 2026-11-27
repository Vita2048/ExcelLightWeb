import { describe, expect, it } from 'vitest';
import { clear } from '../../src/application/editing-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { applyPythonResult } from '../../src/core/python-spills';
import { defaultPythonSource, documentLimits, findCell, newWorkbook, type PythonBlock } from '../../src/core/workbook';
import { loadNativePackage, saveNativePackage } from '../../src/persistence/native';
import { buildPythonInputs } from '../../src/python/python-inputs';
import { pythonBlockFromOutput, pythonInsertCommand } from '../../src/python/python-results';

function harness() {
  let next = 0;
  const makeId = () => `00000000-0000-4000-8000-${String(++next).padStart(12, '0')}`;
  const document = newWorkbook(makeId);
  const sheetId = document.sheets[0]!.id;
  const session = new WorkbookSession(document, { makeId });
  const at = (row: number, column: number) => ({ sheetId, row, column });
  session.execute({ kind: 'set-cells', label: 'Python inputs', edits: [
    { target: at(1, 2), content: { kind: 'number', value: 5 } },
    { target: at(2, 2), content: { kind: 'number', value: 8 } },
    { target: at(1, 3), content: { kind: 'number', value: 60 }, style: { numberFormat: 'yyyy-mm-dd' }, preserveStyle: false },
  ] });
  const source = 'xl("B1:B2")';
  const block: PythonBlock = { id: makeId(), source, bindings: [], lastValue: null, hasResult: false,
    inputFingerprint: null, runtimeFingerprint: null, status: 'never', error: null,
    lastTable: null, spillActive: false, lastFigure: null, placements: [] };
  session.execute({ kind: 'set-cells', label: 'Python anchor', edits: [{ target: at(1, 4),
    content: { kind: 'python', blockId: block.id, source, block } }] });
  const anchor = { ...at(1, 4), blockId: block.id, source };
  return { makeId, session, sheetId, at, block, anchor };
}

function successfulTable(block: PythonBlock, values: number[][] = [[5], [8]]) {
  return pythonBlockFromOutput(block, { status: 'succeeded', result: { kind: 'table', headers: ['Value'], rows: values,
    notices: [] }, stdout: '', stderr: '', outputTruncated: false }, 'inputs-a', 'pyodide-test');
}

describe('S20 manually executed Python results', () => {
  it('snapshots literal xl ranges and validates headers before creating a runtime input', () => {
    const { session, anchor, at } = harness();
    const built = buildPythonInputs(session.capture(), anchor);
    expect(built.input.calls).toEqual([{ reference: 'B1:B2', headers: false, key: '__xl_0' }]);
    expect(built.input.values.__xl_0).toEqual({ rows: [[5], [8]] });
    expect(built.fingerprint).toMatch(/^fnv1a64-[0-9a-f]{16}$/u);

    const sheet = session.active.sheets[0]!;
    const anchorCell = findCell(sheet, { row: 1, column: 4 })!;
    const headersSource = 'xl("B1:B2", headers=True)';
    const headersAnchor = { ...anchor, source: headersSource };
    const priorContent = anchorCell.content;
    if (priorContent.kind !== 'python') throw new Error('Expected Python anchor');
    const withSource = { ...session.capture(), document: { ...session.active, sheets: [{ ...sheet,
      cells: sheet.cells.map(cell => cell.id === anchorCell.id ? { ...cell, content: {
        ...priorContent, source: headersSource,
        ...(priorContent.block && { block: { ...priorContent.block, source: headersSource } }),
      } } : cell),
    }] } };
    expect(() => buildPythonInputs(withSource, headersAnchor)).toThrow(/nonblank unique text headers/u);

    const dateSource = 'xl("C1")';
    const currentContent = anchorCell.content;
    if (currentContent.kind !== 'python') throw new Error('Expected Python anchor');
    const dateDocument = { ...session.active, sheets: [{ ...sheet, cells: sheet.cells.map(cell => cell.id === anchorCell.id ? {
      ...cell, content: { ...currentContent, source: dateSource,
        ...(currentContent.block && { block: { ...currentContent.block, source: dateSource } }) },
    } : cell) }] };
    const dateSnapshot = { ...session.capture(), document: dateDocument };
    const dateAnchor = { ...anchor, source: dateSource };
    expect(() => buildPythonInputs(dateSnapshot, dateAnchor)).toThrow(/1900-02-29/u);
    session.execute({ kind: 'set-cells', label: 'Correct Excel date', edits: [{ target: at(1, 3), content: { kind: 'number', value: 61 } }] });
    const validDateSnapshot = { ...session.capture(), document: { ...session.active, sheets: [{ ...session.active.sheets[0]!,
      cells: session.active.sheets[0]!.cells.map(cell => cell.id === anchorCell.id ? { ...cell, content: {
        ...currentContent, source: dateSource, ...(currentContent.block && { block: { ...currentContent.block, source: dateSource } }),
      } } : cell),
    }] } };
    expect(buildPythonInputs(validDateSnapshot, dateAnchor).input.values.__xl_0).toEqual({ rows: [], scalarCell: true, scalar: '1900-03-01' });
  });

  it('rejects an occupied spill without overwriting and retains the previous successful result', () => {
    const { session, anchor, block, at, makeId } = harness();
    session.execute({ kind: 'set-cells', label: 'occupied output', edits: [{ target: at(2, 4), content: { kind: 'text', value: 'Keep me' } }] });
    const table = successfulTable(block);
    expect(() => applyPythonResult(session.active, at(1, 4), table, makeId)).toThrow(/#SPILL!/u);
    expect(findCell(session.active.sheets[0]!, at(2, 4))?.content).toEqual({ kind: 'text', value: 'Keep me' });

    const previous = { ...block, status: 'succeeded', hasResult: true, lastValue: { kind: 'number' as const, value: 42 },
      inputFingerprint: 'prior', runtimeFingerprint: 'prior-runtime' };
    const failed = pythonBlockFromOutput(previous, { status: 'failed', stdout: 'log', stderr: '', error: 'ValueError', outputTruncated: false },
      'inputs-b', 'pyodide-test');
    expect(failed).toMatchObject({ status: 'failed', error: 'ValueError', hasResult: true,
      lastValue: { kind: 'number', value: 42 }, inputFingerprint: 'inputs-b' });
  });

  it('applies a complete spill in one history transaction, persists it, blocks child edits and removes children when it shrinks', async () => {
    const { session, anchor, block, at, makeId } = harness();
    const result = successfulTable(block);
    const next = applyPythonResult(session.active, at(1, 4), result, makeId);
    session.execute({ kind: 'structure', document: next, label: 'Run Python' });
    expect(session.undoCount).toBe(3);
    expect(findCell(session.active.sheets[0]!, at(1, 4))?.content).toMatchObject({ kind: 'python', block: {
      hasResult: true, spillActive: true, lastTable: { rows: 3, columns: 1 },
    } });
    expect(findCell(session.active.sheets[0]!, at(2, 4))).toMatchObject({ content: { kind: 'number', value: 5 }, spillOwner: block.id });
    expect(findCell(session.active.sheets[0]!, at(3, 4))).toMatchObject({ content: { kind: 'number', value: 8 }, spillOwner: block.id });
    expect(await loadNativePackage(await saveNativePackage(session.active))).toEqual(session.active);
    expect(() => session.execute({ kind: 'set-cells', label: 'Edit child', edits: [
      { target: at(2, 4), content: { kind: 'number', value: 20 } },
    ] })).toThrow(/Generated cell is read-only/u);

    const scalar = pythonBlockFromOutput(result, { status: 'succeeded', result: { kind: 'scalar', value: 12 },
      stdout: '', stderr: '', outputTruncated: false }, 'inputs-b', 'pyodide-test');
    session.execute({ kind: 'structure', document: applyPythonResult(session.active, at(1, 4), scalar, makeId), label: 'Run Python' });
    expect(findCell(session.active.sheets[0]!, at(2, 4))).toBeUndefined();
    expect(findCell(session.active.sheets[0]!, at(3, 4))).toBeUndefined();
    expect(findCell(session.active.sheets[0]!, at(1, 4))?.content).toMatchObject({ kind: 'python', block: {
      lastValue: { kind: 'number', value: 12 }, spillActive: false,
    } });
    expect(session.undo()).toBe(true);
    expect(findCell(session.active.sheets[0]!, at(3, 4))?.spillOwner).toBe(block.id);
  });

  it('replaces styled spill cells in place when a subsequent generation updates the same extent', () => {
    const { session, at, block, makeId } = harness();
    session.execute({ kind: 'set-cells', label: 'Style spill output', edits: [
      { target: at(2, 4), content: { kind: 'blank' }, style: { fillRgb: 'EAF1F7' }, preserveStyle: false },
      { target: at(3, 4), content: { kind: 'blank' }, style: { fillRgb: 'FFFFFF' }, preserveStyle: false },
    ] });
    const first = applyPythonResult(session.active, at(1, 4), successfulTable(block), makeId);
    session.execute({ kind: 'structure', document: first, label: 'First Python generation' });
    const firstAnchor = findCell(session.active.sheets[0]!, at(1, 4));
    if (firstAnchor?.content.kind !== 'python' || !firstAnchor.content.block) throw new Error('Expected Python result block.');
    const refreshed = pythonBlockFromOutput(firstAnchor.content.block, {
      status: 'succeeded', result: { kind: 'table', headers: ['Value'], rows: [[9], [10]], notices: [] },
      stdout: '', stderr: '', outputTruncated: false,
    }, 'inputs-b', 'pyodide-test');

    const second = applyPythonResult(session.active, at(1, 4), refreshed, makeId);
    session.execute({ kind: 'structure', document: second, label: 'Second Python generation' });
    const cells = session.active.sheets[0]!.cells;
    expect(new Set(cells.map(cell => `${cell.row}:${cell.column}`)).size).toBe(cells.length);
    expect(findCell(session.active.sheets[0]!, at(2, 4))).toMatchObject({
      content: { kind: 'number', value: 9 }, style: { fillRgb: 'EAF1F7' }, spillOwner: block.id,
    });
    expect(findCell(session.active.sheets[0]!, at(3, 4))).toMatchObject({
      content: { kind: 'number', value: 10 }, style: { fillRgb: 'FFFFFF' }, spillOwner: block.id,
    });
  });

  it('clears anchor-owned spill state and inserts the normalized result as ordinary values', () => {
    const { session, anchor, block, at, makeId } = harness();
    const result = successfulTable(block);
    session.execute({ kind: 'structure', document: applyPythonResult(session.active, at(1, 4), result, makeId), label: 'Run Python' });
    const command = pythonInsertCommand(session.active, at(1, 4), result, makeId);
    session.execute(command);
    expect(findCell(session.active.sheets[0]!, at(1, 4))?.content).toEqual({ kind: 'text', value: 'Value' });
    expect(findCell(session.active.sheets[0]!, at(2, 4))?.content).toEqual({ kind: 'number', value: 5 });
    expect(findCell(session.active.sheets[0]!, at(3, 4))?.content).toEqual({ kind: 'number', value: 8 });
    expect(session.active.sheets[0]!.cells.some(cell => cell.content.kind === 'python')).toBe(false);
    expect(session.active.sheets[0]!.cells.some(cell => cell.spillOwner)).toBe(false);

    session.undo();
    session.execute(clear(session.capture(), [{ ...at(1, 4), rows: 1, columns: 1 }], 'contents'));
    expect(findCell(session.active.sheets[0]!, at(1, 4))).toBeUndefined();
    expect(findCell(session.active.sheets[0]!, at(2, 4))).toBeUndefined();
    expect(findCell(session.active.sheets[0]!, at(3, 4))).toBeUndefined();
    expect(session.undo()).toBe(true);
    expect(findCell(session.active.sheets[0]!, at(3, 4))?.spillOwner).toBe(block.id);
    expect(anchor.blockId).toBe(block.id);
  });

  it('bounds logs, scalar text and table cell limits before workbook commit', () => {
    const { block } = harness();
    const base = { status: 'succeeded' as const, stdout: '', stderr: '', outputTruncated: false };
    expect(() => pythonBlockFromOutput(block, { ...base, result: { kind: 'scalar', value: 'x'.repeat(32_768) } },
      'input', 'runtime')).toThrow(/32,767/u);
    expect(() => pythonBlockFromOutput(block, { ...base, stdout: 'x'.repeat(1024 * 1024 + 1),
      result: { kind: 'scalar', value: 1 } }, 'input', 'runtime')).toThrow(/log/u);
    expect(() => pythonBlockFromOutput(block, { ...base, result: { kind: 'table', headers: null,
      rows: [Array.from({ length: documentLimits.columns + 1 }, () => 1)], notices: [] } }, 'input', 'runtime')).toThrow(/shape/u);
  });
});
