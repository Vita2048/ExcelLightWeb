import { describe, expect, it } from 'vitest';
import { WorkbookSession, type CellEdit } from '../../src/application/workbook-session';
import { emptySheet, findCell, type PythonBlock, type WorkbookDocument } from '../../src/core/workbook';
import { evaluateNative, indexNativeInputs, prepareNativeGraph } from '../../src/calculation/native-calculator';
import { calculateMixedGeneration, type MixedGenerationServices } from '../../src/calculation/mixed-generation';
import { MixedGraph } from '../../src/calculation/mixed-graph';
import { PythonRuntimeError } from '../../src/python/python-runtime-client';
import type { PythonRuntimeInput, PythonRuntimeOutput } from '../../src/python/runtime-protocol';

function fixture() {
  let serial = 0;
  const makeId = () => 'cell-' + ++serial;
  const document: WorkbookDocument = { id: 'book', formatVersion: 1, dateSystem: '1900',
    sheets: [emptySheet('sheet', 'Sheet1')], names: [] };
  const session = new WorkbookSession(document, { makeId });
  const at = (row: number, column: number) => ({ sheetId: 'sheet', row, column });
  const block = (source: string): PythonBlock => ({ id: makeId(), source, bindings: [], hasResult: false,
    lastValue: null, inputFingerprint: null, runtimeFingerprint: null, status: 'never', error: null,
    lastTable: null, spillActive: false, lastFigure: null, placements: [] });
  return { session, at, makeId, block };
}

function addCells(session: WorkbookSession, edits: readonly CellEdit[]) {
  session.execute({ kind: 'set-cells', edits, label: 'Fixture' });
}

function scalarOutput(value: number): PythonRuntimeOutput {
  return { status: 'succeeded', result: { kind: 'scalar', value }, stdout: '', stderr: '', outputTruncated: false };
}

function inputScalar(input: PythonRuntimeInput): number {
  const value = input.values.__xl_0;
  if (value?.scalarCell) return Number(value.scalar);
  return Number(value?.rows[0]?.[0]);
}

function services(run: (source: string, input: PythonRuntimeInput) => Promise<PythonRuntimeOutput> | PythonRuntimeOutput,
  calls: string[]): MixedGenerationServices {
  return {
    runtimeFingerprint: 'pyodide-mixed-test', makeId: () => 'spill-' + Math.random().toString(16).slice(2),
    async calculateNative(document, targets) {
      const graph = prepareNativeGraph(document);
      const selected = new Set<string>();
      const visit = (id: string) => {
        const node = graph.nodes.get(id);
        if (!node || selected.has(id)) return;
        selected.add(id);
        for (const input of node.inputs) visit(input.id);
      };
      for (const id of targets) visit(id);
      const calculated = evaluateNative(indexNativeInputs(document), graph,
        new Set([...selected].map(id => graph.nodes.get(id)!)), {});
      return { updates: calculated.updates, staleCellIds: [] };
    },
    async runPython(source, input) {
      calls.push(source);
      return { runtimeFingerprint: 'pyodide-mixed-test', output: await run(source, input) };
    },
  };
}

describe('S22 mixed recalculation', () => {
  it('does not mark Python stale when only its cached Figure placements change', () => {
    const { session, at, block } = fixture();
    const pythonBlock = block('');
    addCells(session, [{ target: at(1, 1), content: { kind: 'python', blockId: pythonBlock.id,
      source: pythonBlock.source, block: pythonBlock } }]);
    const before = session.active;
    const anchor = findCell(before.sheets[0]!, at(1, 1))!;
    if (anchor.content.kind !== 'python' || !anchor.content.block) throw new Error('Expected the Python anchor.');
    const placement = { id: 'placement', sheetId: 'sheet', row: 3, column: 3, rows: 4, columns: 5,
      width: 200, height: 120 };
    const after = { ...before, sheets: before.sheets.map(sheet => ({ ...sheet, cells: sheet.cells.map(cell =>
      cell.id === anchor.id && cell.content.kind === 'python' && cell.content.block
        ? { ...cell, content: { ...cell.content, block: { ...cell.content.block, placements: [placement] } } }
        : cell) })) };
    expect(new MixedGraph(after).affectedPython(before, after)).toEqual(new Set());

    const changedSource = { ...after, sheets: after.sheets.map(sheet => ({ ...sheet, cells: sheet.cells.map(cell =>
      cell.id === anchor.id && cell.content.kind === 'python'
        ? { ...cell, content: { ...cell.content, source: '40 + 2', block: { ...cell.content.block!, source: '40 + 2' } } }
        : cell) })) };
    expect(new MixedGraph(changedSource).affectedPython(before, changedSource)).toEqual(new Set([anchor.id]));
  });

  it('refreshes native prerequisites and commits native-to-Python-to-native chains as one history generation', async () => {
    const { session, at, block } = fixture();
    const a = block('xl("B1") + 1');
    const b = block('xl("D1") + 1');
    const commands = [
      { target: at(1, 1), content: { kind: 'number' as const, value: 4 } },
      { target: at(1, 2), content: { kind: 'formula' as const, source: '=A1*2' } },
      { target: at(1, 3), content: { kind: 'python' as const, blockId: a.id, source: a.source, block: a } },
      { target: at(1, 4), content: { kind: 'formula' as const, source: '=C1*2' } },
      { target: at(1, 5), content: { kind: 'python' as const, blockId: b.id, source: b.source, block: b } },
    ];
    addCells(session, commands);
    const before = session.undoCount;
    const basis = session.capture();
    const rootId = findCell(session.active.sheets[0]!, at(1, 5))!.id;
    const calls: string[] = [];
    const result = await calculateMixedGeneration({ basis, requestedPythonIds: new Set([rootId]),
      forcedPythonIds: new Set([rootId]), signal: new AbortController().signal }, services(
      (_source, input) => scalarOutput(inputScalar(input) + 1), calls));
    expect(calls).toEqual(['xl("B1") + 1', 'xl("D1") + 1']);
    expect(findCell(result.document.sheets[0]!, at(1, 2))?.cache).toEqual({ kind: 'number', value: 8 });
    expect(findCell(result.document.sheets[0]!, at(1, 3))?.content).toMatchObject({ kind: 'python', block: {
      lastValue: { kind: 'number', value: 9 }, status: 'succeeded',
    } });
    expect(findCell(result.document.sheets[0]!, at(1, 4))?.cache).toEqual({ kind: 'number', value: 18 });
    expect(findCell(result.document.sheets[0]!, at(1, 5))?.content).toMatchObject({ kind: 'python', block: {
      lastValue: { kind: 'number', value: 19 }, status: 'succeeded',
    } });
    expect(session.tryExecute(basis, { kind: 'structure', document: result.document, label: 'Run Python generation' })).toBe(true);
    expect(session.undoCount).toBe(before + 1);
  });

  it('adds a dependency revealed by a newly created spill and replans the remaining Python cell', async () => {
    const { session, at, block } = fixture();
    const producer = block('table');
    const consumer = block('xl("A2") + 1');
    addCells(session, [
      { target: at(1, 1), content: { kind: 'python', blockId: producer.id, source: producer.source, block: producer } },
      { target: at(1, 2), content: { kind: 'python', blockId: consumer.id, source: consumer.source, block: consumer } },
    ]);
    const producerId = findCell(session.active.sheets[0]!, at(1, 1))!.id;
    const consumerId = findCell(session.active.sheets[0]!, at(1, 2))!.id;
    const calls: string[] = [];
    const result = await calculateMixedGeneration({ basis: session.capture(),
      requestedPythonIds: new Set([producerId, consumerId]), forcedPythonIds: new Set([producerId, consumerId]),
      signal: new AbortController().signal }, services((source, input) => source === 'table'
      ? { status: 'succeeded', result: { kind: 'table', headers: null, rows: [[7], [9]], notices: [] },
        stdout: '', stderr: '', outputTruncated: false }
      : scalarOutput(inputScalar(input) + 1), calls));
    expect(calls).toEqual(['table', 'xl("A2") + 1']);
    expect(findCell(result.document.sheets[0]!, at(2, 1))).toMatchObject({ spillOwner: producer.id,
      content: { kind: 'number', value: 9 } });
    expect(findCell(result.document.sheets[0]!, at(1, 2))?.content).toMatchObject({ kind: 'python', block: {
      lastValue: { kind: 'number', value: 10 }, status: 'succeeded',
    } });
    expect(new MixedGraph(result.document).nodes.get(consumerId)?.dependencies.has(producerId)).toBe(true);
  });

  it('blocks mixed cycles and skips downstream Python after an upstream failure while keeping last success', async () => {
    const cycle = fixture();
    const first = cycle.block('xl("B1")');
    const second = cycle.block('xl("A1")');
    addCells(cycle.session, [
      { target: cycle.at(1, 1), content: { kind: 'python', blockId: first.id, source: first.source, block: first } },
      { target: cycle.at(1, 2), content: { kind: 'python', blockId: second.id, source: second.source, block: second } },
    ]);
    const firstId = findCell(cycle.session.active.sheets[0]!, cycle.at(1, 1))!.id;
    const secondId = findCell(cycle.session.active.sheets[0]!, cycle.at(1, 2))!.id;
    const calls: string[] = [];
    const cyclic = await calculateMixedGeneration({ basis: cycle.session.capture(),
      requestedPythonIds: new Set([firstId, secondId]), signal: new AbortController().signal },
    services(() => scalarOutput(1), calls));
    expect(calls).toEqual([]);
    expect(cyclic.failedPythonIds).toEqual(new Set([firstId, secondId]));

    const failed = fixture();
    const upstream = failed.block('raise');
    const prior: PythonBlock = { ...failed.block('xl("A1")'), hasResult: true, lastValue: { kind: 'number', value: 42 },
      status: 'succeeded', inputFingerprint: 'previous', runtimeFingerprint: 'pyodide-mixed-test' };
    addCells(failed.session, [
      { target: failed.at(1, 1), content: { kind: 'python', blockId: upstream.id, source: upstream.source, block: upstream } },
      { target: failed.at(1, 2), content: { kind: 'python', blockId: prior.id, source: prior.source, block: prior } },
    ]);
    const upstreamId = findCell(failed.session.active.sheets[0]!, failed.at(1, 1))!.id;
    const priorId = findCell(failed.session.active.sheets[0]!, failed.at(1, 2))!.id;
    const failureCalls: string[] = [];
    const generation = await calculateMixedGeneration({ basis: failed.session.capture(),
      requestedPythonIds: new Set([priorId]), forcedPythonIds: new Set([priorId]), signal: new AbortController().signal },
    services(source => source === 'raise' ? { status: 'failed', error: 'ValueError', stdout: '', stderr: '', outputTruncated: false }
      : scalarOutput(100), failureCalls));
    expect(failureCalls).toEqual(['raise']);
    expect(findCell(generation.document.sheets[0]!, failed.at(1, 1))?.content).toMatchObject({ kind: 'python', block: { status: 'failed' } });
    expect(findCell(generation.document.sheets[0]!, failed.at(1, 2))?.content).toMatchObject({ kind: 'python', block: {
      status: 'failed', hasResult: true, lastValue: { kind: 'number', value: 42 },
    } });
    expect(generation.messages.get(priorId)).toMatch(/upstream Python calculation failed/u);
  });

  it('retains timeout and worker exit status when a detached mixed run fails', async () => {
    const { session, at, block } = fixture();
    const python = block('xl("A1")');
    addCells(session, [{ target: at(1, 1), content: { kind: 'number', value: 1 } },
      { target: at(1, 2), content: { kind: 'python', blockId: python.id, source: python.source, block: python } }]);
    const id = findCell(session.active.sheets[0]!, at(1, 2))!.id;
    const calls: string[] = [];
    const generation = await calculateMixedGeneration({ basis: session.capture(), requestedPythonIds: new Set([id]),
      forcedPythonIds: new Set([id]), signal: new AbortController().signal }, services(() => {
      throw new PythonRuntimeError('Python timed out after 30 seconds.', 'timedOut');
    }, calls));
    expect(calls).toHaveLength(1);
    expect(findCell(generation.document.sheets[0]!, at(1, 2))?.content).toMatchObject({
      kind: 'python', block: { status: 'timedOut', error: 'Python timed out after 30 seconds.' },
    });
  });
});

