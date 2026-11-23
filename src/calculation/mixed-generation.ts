import type { CacheUpdate, WorkbookSnapshot } from '../application/workbook-protocol';
import { applyPythonResult } from '../core/python-spills';
import { cellKey, findCell, type CellDocument, type PythonBlock, type WorkbookDocument } from '../core/workbook';
import { buildPythonInputs, type PythonInputAnchor } from '../python/python-inputs';
import { pythonBlockFromOutput } from '../python/python-results';
import { PythonRuntimeError } from '../python/python-runtime-client';
import type { PythonRuntimeOutput, PythonRuntimeInput } from '../python/runtime-protocol';
import { MixedGraph } from './mixed-graph';

export interface NativePassResult {
  readonly updates: readonly CacheUpdate[];
  readonly staleCellIds: readonly string[];
}

export interface MixedGenerationServices {
  readonly calculateNative: (document: WorkbookDocument, targets: readonly string[], signal: AbortSignal,
    onCalculating: (cellIds: readonly string[]) => void) => Promise<NativePassResult>;
  readonly runPython: (source: string, input: PythonRuntimeInput, signal: AbortSignal,
    onProgress: (message: string) => void) => Promise<{
    readonly runtimeFingerprint: string;
    readonly output: PythonRuntimeOutput;
  }>;
  readonly runtimeFingerprint: string;
  readonly makeId: () => string;
  readonly onOutput?: (pythonId: string, output: PythonRuntimeOutput) => void;
  readonly onNativeCalculating?: (pythonId: string, cellIds: readonly string[]) => void;
  readonly onPythonStart?: (pythonId: string) => void;
  readonly onPythonProgress?: (pythonId: string, message: string) => void;
}

export interface MixedGenerationRequest {
  readonly basis: WorkbookSnapshot;
  readonly requestedPythonIds: ReadonlySet<string>;
  readonly forcedPythonIds?: ReadonlySet<string>;
  readonly signal: AbortSignal;
}

export interface MixedGenerationResult {
  readonly document: WorkbookDocument;
  readonly staleFormulaIds: ReadonlySet<string>;
  readonly completedPythonIds: ReadonlySet<string>;
  readonly executedPythonIds: ReadonlySet<string>;
  readonly executedFigureIds: ReadonlyMap<string, string>;
  readonly failedPythonIds: ReadonlySet<string>;
  readonly messages: ReadonlyMap<string, string>;
}

function applyNativeUpdates(document: WorkbookDocument, updates: readonly CacheUpdate[]): WorkbookDocument {
  const bySheet = new Map<string, Map<number, CacheUpdate>>();
  for (const update of updates) {
    let values = bySheet.get(update.target.sheetId);
    if (!values) bySheet.set(update.target.sheetId, values = new Map());
    values.set(cellKey(update.target), update);
  }
  let changed = false;
  const sheets = document.sheets.map(sheet => {
    const values = bySheet.get(sheet.id);
    if (!values) return sheet;
    const cells = sheet.cells.map(cell => {
      const update = values.get(cellKey(cell));
      if (!update) return cell;
      values.delete(cellKey(cell));
      if (cell.content.kind !== 'formula') throw new Error('Native worker targeted a non-formula cell.');
      if (JSON.stringify(cell.cache) === JSON.stringify(update.cache)) return cell;
      changed = true;
      const { cache: _cache, ...rest } = cell;
      return update.cache === undefined ? rest : { ...rest, cache: update.cache };
    });
    if (values.size) throw new Error('Native worker returned a missing formula target.');
    return cells.every((cell, index) => cell === sheet.cells[index]) ? sheet : { ...sheet, cells };
  });
  if ([...bySheet.values()].some(values => values.size)) throw new Error('Native worker returned a missing worksheet.');
  return changed ? { ...document, sheets } : document;
}

function pythonAnchor(document: WorkbookDocument, id: string): PythonInputAnchor | undefined {
  for (const sheet of document.sheets) {
    const cell = sheet.cells.find(item => item.id === id);
    if (cell?.content.kind === 'python') return { sheetId: sheet.id, row: cell.row, column: cell.column,
      blockId: cell.content.blockId, source: cell.content.source };
  }
  return undefined;
}

function updatePythonBlock(document: WorkbookDocument, id: string,
  update: (block: PythonBlock) => PythonBlock,
  makeId: () => string): WorkbookDocument {
  const anchor = pythonAnchor(document, id);
  if (!anchor) return document;
  const cell = findCell(document.sheets.find(sheet => sheet.id === anchor.sheetId)!, anchor);
  if (cell?.content.kind !== 'python' || !cell.content.block) return document;
  const block = update(cell.content.block);
  return applyPythonResult(document, anchor, block, makeId);
}

function messageOf(error: unknown) { return error instanceof Error ? error.message : String(error); }
/** Runs one detached mixed generation and returns only its final document. */
export async function calculateMixedGeneration(request: MixedGenerationRequest,
  services: MixedGenerationServices): Promise<MixedGenerationResult> {
  const { signal, basis } = request;
  let document = basis.document;
  let graph = new MixedGraph(document);
  const roots = new Set([...request.requestedPythonIds].filter(id => graph.nodes.get(id)?.kind === 'python'));
  let selected = graph.closure(roots);
  const pending = new Set([...selected].filter(id => graph.nodes.get(id)?.kind === 'python'));
  const completed = new Set<string>();
  const executed = new Set<string>();
  const figureIds = new Map<string, string>();
  const failed = new Set<string>();
  const messages = new Map<string, string>();
  const staleFormulaIds = new Set<string>();
  let replans = 0;

  const fail = (id: string, message: string, status: 'failed' | 'timedOut' | 'workerExited' = 'failed') => {
    if (!graph.nodes.has(id)) return;
    try {
      document = updatePythonBlock(document, id, block => ({ ...block, status, error: message }), services.makeId);
    } catch { /* A failed prior result remains intact when its old spill cannot be reapplied. */ }
    failed.add(id);
    messages.set(id, message);
  };

  for (const id of graph.blocked) if (pending.has(id) && graph.nodes.get(id)?.kind === 'python') {
    fail(id, 'Circular mixed dependency or dependent of a cycle. Repair the references and run again.');
    pending.delete(id);
  }

  while (pending.size) {
    signal.throwIfAborted();
    graph = new MixedGraph(document);
    selected = graph.closure(roots);
    for (const id of selected) if (graph.nodes.get(id)?.kind === 'python' && !completed.has(id) && !failed.has(id)) pending.add(id);
    for (const id of [...pending]) if (!selected.has(id) || completed.has(id) || failed.has(id)) pending.delete(id);
    if (!pending.size) break;

    const ready = graph.order.find(node => node.kind === 'python' && pending.has(node.id));
    if (!ready) {
      for (const id of pending) fail(id, 'Circular mixed dependency or dependent of a cycle. Repair the references and run again.');
      pending.clear();
      break;
    }
    const id = ready.id;
    pending.delete(id);
    if (ready.error) { fail(id, ready.error); continue; }
    if ([...ready.dependencies].some(dependency => failed.has(dependency))) {
      fail(id, 'Skipped because an upstream Python calculation failed.');
      continue;
    }

    try {
      const nativeTargets = graph.nativePrerequisites(id);
      if (nativeTargets.size) {
        const native = await services.calculateNative(document, [...nativeTargets], signal,
          cellIds => services.onNativeCalculating?.(id, cellIds));
        document = applyNativeUpdates(document, native.updates);
        staleFormulaIds.clear();
        for (const stale of native.staleCellIds) staleFormulaIds.add(stale);
      }

      const anchor = pythonAnchor(document, id);
      if (!anchor) throw new Error('Python anchor changed; select it again.');
      const cell = findCell(document.sheets.find(sheet => sheet.id === anchor.sheetId)!, anchor);
      if (cell?.content.kind !== 'python' || !cell.content.block) throw new Error('Python anchor has no result block.');
      const block = cell.content.block;
      const snapshot = { ...basis, document };
      const built = buildPythonInputs(snapshot, anchor, staleFormulaIds);
      const force = request.forcedPythonIds?.has(id) ?? false;
      if (!force && block.status === 'succeeded' && block.hasResult &&
          block.inputFingerprint === built.fingerprint && block.runtimeFingerprint === services.runtimeFingerprint) {
        completed.add(id);
        continue;
      }

      services.onPythonStart?.(id);
      const result = await services.runPython(anchor.source, built.input, signal,
        message => services.onPythonProgress?.(id, message));
      executed.add(id);
      if (result.output.status === 'succeeded' && result.output.result?.kind === 'figure' &&
          result.output.result.figure.figureId) figureIds.set(id, result.output.result.figure.figureId.toUpperCase());
      services.onOutput?.(id, result.output);
      const next = pythonBlockFromOutput(block, result.output, built.fingerprint, result.runtimeFingerprint);
      document = applyPythonResult(document, anchor, next, services.makeId);
      if (next.status === 'succeeded') completed.add(id);
      else {
        failed.add(id);
        messages.set(id, next.error ?? 'Python execution failed.');
      }
      const nextGraph = new MixedGraph(document);
      if (!graph.topologyMatches(nextGraph)) {
        const changedEdges = [...nextGraph.nodes.values()].filter(node => {
          const prior = graph.nodes.get(node.id);
          return !prior || prior.dependencies.size !== node.dependencies.size ||
            [...prior.dependencies].some(dependency => !node.dependencies.has(dependency));
        }).map(node => node.id);
        if (++replans > Math.max(16, nextGraph.nodes.size * 2 + 8)) {
          throw new Error('Python output dependencies did not settle; use stable spill extents.');
        }
        for (const changedId of changedEdges) {
          if (changedId !== id) { completed.delete(changedId); failed.delete(changedId); pending.add(changedId); }
        }
        graph = nextGraph;
        selected = graph.closure(roots);
        for (const candidate of selected) if (graph.nodes.get(candidate)?.kind === 'python' &&
            !completed.has(candidate) && !failed.has(candidate)) pending.add(candidate);
      }
    } catch (error) {
      signal.throwIfAborted();
      fail(id, messageOf(error), error instanceof PythonRuntimeError &&
        (error.status === 'timedOut' || error.status === 'workerExited') ? error.status : 'failed');
    }
  }

  signal.throwIfAborted();
  graph = new MixedGraph(document);
  const downstream = graph.dependents([...roots, ...completed]);
  const nativeOutputs = new Set([...downstream].filter(id => graph.nodes.get(id)?.kind === 'formula'));
  if (nativeOutputs.size) {
    try {
      const native = await services.calculateNative(document, [...nativeOutputs], signal,
        cellIds => services.onNativeCalculating?.([...roots][0] ?? '', cellIds));
      document = applyNativeUpdates(document, native.updates);
      staleFormulaIds.clear();
      for (const stale of native.staleCellIds) staleFormulaIds.add(stale);
    } catch (error) {
      signal.throwIfAborted();
      throw new Error(`Native formula refresh failed: ${messageOf(error)}`);
    }
  }
  return { document, staleFormulaIds, completedPythonIds: completed, executedPythonIds: executed,
    executedFigureIds: figureIds, failedPythonIds: failed, messages };
}

