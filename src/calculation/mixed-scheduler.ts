import type { WorkbookSnapshot } from '../application/workbook-protocol';
import type { WorkbookSession } from '../application/workbook-session';
import type { WorkbookDocument } from '../core/workbook';
import { PythonRuntimeClient } from '../python/python-runtime-client';
import { pythonRuntimeFingerprint } from '../python/runtime-fingerprint';
import { MixedGraph } from './mixed-graph';
import { calculateMixedGeneration } from './mixed-generation';
import { NativePassClient } from './native-pass-client';
import type { PythonExecutionState } from '../features/python/usePythonExecution';

const singleEditDelay = 75;
const coalescingDelay = 220;

function pythonIds(document: WorkbookDocument) {
  return document.sheets.flatMap(sheet => sheet.cells.flatMap(cell =>
    cell.content.kind === 'python' ? [cell.id] : []));
}

function pythonStatus(document: WorkbookDocument, id: string) {
  for (const sheet of document.sheets) for (const cell of sheet.cells) {
    if (cell.id === id && cell.content.kind === 'python') return cell.content.block?.status ?? 'failed';
  }
  return 'failed';
}

function targetKey(document: WorkbookDocument, id?: string) {
  if (!id) return undefined;
  for (const sheet of document.sheets) {
    const cell = sheet.cells.find(item => item.id === id);
    if (cell) return sheet.id + ':' + cell.id;
  }
}

/** Coordinates debounce, cancellation and single-transaction mixed generations. */
export class MixedPythonScheduler {
  #runtime = new PythonRuntimeClient();
  #observed: WorkbookDocument;
  #graph: MixedGraph;
  #automatic = false;
  #timer: ReturnType<typeof setTimeout> | undefined;
  #queued = new Set<string>();
  #controller: AbortController | undefined;
  #nativeWorker: NativePassClient | undefined;
  #activeRoots: Set<string> | undefined;
  #activeAutomatic = false;
  #serial = 0;
  #displayId: string | undefined;
  #disposed = false;

  constructor(private readonly session: WorkbookSession, private readonly locale: string,
    private readonly publish: (state: PythonExecutionState | undefined) => void,
    private readonly onSnapshot: (snapshot: WorkbookSnapshot) => void,
    private readonly pauseNative: (paused: boolean) => void) {
    this.#observed = session.active;
    this.#graph = new MixedGraph(this.#observed);
  }

  changed(compound: boolean): boolean {
    if (this.#disposed) return false;
    const current = this.session.active;
    const before = this.#observed;
    const coalescing = this.#timer !== undefined || this.#controller !== undefined;
    const priorAutomaticRoots = this.#activeAutomatic ? this.#activeRoots : undefined;
    this.#observed = current;
    this.#graph = new MixedGraph(current);
    const affected = this.#graph.affectedPython(before, current);
    if (!affected.size) {
      if (this.#controller) {
        this.#cancel(false);
        if (priorAutomaticRoots?.size && this.#automatic) {
          this.#queued = new Set([...this.#queued, ...priorAutomaticRoots]);
          this.pauseNative(true);
          this.#schedule(true);
          return true;
        }
        this.pauseNative(false);
      }
      return false;
    }
    this.#cancel(false);
    if (!this.#automatic) {
      this.publish({ key: targetKey(current, [...affected][0]) ?? '', running: false, stdout: '', stderr: '',
        outputTruncated: false, message: 'Manual · Python results may be stale. Choose Run to refresh.' });
      this.pauseNative(false);
      return false;
    }
    this.#queued = new Set([...this.#queued, ...affected, ...(priorAutomaticRoots ?? [])]);
    this.pauseNative(true);
    this.#schedule(compound || coalescing);
    return true;
  }

  setAutomatic(automatic: boolean) {
    if (this.#automatic === automatic) return;
    this.#automatic = automatic;
    if (!automatic) {
      this.#cancel(true);
      this.#queued.clear();
      this.pauseNative(false);
      this.publish({ key: '', running: false, stdout: '', stderr: '', outputTruncated: false,
        message: 'Manual · Run Python to refresh results.' });
      return;
    }
    this.#observed = this.session.active;
    this.#graph = new MixedGraph(this.#observed);
    const all = pythonIds(this.#observed);
    if (all.length) {
      this.#queued = new Set(all);
      this.pauseNative(true);
      this.#schedule(true);
    }
  }

  adopted() {
    this.#cancel(true);
    this.#observed = this.session.active;
    this.#graph = new MixedGraph(this.#observed);
    if (this.#automatic) {
      this.#queued = new Set(pythonIds(this.#observed));
      if (this.#queued.size) { this.pauseNative(true); this.#schedule(true); }
    } else this.publish({ key: '', running: false, stdout: '', stderr: '', outputTruncated: false,
      message: 'Manual · opening a workbook never runs Python.' });
  }

  historyRestored() {
    this.#cancel(true);
    this.#queued.clear();
    this.#observed = this.session.active;
    this.#graph = new MixedGraph(this.#observed);
    this.pauseNative(false);
    this.publish({ key: '', running: false, stdout: '', stderr: '', outputTruncated: false,
      message: 'History restored; Python did not run.' });
  }

  run(pythonId: string, force = true) {
    this.#observed = this.session.active;
    this.#graph = new MixedGraph(this.#observed);
    if (this.#graph.nodes.get(pythonId)?.kind !== 'python') {
      this.publish({ key: '', running: false, stdout: '', stderr: '', outputTruncated: false,
        message: 'Python anchor is no longer current.' });
      return;
    }
    this.#queued.clear();
    this.#cancel(false);
    this.pauseNative(true);
    void this.#run(new Set([pythonId]), force ? new Set([pythonId]) : new Set(), false);
  }

  runAll() {
    const all = new Set(pythonIds(this.session.active));
    if (!all.size) return;
    this.#queued.clear();
    this.#cancel(false);
    this.pauseNative(true);
    void this.#run(all, all, false);
  }

  stop() {
    this.#cancel(true);
    this.#queued.clear();
    this.pauseNative(false);
    this.publish({ key: this.#displayId ? targetKey(this.session.active, this.#displayId) ?? '' : '', running: false,
      stdout: '', stderr: '', outputTruncated: false, message: 'Python run stopped.' });
  }

  dispose() {
    this.#disposed = true;
    clearTimeout(this.#timer);
    this.#controller?.abort();
    this.#runtime.dispose();
    this.#nativeWorker?.dispose();
  }

  #schedule(compound: boolean) {
    const coalescing = this.#timer !== undefined || this.#controller !== undefined;
    clearTimeout(this.#timer);
    const delay = compound || coalescing ? coalescingDelay : singleEditDelay;
    this.#timer = setTimeout(() => {
      this.#timer = undefined;
      const roots = new Set(this.#queued);
      this.#queued.clear();
      void this.#run(roots, new Set(), true);
    }, delay);
    this.publish({ key: targetKey(this.session.active, [...this.#queued][0]) ?? '', running: false,
      stdout: '', stderr: '', outputTruncated: false, message: 'Automatic · Python refresh queued.' });
  }

  #cancel(publishStopped: boolean) {
    clearTimeout(this.#timer);
    this.#timer = undefined;
    this.#serial++;
    this.#controller?.abort();
    this.#controller = undefined;
    this.#nativeWorker?.dispose();
    this.#nativeWorker = undefined;
    this.#activeRoots = undefined;
    this.#activeAutomatic = false;
    if (publishStopped && !this.#disposed) this.publish({ key: '', running: false, stdout: '', stderr: '',
      outputTruncated: false, message: 'Python stopped.' });
  }

  async #run(roots: ReadonlySet<string>, forced: ReadonlySet<string>, automatic: boolean) {
    if (!roots.size || this.#disposed) { this.pauseNative(false); return; }
    this.#cancel(false);
    const serial = ++this.#serial;
    const basis = this.session.capture();
    const controller = new AbortController();
    this.#controller = controller;
    this.#activeRoots = new Set(roots);
    this.#activeAutomatic = automatic;
    this.#displayId = [...roots][0];
    let stdout = '', stderr = '', outputTruncated = false;
    const isCurrent = () => serial === this.#serial && !this.#disposed;
    this.publish({ key: targetKey(basis.document, this.#displayId) ?? '', running: true, stdout, stderr,
      outputTruncated, message: automatic ? 'Automatic · calculating mixed dependencies…' : 'Starting local Python runtime…' });
    const native = new NativePassClient(basis, this.locale);
    this.#nativeWorker = native;
    try {
      const result = await calculateMixedGeneration({ basis, requestedPythonIds: roots, forcedPythonIds: forced,
        signal: controller.signal }, {
        calculateNative: async (document, targets, signal) => {
          const reply = await native.calculate(document, targets, signal);
          return { updates: reply.updates, staleCellIds: reply.staleCellIds };
        },
        runPython: async (source, input, signal) => this.#runtime.run(source, input, signal),
        runtimeFingerprint: pythonRuntimeFingerprint, makeId: () => crypto.randomUUID(),
        onOutput: (id, output) => {
          if (!isCurrent() || id !== this.#displayId) return;
          stdout = output.stdout;
          stderr = output.stderr;
          outputTruncated = output.outputTruncated;
          this.publish({ key: targetKey(basis.document, id) ?? '', running: true, stdout, stderr, outputTruncated,
            ...(output.error && { error: output.error }), message: 'Running Python locally…' });
        },
      });
      if (!isCurrent()) return;
      if (!this.session.tryCommitGeneration(basis, result.document,
        automatic ? 'Automatic Python calculation' : 'Run Python generation')) {
        this.#observed = this.session.active;
        this.#graph = new MixedGraph(this.#observed);
        this.publish({ key: '', running: false, stdout, stderr, outputTruncated,
          message: 'Workbook changed while Python was running. The generation was discarded.' });
        if (this.#automatic && automatic) {
          const changed = this.#graph.affectedPython(basis.document, this.session.active);
          this.#queued = new Set([...this.#queued, ...changed]);
          if (this.#queued.size) this.#schedule(true);
        }
        return;
      }
      const snapshot = this.session.capture();
      this.#observed = snapshot.document;
      this.#graph = new MixedGraph(snapshot.document);
      this.onSnapshot(snapshot);
      const errors = result.failedPythonIds.size;
      const firstError = result.messages.values().next().value;
      const failedStatuses = [...result.failedPythonIds].map(id => pythonStatus(result.document, id));
      this.publish({ key: targetKey(snapshot.document, this.#displayId) ?? '', running: false, stdout, stderr,
        outputTruncated, ...(firstError && { error: firstError }),
        message: errors ? automatic ? `Automatic · ${errors} Python cell${errors === 1 ? '' : 's'} failed or was skipped.` :
          failedStatuses.includes('timedOut') ? 'Python run timed out.' :
            failedStatuses.includes('workerExited') ? 'Python worker exited.' : 'Python run failed.' :
          automatic ? 'Automatic · Python results are current.' : 'Python run completed.' });
    } catch (error) {
      if (!isCurrent()) return;
      if (controller.signal.aborted) return;
      const message = error instanceof Error ? error.message : String(error);
      this.publish({ key: targetKey(basis.document, this.#displayId) ?? '', running: false, stdout, stderr,
        outputTruncated, error: message, message: automatic ? 'Automatic Python calculation failed.' : 'Python run failed.' });
    } finally {
      native.dispose();
      if (this.#nativeWorker === native) this.#nativeWorker = undefined;
      if (serial === this.#serial) {
        this.#controller = undefined;
        this.#activeRoots = undefined;
        this.#activeAutomatic = false;
        this.pauseNative(false);
      }
    }
  }
}

