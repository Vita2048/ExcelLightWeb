import {
  applyCellChanges, applyDelta, diffWorkbook, equalDto, estimateDtoBytes, invertDelta,
  sameAuthoredCell, sameAuthoredWorkbook, type CellChange, type WorkbookDelta,
} from '../core/workbook-delta';
import {
  cellKey, DocumentError, documentLimits, findReferencedCell, ownDto, type CellContent, type CellReference,
  type CellStyle, type DocumentId, type IdFactory, type WorkbookDocument,
} from '../core/workbook';
import {
  requireDocument, validateCellValue, validateId, validateTarget, validateWorkbook,
} from '../core/workbook-validation';
import {
  workbookProtocolVersion, type NativeCalculationRequest, type NativeCalculationReply,
  type SnapshotBasis, type WorkerBasis, type WorkbookSnapshot,
} from './workbook-protocol';

export interface CellEdit {
  readonly target: CellReference;
  readonly content: CellContent;
  readonly style?: CellStyle;
  /** Preserve by default. Set false to replace style or remove it by omitting style. */
  readonly preserveStyle?: boolean;
  readonly cellId?: DocumentId;
  readonly requiresExternalCalculation?: boolean;
}
export interface SetCellsCommand {
  readonly kind: 'set-cells';
  readonly edits: readonly CellEdit[];
  readonly label?: string;
}
export interface StructureCommand {
  readonly kind: 'structure';
  readonly document: WorkbookDocument;
  readonly label: string;
}
export type WorkbookCommand = SetCellsCommand | StructureCommand;
export interface CommandOutcome {
  readonly changed: boolean;
  readonly revision: number;
  readonly delta: WorkbookDelta;
}

interface HistoryBudget {
  readonly transactions: number;
  readonly bytes: number;
}
export const defaultHistoryBudget: HistoryBudget = Object.freeze({ transactions: 100, bytes: 64 * 1024 * 1024 });
export interface SessionOptions {
  readonly makeId: IdFactory;
  readonly initiallySaved?: boolean;
  readonly budget?: Partial<HistoryBudget>;
}

interface HistoryEntry {
  readonly forward: WorkbookDelta;
  readonly inverse: WorkbookDelta;
  readonly before: string;
  readonly after: string;
  readonly bytes: number;
}

function matchesBasis(left: SnapshotBasis, right: SnapshotBasis): boolean {
  return left.sessionId === right.sessionId && left.revision === right.revision &&
    left.generation === right.generation && left.sourceFingerprint === right.sourceFingerprint;
}

const emptyDelta: WorkbookDelta = ownDto({ kind: 'cells', changes: [] });

/** Owns the active workbook. All state transitions are synchronous between detached I/O tasks. */
export class WorkbookSession {
  readonly #makeId: IdFactory;
  readonly #budget: HistoryBudget;
  readonly #undo: HistoryEntry[] = [];
  readonly #redo: HistoryEntry[] = [];
  #issuedSnapshots = new WeakSet<WorkbookSnapshot>();
  #active: WorkbookSnapshot;
  #saved: WorkbookSnapshot | undefined;
  #dirty: boolean;
  #historyBytes = 0;
  #requestSerial = 0;
  #nativeRequest: WorkerBasis | undefined;

  constructor(document: WorkbookDocument, options: SessionOptions) {
    this.#makeId = options.makeId;
    this.#budget = { ...defaultHistoryBudget, ...options.budget };
    requireDocument(Number.isSafeInteger(this.#budget.transactions) && this.#budget.transactions > 0 &&
      this.#budget.transactions <= defaultHistoryBudget.transactions &&
      Number.isSafeInteger(this.#budget.bytes) && this.#budget.bytes > 0 && this.#budget.bytes <= defaultHistoryBudget.bytes,
    'Invalid history budget.');
    const owned = ownDto(document);
    validateWorkbook(owned);
    this.#active = this.#newSnapshot(owned, 0, 0);
    this.#saved = options.initiallySaved === false ? undefined : this.#active;
    this.#dirty = !this.#saved;
  }

  get active() { return this.#active.document; }
  get sessionId() { return this.#active.sessionId; }
  get revision() { return this.#active.revision; }
  get sourceFingerprint() { return this.#active.sourceFingerprint; }
  get savedIdentity() { return this.#saved?.sourceFingerprint; }
  get isDirty() { return this.#dirty; }
  get canUndo() { return this.#undo.length > 0; }
  get canRedo() { return this.#redo.length > 0; }
  get undoCount() { return this.#undo.length; }
  get historyBytes() { return this.#historyBytes; }

  capture(): WorkbookSnapshot {
    this.#issuedSnapshots.add(this.#active);
    return this.#active;
  }

  isCurrent(basis: SnapshotBasis): boolean {
    return matchesBasis(basis, this.#active);
  }

  adopt(document: WorkbookDocument, saved = true): void {
    const owned = ownDto(document);
    validateWorkbook(owned);
    const snapshot = this.#newSnapshot(owned, this.revision + 1, this.#active.generation + 1);
    this.#active = snapshot;
    this.#issuedSnapshots = new WeakSet();
    this.#saved = saved ? snapshot : undefined;
    this.#dirty = !saved;
    this.#nativeRequest = undefined;
    this.#undo.length = 0;
    this.#redo.length = 0;
    this.#historyBytes = 0;
  }

  tryAdopt(basis: SnapshotBasis, document: WorkbookDocument, saved = true): boolean {
    if (!this.isCurrent(basis)) return false;
    this.adopt(document, saved);
    return true;
  }

  markSaved(snapshot: WorkbookSnapshot): boolean {
    if (!this.#issuedSnapshots.has(snapshot) || snapshot.sessionId !== this.sessionId) return false;
    const dirty = !sameAuthoredWorkbook(snapshot.document, this.active);
    this.#saved = snapshot;
    this.#dirty = dirty;
    return true;
  }

  execute(command: WorkbookCommand): CommandOutcome {
    const label = command.kind === 'set-cells' ? command.label ?? 'Edit cells' : command.label;
    requireDocument(typeof label === 'string' && label.length <= 256, 'Invalid command label.');
    let candidate: WorkbookDocument;
    switch (command.kind) {
      case 'structure':
        candidate = ownDto(command.document);
        requireDocument(candidate?.id === this.active.id, 'Structure cannot replace workbook identity.');
        break;
      case 'set-cells':
        candidate = this.#editCells(command.edits);
        break;
      default:
        throw new DocumentError('Unsupported workbook command.');
    }
    if (candidate === this.active) return { changed: false, revision: this.revision, delta: emptyDelta };
    validateWorkbook(candidate);
    const delta = diffWorkbook(this.active, candidate);
    const authoredChanged = delta.kind === 'structure' ? !sameAuthoredWorkbook(this.active, candidate) :
      delta.changes.some(change => !change.before || !change.after || !sameAuthoredCell(change.before, change.after));
    if (!authoredChanged) return { changed: false, revision: this.revision, delta: emptyDelta };

    const after = this.#nextIdentity(candidate);
    const bytes = 256 + 2 * label.length + 2 * estimateDtoBytes(delta);
    requireDocument(bytes <= this.#budget.bytes, 'Edit exceeds the Undo budget. Split it into smaller edits.');
    const entry = { forward: delta, inverse: invertDelta(delta), before: this.sourceFingerprint, after, bytes };
    for (const old of this.#redo) this.#historyBytes -= old.bytes;
    this.#redo.length = 0;
    while (this.#undo.length >= this.#budget.transactions || this.#historyBytes + bytes > this.#budget.bytes) {
      this.#historyBytes -= this.#undo.shift()!.bytes;
    }
    this.#undo.push(entry);
    this.#historyBytes += bytes;
    this.#advance(candidate, after, after !== this.savedIdentity);
    return { changed: true, revision: this.revision, delta };
  }

  tryExecute(basis: SnapshotBasis, command: WorkbookCommand): boolean {
    if (!this.isCurrent(basis)) return false;
    this.execute(command);
    return true;
  }

  undo(): boolean { return this.#moveHistory(this.#undo, this.#redo, true); }
  redo(): boolean { return this.#moveHistory(this.#redo, this.#undo, false); }

  beginNativeCalculation(runtimeFingerprint: string): NativeCalculationRequest {
    requireDocument(typeof runtimeFingerprint === 'string' && runtimeFingerprint.length > 0 &&
      runtimeFingerprint.length <= 128, 'Invalid calculation runtime fingerprint.');
    const request = ownDto<NativeCalculationRequest>({
      ...this.#active, kind: 'calculate-native', protocolVersion: workbookProtocolVersion,
      requestId: String(++this.#requestSerial), runtimeFingerprint,
    });
    this.#nativeRequest = request;
    return request;
  }

  cancelNativeCalculation(): void {
    this.#nativeRequest = undefined;
  }

  tryCommitNativeCalculation(reply: NativeCalculationReply): boolean {
    const expected = this.#nativeRequest;
    if (!expected || reply?.kind !== 'native-calculated' || reply.protocolVersion !== expected.protocolVersion ||
        reply.requestId !== expected.requestId || reply.runtimeFingerprint !== expected.runtimeFingerprint ||
        !matchesBasis(reply, expected) || !this.isCurrent(expected)) return false;

    requireDocument(Array.isArray(reply.updates) && reply.updates.length <= documentLimits.formulas, 'Invalid calculation batch size.');
    const sheetIds = new Set(this.active.sheets.map(sheet => sheet.id));
    const seen = new Set<string>();
    const changes: CellChange[] = [];
    for (const update of reply.updates) {
      requireDocument(update?.target, 'Missing calculation target.');
      const { target, cache } = update;
      const key = this.#validateTarget(target, sheetIds, seen);
      seen.add(key);
      const cell = findReferencedCell(this.active, target);
      requireDocument(cell?.content.kind === 'formula', 'Calculation can update only existing native formula caches.');
      if (cache !== undefined) validateCellValue(cache, true);
      if (equalDto(cell.cache, cache)) continue;
      const { cache: _prior, ...rest } = cell;
      changes.push({ target, after: cache === undefined ? rest : { ...rest, cache } });
    }
    const candidate = applyCellChanges(this.active, changes);
    if (candidate !== this.active) this.#advance(candidate, this.sourceFingerprint, this.isDirty);
    this.#nativeRequest = undefined;
    return true;
  }

  #newSnapshot(document: WorkbookDocument, revision: number, generation: number): WorkbookSnapshot {
    const sessionId = this.#makeId();
    validateId(sessionId);
    requireDocument(sessionId !== this.#active?.sessionId, 'A new workbook session needs a fresh identity.');
    return ownDto({ document, sessionId, revision, generation, sourceFingerprint: `${sessionId}:${revision}` });
  }

  #nextIdentity(document: WorkbookDocument): string {
    if (this.#saved && sameAuthoredWorkbook(document, this.#saved.document)) return this.#saved.sourceFingerprint;
    return `${this.sessionId}:${this.revision + 1}`;
  }

  #isDirtyAt(document: WorkbookDocument, identity: string): boolean {
    return !this.#saved || identity !== this.#saved.sourceFingerprint &&
      !sameAuthoredWorkbook(document, this.#saved.document);
  }

  #advance(document: WorkbookDocument, sourceFingerprint: string, dirty: boolean): void {
    this.#active = ownDto({
      document, sessionId: this.sessionId, revision: this.revision + 1,
      generation: this.#active.generation + 1, sourceFingerprint,
    });
    this.#dirty = dirty;
    this.#nativeRequest = undefined;
  }

  #moveHistory(from: HistoryEntry[], to: HistoryEntry[], undo: boolean): boolean {
    const entry = from.at(-1);
    if (!entry) return false;
    requireDocument(this.sourceFingerprint === (undo ? entry.after : entry.before), 'History does not match the workbook.');
    const document = applyDelta(this.active, undo ? entry.inverse : entry.forward);
    validateWorkbook(document);
    const identity = undo ? entry.before : entry.after;
    const dirty = this.#isDirtyAt(document, identity);
    from.pop();
    to.push(entry);
    this.#advance(document, identity, dirty);
    return true;
  }

  #validateTarget(target: CellReference, sheetIds: ReadonlySet<string>, seen: ReadonlySet<string>): string {
    validateTarget(target, sheetIds);
    const key = `${target.sheetId}:${cellKey(target)}`;
    requireDocument(!seen.has(key), 'Duplicate target in workbook transaction.');
    return key;
  }

  #editCells(edits: readonly CellEdit[]): WorkbookDocument {
    requireDocument(Array.isArray(edits), 'Missing cell edits.');
    const sheetIds = new Set(this.active.sheets.map(sheet => sheet.id));
    const seen = new Set<string>();
    const changes: CellChange[] = [];
    for (const edit of edits) {
      requireDocument(edit?.target, 'Missing edit target.');
      const { target, content } = edit;
      seen.add(this.#validateTarget(target, sheetIds, seen));
      validateCellValue(content);
      if (edit.cellId !== undefined) validateId(edit.cellId);
      const prior = findReferencedCell(this.active, target);
      const style = edit.preserveStyle === false ? edit.style : prior?.style;
      const sameContent = equalDto(content, prior?.content);
      const external = edit.requiresExternalCalculation || sameContent && prior?.requiresExternalCalculation;
      requireDocument(!external || content.kind === 'formula', 'External calculation requires a formula.');
      const cache = sameContent || content.kind === 'formula' ? prior?.cache : undefined;
      if (content.kind === 'blank' && style === undefined) {
        if (prior) changes.push({ target });
        continue;
      }
      const after = {
        id: edit.cellId ?? prior?.id ?? this.#makeId(), row: target.row, column: target.column, content,
        ...(style !== undefined && { style }), ...(cache && { cache }),
        ...(external && { requiresExternalCalculation: true }),
      };
      if (!equalDto(prior, after)) changes.push({ target, after });
    }
    return applyCellChanges(this.active, changes);
  }
}
