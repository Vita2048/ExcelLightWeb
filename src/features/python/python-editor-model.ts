import type { SnapshotBasis, WorkbookSnapshot } from '../../application/workbook-protocol';
import type { SetCellsCommand } from '../../application/workbook-session';
import { findCell, type CellDocument, type PythonBlock } from '../../core/workbook';

export interface PythonAnchorTarget {
  readonly sessionId: string;
  readonly sheetId: string;
  readonly cellId: string;
  readonly blockId: string;
  readonly row: number;
  readonly column: number;
  readonly source: string;
}

export interface PythonEditorMessage {
  readonly contextId: string;
  readonly text: string;
}


/** Resolve an anchor or an owned spill child to the Python anchor that owns the source. */
export function pythonAnchorAt(snapshot: WorkbookSnapshot, sheetId: string, point: { row: number; column: number }): PythonAnchorTarget | undefined {
  const sheet = snapshot.document.sheets.find(candidate => candidate.id === sheetId);
  if (!sheet) return undefined;
  let cell: CellDocument | undefined = findCell(sheet, point);
  if (cell?.spillOwner && cell.spillKind !== 'formula') {
    const owner = cell.spillOwner;
    cell = sheet.cells.find(candidate => candidate.content.kind === 'python' && candidate.content.blockId === owner);
  }
  if (cell?.content.kind !== 'python') return undefined;
  return {
    sessionId: snapshot.sessionId, sheetId, cellId: cell.id, blockId: cell.content.blockId,
    row: cell.row, column: cell.column, source: cell.content.source,
  };
}

export function samePythonTarget(left: PythonAnchorTarget | undefined, right: PythonAnchorTarget | undefined): boolean {
  return left === right || !!left && !!right && left.sessionId === right.sessionId &&
    left.sheetId === right.sheetId && left.cellId === right.cellId && left.blockId === right.blockId &&
    left.row === right.row && left.column === right.column;
}

export function isCurrentPythonTarget(basis: SnapshotBasis, target: PythonAnchorTarget): boolean {
  return basis.sessionId === target.sessionId && target.sessionId.length > 0;
}

/** The live source for a target, or undefined when the anchor moved, was cleared or belongs to another session. */
export function currentPythonSource(snapshot: WorkbookSnapshot, target: PythonAnchorTarget): string | undefined {
  if (snapshot.sessionId !== target.sessionId) return undefined;
  const sheet = snapshot.document.sheets.find(candidate => candidate.id === target.sheetId);
  const cell = sheet && findCell(sheet, { row: target.row, column: target.column });
  if (!cell || cell.id !== target.cellId || cell.content.kind !== 'python' || cell.content.blockId !== target.blockId) {
    return undefined;
  }
  return cell.content.source;
}

/**
 * Build one session command for a source draft. The block keeps its identity and legacy
 * bindings; only authored source and the previous run state change. This never executes Python.
 */
export function pythonSourceCommand(snapshot: WorkbookSnapshot, target: PythonAnchorTarget, text: string): SetCellsCommand | undefined {
  const source = currentPythonSource(snapshot, target);
  if (source === undefined || source === text) return undefined;
  const sheet = snapshot.document.sheets.find(candidate => candidate.id === target.sheetId)!;
  const cell = findCell(sheet, { row: target.row, column: target.column })!;
  const prior = cell.content.kind === 'python' ? cell.content : undefined;
  const existing: PythonBlock | undefined = prior?.block;
  const block: PythonBlock = existing ?
    { ...existing, source: text, status: 'never', error: null } :
    { id: target.blockId, source: text, bindings: [], lastValue: null, hasResult: false,
      inputFingerprint: null, runtimeFingerprint: null, status: 'never', error: null,
      lastTable: null, spillActive: false, lastFigure: null, placements: [] };
  const content = { kind: 'python' as const, blockId: block.id, source: text, block };
  return { kind: 'set-cells', label: 'Edit Python source', edits: [{ target: {
    sheetId: target.sheetId, row: target.row, column: target.column,
  }, content }] };
}

export function acceptsPythonEditorMessage(contextId: string | undefined, message: PythonEditorMessage): boolean {
  return !!contextId && message.contextId === contextId;
}
