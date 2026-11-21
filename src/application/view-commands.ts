import { documentLimits, type FrozenPanes, type SheetDocument, type WorkbookDocument } from '../core/workbook';
import { requireDocument } from '../core/workbook-validation';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { StructureCommand } from './workbook-session';

// Per-sheet frozen panes, matching DataCommands.SetFreeze. The grid projects the stored split;
// zoom and visibility are view state and never rewrite authored styles.
export const noFreeze: FrozenPanes = Object.freeze({ rows: 0, columns: 0 });

export function setFreeze(basis: WorkbookSnapshot, sheetId: string, panes: FrozenPanes): StructureCommand {
  requireDocument(Number.isSafeInteger(panes.rows) && Number.isSafeInteger(panes.columns) &&
    panes.rows >= 0 && panes.rows < documentLimits.rows && panes.columns >= 0 && panes.columns < documentLimits.columns,
  'Frozen pane boundaries are invalid.');
  const sheet = basis.document.sheets.find(item => item.id === sheetId);
  requireDocument(sheet, 'The worksheet no longer exists.');
  const unfrozen = panes.rows === 0 && panes.columns === 0;
  return {
    kind: 'structure',
    label: unfrozen ? 'Unfreeze' : 'Freeze',
    document: replaceSheet(basis.document, { ...sheet, frozenPanes: { rows: panes.rows, columns: panes.columns } }),
  };
}

function replaceSheet(document: WorkbookDocument, next: SheetDocument): WorkbookDocument {
  return { ...document, sheets: document.sheets.map(sheet => sheet.id === next.id ? next : sheet) };
}

/** The frozen split for a caret: rows above and columns to the left stay pinned. */
export function freezeFromCaret(point: { readonly row: number; readonly column: number }): FrozenPanes {
  return { rows: Math.max(0, point.row - 1), columns: Math.max(0, point.column - 1) };
}
