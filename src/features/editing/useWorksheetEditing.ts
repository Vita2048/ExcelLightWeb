import { useRef, useState } from 'react';
import { autoSum, captureClipboard, clear, clipboardText, entryText, fill, parseEntry, pasteBlock, pasteText,
  type Aggregate, type ClearMode, type ClipboardBlock, type EditRange, type FillDirection, type PasteMode,
} from '../../application/editing-commands';
import type { WorkbookSession } from '../../application/workbook-session';
import type { WorkbookSnapshot } from '../../application/workbook-protocol';
import { editingPreview } from '../../calculation/editing-preview';
import { equalDto } from '../../core/workbook-delta';
import { findReferencedCell, type CellDocument, type CellReference } from '../../core/workbook';
import { requireDocument } from '../../core/workbook-validation';

export type EditSurface = 'inline' | 'formula';
export type AfterEditKey = 'Enter' | 'Tab' | 'ArrowUp' | 'ArrowDown' | 'ArrowLeft' | 'ArrowRight';
export type EditAction =
  | { kind: 'copy' | 'cut' }
  | { kind: 'paste'; mode: PasteMode }
  | { kind: 'clear'; mode: ClearMode }
  | { kind: 'fill'; direction: FillDirection }
  | { kind: 'autosum'; aggregate: Aggregate };

interface Draft {
  readonly surface: EditSurface;
  readonly sessionId: string;
  readonly target: CellReference;
  readonly original: CellDocument | undefined;
  readonly text: string;
  readonly direct: boolean;
}

interface LocalClipboard {
  readonly token: string;
  readonly block: ClipboardBlock;
  readonly text: string;
}
interface ManualPaste {
  readonly basis: WorkbookSnapshot;
  readonly range: EditRange;
  readonly mode: PasteMode;
}

const clipboardType = 'application/x-excellight-cells';

export function useWorksheetEditing(
  session: WorkbookSession, publish: (snapshot: WorkbookSnapshot) => void, focusGrid: () => void,
  focusEditor: (surface: EditSurface) => void,
) {
  const liveDraft = useRef<Draft | null>(null);
  const composing = useRef(false);
  const [draft, renderDraft] = useState<Draft | null>(null);
  const [status, setStatus] = useState('');
  const localClipboard = useRef<LocalClipboard | null>(null);
  const [manualPaste, setManualPaste] = useState<ManualPaste | null>(null);
  const intent = useRef(0);
  const previewIds = useRef<ReadonlySet<string>>(new Set());

  function showDraft(next: Draft | null) {
    liveDraft.current = next;
    renderDraft(next);
  }

  function begin(target: CellReference, surface: EditSurface, replacement?: string) {
    const existing = liveDraft.current;
    if (existing) {
      if (surface !== existing.surface) showDraft({ ...existing, surface });
      return;
    }
    intent.current++;
    setStatus('');
    const original = findReferencedCell(session.active, target);
    showDraft({ surface, sessionId: session.sessionId, target, original, text: replacement ?? entryText(original),
      direct: surface === 'inline' && replacement !== undefined });
  }

  function update(text: string) {
    const current = liveDraft.current;
    if (current && current.text !== text) showDraft({ ...current, text });
  }

  function recalculate() {
    const updates = [...editingPreview(session.active)];
    const computed = new Set<string>();
    for (const update of updates) computed.add(findReferencedCell(session.active, update.target)!.id);
    for (const sheet of session.active.sheets) for (const cell of sheet.cells) {
      if (cell.content.kind === 'formula' && cell.cache && previewIds.current.has(cell.id) && !computed.has(cell.id)) {
        updates.push({ target: { sheetId: sheet.id, row: cell.row, column: cell.column } });
      }
    }
    if (updates.length > 0) {
      const { document: _document, kind: _kind, ...basis } = session.beginNativeCalculation('editing-preview-v1');
      session.tryCommitNativeCalculation({ ...basis, kind: 'native-calculated', updates });
    }
    previewIds.current = computed;
    publish(session.capture());
  }

  function reject(error: unknown) {
    setStatus(error instanceof Error ? error.message : 'Worksheet command failed.');
  }

  function commit(): boolean {
    const current = liveDraft.current;
    if (!current) return true;
    if (composing.current) return false;
    try {
      if (current.text !== entryText(current.original)) {
        const basis = session.capture();
        const actual = findReferencedCell(basis.document, current.target);
        requireDocument(current.sessionId === basis.sessionId && actual?.id === current.original?.id &&
          equalDto(actual?.content, current.original?.content) && equalDto(actual?.style, current.original?.style),
        'The cell changed while you were editing it.');
        const outcome = session.execute({ kind: 'set-cells', label: 'Edit cell', edits: [{
          target: current.target, content: parseEntry(current.text, actual?.style), resetCache: true,
        }] });
        if (outcome.changed) recalculate();
      }
      showDraft(null);
      setStatus('');
      return true;
    } catch (error) {
      reject(error);
      queueMicrotask(() => focusEditor(current.surface));
      return false;
    }
  }

  function cancel() {
    composing.current = false;
    showDraft(null);
    setStatus('');
    focusGrid();
  }

  function moveHistory(undo: boolean) {
    if (!commit()) return;
    intent.current++;
    if (undo ? session.undo() : session.redo()) recalculate();
    setStatus('');
    focusGrid();
  }

  function onlyRange(ranges: readonly EditRange[]): EditRange {
    requireDocument(ranges.length === 1, 'This command requires one contiguous selection.');
    return ranges[0]!;
  }

  function paste(basis: WorkbookSnapshot, range: EditRange, mode: PasteMode, rich: LocalClipboard | null, text?: string) {
    const operation = rich ? pasteBlock(basis, range, rich.block, mode) : pasteText(basis, range, text ?? '', mode);
    const outcome = session.execute(operation);
    if (rich?.block.cut && localClipboard.current === rich) localClipboard.current = null;
    if (outcome.changed) recalculate();
  }

  function closeManualPaste() {
    setManualPaste(null);
    setStatus('');
    requestAnimationFrame(focusGrid);
  }

  function submitManualPaste(text: string) {
    if (!manualPaste) return;
    try {
      requireDocument(session.isCurrent(manualPaste.basis), 'The worksheet changed before the paste completed.');
      paste(manualPaste.basis, manualPaste.range, manualPaste.mode, null, text);
      closeManualPaste();
    } catch (error) { reject(error); }
  }

  async function run(action: EditAction, ranges: readonly EditRange[], event?: DataTransfer): Promise<CellReference | undefined> {
    if (!commit()) return;
    const serial = ++intent.current;
    const basis = session.capture();
    setStatus('');
    try {
      if (action.kind === 'cut' || action.kind === 'copy') {
        const block = captureClipboard(basis, onlyRange(ranges), action.kind === 'cut');
        const text = clipboardText(block);
        const token = crypto.randomUUID();
        localClipboard.current = { text, block, token };
        if (block.cut) setStatus('Cut selection. Paste to move it.');
        if (event) {
          event.setData('text/plain', text);
          try { event.setData(clipboardType, token); } catch { /* Plain text is already available. */ }
        } else {
          focusGrid();
          try { await navigator.clipboard.writeText(text); }
          catch {
            if (serial === intent.current) setStatus(block.cut ? 'Cut selection. Use the ribbon to paste it.' :
              'Copied to this workbook. System clipboard access is unavailable.');
          }
        }
      } else if (action.kind === 'paste') {
        const range = onlyRange(ranges);
        const rich = localClipboard.current;
        if (event) {
          const text = event.getData('text/plain');
          const token = event.getData(clipboardType);
          // Browsers may strip custom formats. An exact plain-text match preserves the in-page block.
          if (rich && (token === rich.token || !rich.block.cut && !token &&
              text.replaceAll('\r\n', '\n') === rich.text.replaceAll('\r\n', '\n'))) {
            paste(basis, range, action.mode, rich);
          } else if (text || Array.from(event.types).includes('text/plain')) {
            paste(basis, range, action.mode, null, text);
          } else requireDocument(false, 'The clipboard contains no worksheet text.');
        } else {
          let text: string | undefined;
          let denied = false;
          // Base the completion on the original selection and revision.
          try {
            if (rich) {
              // Firefox and WebKit may show a browser paste prompt for programmatic reads.
              const permission = await navigator.permissions.query({ name: 'clipboard-read' as PermissionName });
              if (permission.state !== 'granted') throw new Error('System clipboard permission is unavailable.');
            }
            text = await navigator.clipboard.readText();
          }
          catch { denied = true; }
          if (serial !== intent.current || !session.isCurrent(basis)) return;
          if (rich && (denied || text?.replaceAll('\r\n', '\n') === rich.text.replaceAll('\r\n', '\n'))) {
            paste(basis, range, action.mode, rich);
          } else {
            if (denied || text === undefined) setManualPaste({ basis, range, mode: action.mode });
            else {
              requireDocument(text !== '', 'The system clipboard contains no worksheet text.');
              paste(basis, range, action.mode, null, text);
            }
          }
        }
      } else if (action.kind === 'clear') {
        if (session.execute(clear(basis, ranges, action.mode)).changed) recalculate();
      } else if (action.kind === 'fill') {
        if (session.execute(fill(basis, onlyRange(ranges), action.direction)).changed) recalculate();
      } else if (action.kind === 'autosum') {
        const result = autoSum(basis, onlyRange(ranges), action.aggregate);
        if (session.execute(result.command).changed) recalculate();
        return result.target;
      }
    } catch (error) { reject(error); }
    finally { if (serial === intent.current) focusGrid(); }
  }

  return {
    draft, status, begin, update, commit, cancel, moveHistory, run, manualPaste, closeManualPaste, submitManualPaste,
    recalculate, reject, clearStatus: () => setStatus(''),
    composing: () => composing.current,
    compose: (active: boolean) => { composing.current = active; },
    keepCaret: () => { if (liveDraft.current?.direct) showDraft({ ...liveDraft.current, direct: false }); },
    selectionChanged: () => { intent.current++; },
    clearCut: () => { if (localClipboard.current?.block.cut) localClipboard.current = null; },
  };
}
